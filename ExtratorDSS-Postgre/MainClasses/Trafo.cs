using System;
using System.Data.SqlClient;
using System.Text;
using Npgsql;

namespace ExportadorGeoPerdasDSS
{
    class Trafo
    {
        // membros privados
        private static readonly string _trafos = "Transformadores.dss";
        private static NpgsqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqTrafo;
        private readonly Param _par;
        private readonly ModeloSDEE _SDEE;

        public Trafo(NpgsqlConnectionStringBuilder connBuilder, Param par, ModeloSDEE SDEE)
        {
            _par = par;
            _connBuilder = connBuilder;
            _SDEE = SDEE;
        }

        //new transformer.TRF1151404 Phases=1,Windings=3,Buses=[162221270.3 BBT32992241.1.0 BBT32992241.0.2],Conns=[wye wye wye],kvs=[7.97 0.12 0.12],kvas=[15 15 15],Taps=[1.0 1 1],XscArray=[2.124,2.124,1.416],%loadloss=1.766666667 ,%noloadloss=0.433333333
        //new transformer.TRF1174742 Phases = 3, Windings = 2, Buses =[98410048.1.2.3 BBT36975913.1.2.3.0], Conns =[delta wye], kvs =[13.80 0.220], kvas =[75 75], Taps =[1.0 1.0], XHL = 3.72,%loadloss=1.466666667 ,%noloadloss=0.393333333
        //CodBase	CodTrafo	CodBnc	CodAlim	CodPonAcopl1	CodPonAcopl2	PotNom_kVA	MRT	TipTrafo	CodFasPrim	CodFasSecu	CodFasTerc	TenSecu_kV	Tap_pu	Resis_%	ReatHL_%	ReatHT_%	ReatLT_%	PerdTtl_%	PerdVz_%	ClssTrafo	Propr	Descr	CodSubAtrib	CodAlimAtrib	Ordm	De	Para	TnsLnh1_kV	TnsLnh2_kV
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqTrafo = new StringBuilder();

            using (NpgsqlConnection conn = new NpgsqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select Propr,CodTrafo,CodPonAcopl1,CodPonAcopl2,PotNom_kVA,TipTrafo,CodFasPrim,CodFasSecu,CodFasTerc,TenSecu_kV,"
                        + "Tap_pu,\"Resis_%\",\"ReatHL_%\",\"ReatHT_%\",\"ReatLT_%\",\"PerdVz_%\",TnsLnh1_kV,CodBnc,Descr"
                        + " from " + _par._DBschema + "StoredTrafoMTMTMTBT ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "where CodBase=@codbase and CodAlim=@CodAlim";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                        command.Parameters.AddWithValue("@CodAlim", _par._alim);
                    }

                    using (var rs = command.ExecuteReader())
                    {
                        // verifica ocorrencia de elemento no banco
                        if (!rs.HasRows)
                        {
                            return false;
                        }

                        while (rs.Read())
                        {
                            /* // TODO
                            //skipa trafos de terceiros
                            if (rs["Propr"].ToString().Equals("TC"))
                            {
                                // Console.Write("trafo de terceiros: " + rs["CodTrafo"].ToString() + Environment.NewLine );
                                continue;
                            }*/

                            string barraBT = rs["CodPonAcopl2"].ToString();
                            string pot = rs["PotNom_kVA"].ToString();

                            // double pot2 =Double.TryParse(rs["PotNom_kVA"];
                            // pot = string.Format("{0:.##}}", pot2);

                            string tensaoFF = rs["TnsLnh1_kV"].ToString();
                            if (tensaoFF.Equals(""))
                            {
                                tensaoFF = "13.8";
                            }
                            string tensaoFN = AuxFunc.GetTensaoFN(tensaoFF);
                            string tensaoFFsec = rs["TenSecu_kV"].ToString();

                            string fasePrim = AuxFunc.GetFasesDSS(rs["CodFasPrim"].ToString());

                            // verifies if transformer MTMT (e.g. 34.5/13.8kV or 13.8/34.5 kV) 
                            bool hasPT = IsMVMV_Transformer(tensaoFF, tensaoFFsec);

                            string CodTrafo = rs["CodTrafo"].ToString();

                            /* // DEBUG
                            if (CodTrafo.Equals("308866"))
                            {
                                int teste=0;
                            }*/

                            string Descr = rs["Descr"].ToString();

                            string linha;

                            // OBS: verificado que no geoperdas2018, o trafo mono apresenta TipTrafo=1
                            switch (rs["TipTrafo"].ToString())
                            {
                                //monofasico
                                case "1":
                                    linha = CriaStringTrafoMonofasico(rs, fasePrim, barraBT, tensaoFN, pot, Descr, tensaoFFsec);
                                    break;
                                //monofasico 3 fios
                                case "2":
                                    linha = CriaStringTrafoMonofasico_3fios(rs, fasePrim, barraBT, tensaoFN, pot, Descr, tensaoFFsec);
                                    break;
                                //BIFASICO 3 fios
                                case "3":
                                    linha = CriaStringTrafoBifasico_3fios(rs, fasePrim, barraBT, tensaoFF, pot, Descr, tensaoFFsec);
                                    break;
                                //posto transformador
                                case "5":
                                    linha = CriaStringPostoTransformador(rs, fasePrim, barraBT, tensaoFF, pot, Descr, tensaoFFsec);
                                    break;

                                default:
                                    linha = CriaStringTrafoTrifasico(rs, fasePrim, barraBT, tensaoFF, pot, Descr, tensaoFFsec);
                                    break;
                            }
                            _arqTrafo.Append(linha);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        // 
        private bool IsMVMV_Transformer(string tensaoPrim, string tensaoSecu)
        {
            // condicao retorn
            if (tensaoPrim.Equals("") || tensaoSecu.Equals(""))
            {
                return false;
            }

            double tensaoPrim_D = double.Parse(tensaoPrim);
            double tensaoSecu_D = double.Parse(tensaoSecu);

            // set flag in _Par
            if ((tensaoPrim_D == 13.8 && tensaoSecu_D == 34.5) ||
                  (tensaoPrim_D == 34.5 && tensaoSecu_D == 13.8))
            {
                return true;
            }
            return false;
        }

        //
        private string TensaoLinha2TensaoFase(string tensaoFF)
        {
            /*
            //verifica se eh vazia 
            if ( tensaoFF.Equals("") && tipoBanco.Equals("5") )
            {
                return "34.50";
            }
            if ( tensaoFF.Equals("") && tipoBanco.Equals("4") )
            {
                return "13.80";
            }*/

            //necessario transformar para double
            double tensaoFFd = double.Parse(tensaoFF);

            // Se tensao fasePrimDSS, o posto eh abaixador 
            if (tensaoFFd.Equals(34.5))
            {
                return "19.92";
            }
            if (tensaoFFd.Equals(13.8))
            {
                return "7.97";
            }
            return "7.97";
        }


        /* Exemplo de declaracao PT
! !elevador
new transformer.873922PT_1 Phases=1,Windings=2,Buses=[174122367.1.2 174122366.1.0],Conns=[delta wye],kvs=[13.8  19.92]
! !abaixador 
new transformer.873969PT_1 Phases=1,Windings=2,Buses=[174122333.1.2 174122332.1.0],Conns=[delta wye],kvs=[34.5 7.97]
        */
        private string CriaStringPostoTransformador(NpgsqlDataReader rs, string faseDSS, string barraBT, string tensaoPrim, string pot, string descr, string tensaoSecu)
        {
            // TODO 
            // string tensaoSecu

            //
            string faseSecuDSS = AuxFunc.GetFasesDSS(rs["CodFasSecu"].ToString());

            // OBS: de acordo com o padrao do GeoPerdas, o CodFasSec Como o PT eh delta/estrela, a tensao secundaria de linha deve ser transformada em tensao de fase.
            string tensaoSec = TensaoLinha2TensaoFase(rs["TenSecu_kV"].ToString());

            // 
            string linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "_" + rs["CodBnc"].ToString()
                 + " Phases=1"
                 + ",Windings=2"
                 + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS
                 + " " + barraBT + faseSecuDSS + "]"
                 + ",Conns=[delta wye]"
                 + ",kvs=[" + tensaoPrim + " " + tensaoSec + "]"
                 + ",kvas=[" + pot + " " + pot + "]"
                 + ",Taps=[1," + rs["Tap_pu"].ToString() + "]";

            // se modo reatancia
            if (_SDEE._reatanciaTrafos)
            {
                linha += ",XHL=" + rs["ReatHL_%"].ToString();
            }

            linha += ",%loadloss=" + rs["Resis_%"].ToString()
                 + ",%noloadloss=" + rs["PerdVz_%"].ToString()
                 + " !" + descr + Environment.NewLine;

            return linha;
        }

        // cria string trafo monofasico tap central
        private string CriaStringTrafoMonofasico_3fios(NpgsqlDataReader rs, string faseDSS, string barraBT, string tensaoFN, string pot, string descr, string tensaoFFsec)
        {
            string linha;
            if (!_par._modelo4condutores)
            {
                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "A"
                    + " Phases=1"
                    + ",Windings=3"
                    + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS //OBS1
                    + " " + barraBT + ".1.0 " + barraBT + ".0.2]"  //OBS1 + "BBT" //OBS: atenção para a polaridade .0.2
                    + ",Conns=[wye wye wye]"
                    + ",kvs=[" + tensaoFN + " " + tensaoFFsec + " " + tensaoFFsec + "]"
                    + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                    + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString()
                    + ",XHT=" + rs["ReatHT_%"].ToString()
                    + ",XLT=" + rs["ReatLT_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                    + ",%noloadloss=" + rs["PerdVz_%"].ToString()
                    + " !" + descr + Environment.NewLine;
            }
            else
            {
                /*
 New "Transformer.TRF_1042516A" phases=1 windings=3 buses=["119948602.2" "107459020.1.4" "107459020.4.2"] conns=[Wye Wye Wye] kvs=[7.96743371481684 0.12 0.12] taps=[1 1 1] kvas=[10 10 10] %loadloss=1.8 %noloadloss=0.45
 New "Reactor.TRF_1042516A_R" phases=1 bus1=107459020.4 R=15 X=0 basefreq=60 */

                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "A"
                + " Phases=1"
                + ",Windings=3"
                + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS //
                + " " + barraBT + ".1.4 " + barraBT + ".4.2]"  //+ "BBT" //+ "BBT" : atenção para a polaridade .0.2
                + ",Conns=[wye wye wye]"
                + ",kvs=[" + tensaoFN + " " + tensaoFFsec + " " + tensaoFFsec + "]"
                + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString()
                    + ",XHT=" + rs["ReatHT_%"].ToString()
                    + ",XLT=" + rs["ReatLT_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                    + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

                linha += "New Reactor." + rs["CodTrafo"].ToString() + "R" + " phases=1,bus1=" + barraBT + ".4,R=15,X=0,basefreq=60" + Environment.NewLine; //+"BBT"
            }

            return linha;
        }

        // cria string trafo monofasico tap central
        private string CriaStringTrafoMonofasico(NpgsqlDataReader rs, string faseDSS, string barraBT, string tensaoFN, string pot, string descr, string tensaoFFsec)
        {
            string linha;
            if (!_par._modelo4condutores)
            {
                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "A"
                    + " Phases=1"
                    + ",Windings=2"
                    + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS //OBS1
                    + " " + barraBT + ".1.0]"  //OBS1 + "BBT" //OBS: atenção para a polaridade .0.2
                    + ",Conns=[wye wye]"
                    + ",kvs=[" + tensaoFN + " " + tensaoFFsec + "]"
                    + ",kvas=[" + pot + " " + pot + "]"
                    + ",Taps=[1," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                    + ",%noloadloss=" + rs["PerdVz_%"].ToString()
                    + " !" + descr + Environment.NewLine;
            }
            else
            {
                /*
 New "Transformer.TRF_1042516A" phases=1 windings=3 buses=["119948602.2" "107459020.1.4" "107459020.4.2"] conns=[Wye Wye Wye] kvs=[7.96743371481684 0.12 0.12] taps=[1 1 1] kvas=[10 10 10] %loadloss=1.8 %noloadloss=0.45
 New "Reactor.TRF_1042516A_R" phases=1 bus1=107459020.4 R=15 X=0 basefreq=60 */

                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "A"
                + " Phases=1"
                + ",Windings=3"
                + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS //
                + " " + barraBT + ".1.4 " + barraBT + ".4.2]"  //+ "BBT" //+ "BBT" : atenção para a polaridade .0.2
                + ",Conns=[wye wye wye]"
                + ",kvs=[" + tensaoFN + " " + tensaoFFsec + " " + tensaoFFsec + "]"
                + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString()
                    + ",XHT=" + rs["ReatHT_%"].ToString()
                    + ",XLT=" + rs["ReatLT_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                    + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

                linha += "New Reactor." + rs["CodTrafo"].ToString() + "R" + " phases=1,bus1=" + barraBT + ".4,R=15,X=0,basefreq=60" + Environment.NewLine; //+"BBT"
            }

            return linha;
        }

        // cria string trafo 
        // Primary Delta - Seocndary Y
        private string CriaStringTrafoBifasico_3fios(NpgsqlDataReader rs, string faseDSS, string barraBT, string tensaoFF, string pot, string descr, string tensaoFFsec)
        {
            string linha;
            if (!_par._modelo4condutores)
            {
                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "A"
                    + " Phases=2"
                    + ",Windings=3"
                    + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS
                    + " " + barraBT + ".1.0 " + barraBT + ".0.2]" //OBS: atenção para a polaridade .0.2
                    + ",Conns=[delta wye wye]"
                    + ",kvs=[" + tensaoFF + " " + tensaoFFsec + " " + tensaoFFsec + "]"
                    + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                    + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString()
                    + ",XHT=" + rs["ReatHT_%"].ToString()
                    + ",XLT=" + rs["ReatLT_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                    + ",%noloadloss=" + rs["PerdVz_%"].ToString()
                    + " !" + descr + Environment.NewLine;
            }
            else
            {
                /*
 New "Transformer.TRF_1042516A" phases=1 windings=3 buses=["119948602.2" "107459020.1.4" "107459020.4.2"] conns=[Wye Wye Wye] kvs=[7.96743371481684 0.12 0.12] taps=[1 1 1] kvas=[10 10 10] %loadloss=1.8 %noloadloss=0.45
 New "Reactor.TRF_1042516A_R" phases=1 bus1=107459020.4 R=15 X=0 basefreq=60 */

                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString() + "A"
                + " Phases=2"
                + ",Windings=3"
                + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS
                + " " + barraBT + ".1.4 " + barraBT + ".4.2]" // atenção para a polaridade .0.2
                + ",Conns=[delta wye wye]"
                + ",kvs=[" + tensaoFF + " " + tensaoFFsec + " " + tensaoFFsec + "]"
                + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString()
                    + ",XHT=" + rs["ReatHT_%"].ToString()
                    + ",XLT=" + rs["ReatLT_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                    + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

                linha += "New Reactor." + rs["CodTrafo"].ToString() + "R" + " phases=1,bus1=" + barraBT + ".4,R=15,X=0,basefreq=60" + Environment.NewLine; //+"BBT"
            }
            return linha;
        }

        // cria string trafo trifasico
        private string CriaStringTrafoTrifasico(NpgsqlDataReader rs, string faseDSS, string barraBT, string tensaoFF, string pot, string descr, string tensaoFFsec)
        {
            string consSec = "wye";

            // TODO OBS: comentado pois alterava a tensao de linha correta do PT 34,5 kV 
            //tensaoFF = TrataTensaoLinhaANEEL(tensaoFF,"4");
            string linha;

            if (!_par._modelo4condutores)
            {
                //
                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString()
                     + " Phases=3"
                     + ",Windings=2"
                     + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS
                     + " " + barraBT + ".1.2.3.0]"
                     + ",Conns=[delta," + consSec + "]"
                     + ",kvs=[" + tensaoFF + "," + tensaoFFsec + "]"
                     + ",kvas=[" + pot + " " + pot + "]"
                     + ",Taps=[1," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                     + ",%noloadloss=" + rs["PerdVz_%"].ToString()
                     + " !" + descr + Environment.NewLine;
            }
            else
            {
                //
                linha = "new transformer.TRF_" + rs["CodTrafo"].ToString()
                     + " Phases=3"
                     + ",Windings=2"
                     + ",Buses=[" + rs["CodPonAcopl1"].ToString() + faseDSS
                     + " " + barraBT + ".1.2.3.4]"
                     + ",Conns=[delta," + consSec + "]"
                     + ",kvs=[" + tensaoFF + "," + tensaoFFsec + "]"
                     + ",kvas=[" + pot + " " + pot + "]"
                     + ",Taps=[1," + rs["Tap_pu"].ToString() + "]";

                // se modo reatancia
                if (_SDEE._reatanciaTrafos)
                {
                    linha += ",XHL=" + rs["ReatHL_%"].ToString();
                }

                linha += ",%loadloss=" + rs["Resis_%"].ToString()
                     + ",%noloadloss=" + rs["PerdVz_%"].ToString()
                     + " !" + descr + Environment.NewLine;

                linha += "New Reactor." + rs["CodTrafo"].ToString() + "R" + " phases=1,bus1=" + barraBT + ".4,R=15,X=0,basefreq=60" + Environment.NewLine; //+"BBT"
            }

            return linha;
        }

        private string GetNomeArq()
        {
            return _par._pathAlim + _par._alim + _trafos;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqTrafo.ToString(), GetNomeArq());
        }
    }
}
