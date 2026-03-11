using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;
using Npgsql;

namespace ExportadorGeoPerdasDSS
{
    class CargaBT
    {
        // membros privados
        private static readonly string _cargaBT = "CargaBT_";
        private static NpgsqlConnectionStringBuilder _connBuilder;
        private int _iMes;
        private string _ano;
        private readonly Param _par;
        private StringBuilder _arqSegmentoBT;
        private readonly List<List<int>> _numDiasFeriadoXMes;
        private readonly Dictionary<string, double> _somaCurvaCargaDiariaPU;
        private readonly ModeloSDEE _SDEE;

        public CargaBT(NpgsqlConnectionStringBuilder connBuilder, int iMes, string ano,
            List<List<int>> numDiasFeriadoXMes, Dictionary<string, double> somaCurvaCargaDiariaPU, ModeloSDEE sdee,
            Param par)
        {
            _par = par;
            _connBuilder = connBuilder;
            _iMes = iMes;
            _ano = ano;
            _numDiasFeriadoXMes = numDiasFeriadoXMes;
            _somaCurvaCargaDiariaPU = somaCurvaCargaDiariaPU;
            _SDEE = sdee;
        }

        //modelo
        //new load.3001215463M1 bus1=R9772.1.3.0,Phases=2,kv=0.22,kw=1.29794758726823,pf=0.92,Vminpu=0.92,Vmaxpu=1.5,model=2,daily=arqCurvaNormRES4_11,status=variable
        //new load.3001215463M2 bus1=R9772.1.3.0,Phases=2,kv=0.22,kw=1.29794758726823,pf=0.92,Vminpu=0.92,Vmaxpu=1.5,model=3,daily=arqCurvaNormRES4_11,status=variable
        // CodBase	CodConsBT	CodAlim	CodTrafo	CodRmlBT	CodFas	CodPonAcopl	SemRedAssoc	TipMedi	TipCrvaCarga	EnerMedid01_MWh	EnerMedid02_MWh	EnerMedid03_MWh	EnerMedid04_MWh	EnerMedid05_MWh	EnerMedid06_MWh	EnerMedid07_MWh	EnerMedid08_MWh	EnerMedid09_MWh	EnerMedid10_MWh	EnerMedid11_MWh	EnerMedid12_MWh	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	TnsLnh_kV	TnsFas_kV
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqSegmentoBT = new StringBuilder();

            using (NpgsqlConnection conn = new NpgsqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    string sql = "select TipTrafo,TenSecu_kV,CodConsBT,CodFas,CodPonAcopl,TipCrvaCarga,tnslnh_kv,tnsfas_kv,EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh,EnerMedid07_MWh," +
                            "EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh from " +
                            _par._DBschema + "StoredCargaBT as car inner join " + _par._DBschema + "StoredTrafoMTMTMTBT as tr on tr.CodTrafo = car.CodTrafo ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText = sql + "where car.CodBase=@codbase and tr.CodBase=@codbase and car.CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = sql + "where car.CodBase=@codbase and tr.CodBase=@codbase and car.CodAlim=@CodAlim";
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
                            CriaDSSCargaBT_Geral(rs);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        private void CriaDSSCargaBT_Geral(NpgsqlDataReader rs)
        {
            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString(), _par._modelo4condutores);
            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());

            // obtem tensao base de acordo com tipo da carga (mono, bi ou tri)
            string Kv = GetTensaoBase(numFases, rs);

            //obtem o consumo de acordo com o mes 
            string consumoMes = AuxFunc.GetConsumoMesCorrente(rs, _iMes);

            // se consumo nao eh vazio, transforma para double
            // OBS: optou-se por esta funcao visto que o banco pode retornar: "0","0.00000" e etc...
            if (!consumoMes.Equals(""))
            {
                double dConsumoMes = double.Parse(consumoMes);

                // skipa consumo = 0
                if (dConsumoMes == 0)
                {
                    return;
                }

                string demanda = AuxFunc.CalcDemanda(consumoMes, _iMes, _ano, rs["TipCrvaCarga"].ToString(), _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU);

                string linha = "";

                // se modelo de carga ANEEL
                switch (_SDEE._modeloCarga)
                {
                    case "ANEEL":

                        linha = CriaDSSCargaBTAneel(rs, demanda, fases, numFases, Kv);

                        break;

                    // modelo P constante
                    case "PCONST":

                        linha = CriaDSSCargaPconstBT(rs, demanda, fases, numFases, Kv);

                        break;
                }

                _arqSegmentoBT.Append(linha);

            }
            /*
            else
            {
                continue;
            }*/

            /* //OBS: DEBUG exclui IP
            if (rs["TipCrvaCarga"].ToString().Equals("IP"))
            {
                continue;
            }*/

        }

        private string CriaDSSCargaPconstBT(NpgsqlDataReader rs, string demanda, string fases, string numFases, string Kv)
        {
            string linha;

            if (!_par._modelo4condutores)
            {
                linha = "new load.BT_" + rs["CodConsBT"].ToString() + "_M1"
                   + " bus1=" + rs["CodPonAcopl"] + fases
                   + ",Phases=" + numFases
                   + ",kv=" + Kv
                   + ",kW=" + demanda
                   + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                   + ",model=1"
                   + ",daily=" + rs["TipCrvaCarga"].ToString()
                   + ",status=variable";
            }
            else
            {
                /* // OBS: DEBUG
                // nao grava demanda abaixo de 0.001 KWh
                if (demandaD < 0.001)
                {
                    continue;
                }
                */

                string tipLig = "wye";
                if (numFases == "1")
                {
                    tipLig = "wye";
                }
                /* //maneira como ANEEL simula
                if (numFases == "2")
                {
                    numFases = "1"; 
                    tipLig = "delta"; 
                }*/
                // maneira correta de se simular cargas BI
                if (numFases == "2")
                {
                    numFases = "2";
                    tipLig = "wye";
                }
                if (numFases == "3")
                {
                    tipLig = "delta";
                }

                linha = "new load.BT_" + rs["CodConsBT"].ToString() + "_M1"
                   + " bus1=" + rs["CodPonAcopl"] + fases //" + prefixoBarraBT
                   + ",Phases=" + numFases + ",Conn=" + tipLig
                   + ",kv=" + Kv
                   + ",kW=" + demanda
                   + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                   + ",model=1"
                   + ",daily=" + rs["TipCrvaCarga"].ToString()
                   + ",status=variable";
            }

            // alterar numCust=0 p/ cargas do tipo IP (iluminacao publica)
            if (rs["TipCrvaCarga"].ToString().Equals("IP"))
            {
                linha += ",NumCust=0" + Environment.NewLine;
            }
            else
            {
                linha += Environment.NewLine;
            }

            return linha;
        }

        private string CriaDSSCargaBTAneel(NpgsqlDataReader rs, string demanda, string fases, string numFases, string Kv)
        {
            // divide demanda entre 2 cargas
            double demandaD = double.Parse(demanda) / 2;

            string linha;

            // carga model=2
            linha = "new load.BT_" + rs["CodConsBT"].ToString() + "_M2"
                + " bus1=" + rs["CodPonAcopl"] + fases //OBS1
                + ",Phases=" + numFases
                + ",kv=" + Kv
                + ",kW=" + demandaD.ToString()
                + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                + ",model=2"
                + ",daily=" + rs["TipCrvaCarga"].ToString()
                + ",status=variable" + Environment.NewLine;

            // carga model=3
            linha += "new load.BT_" + rs["CodConsBT"].ToString() + "_M3"
                + " bus1=" + rs["CodPonAcopl"] + fases //OBS1
                + ",Phases=" + numFases
                + ",kv=" + Kv
                + ",kW=" + demandaD.ToString()
                + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                + ",model=3"
                + ",daily=" + rs["TipCrvaCarga"].ToString()
                + ",status=variable";

            // alterar numCust=0 p/ cargas do tipo IP (iluminacao publica)
            if (rs["TipCrvaCarga"].ToString().Equals("IP"))
            {
                linha += ",NumCust=0" + Environment.NewLine;
            }
            else
            {
                linha += Environment.NewLine;
            }

            return linha;
        }

        // A tensao base depende do numero de fases e tambem do trafo.
        private string GetTensaoBase(string numFases, NpgsqlDataReader rs)
        {
            //string tipoTrafo = rs["TipTrafo"].ToString();
            string tnslnh_kv = rs["tnslnh_kv"].ToString();
            string tnsfas_kv = rs["tnsfas_kv"].ToString();

            // se monofasico, usa tensao fas
            if ( numFases.Equals("1") )
            {
                //
                if (tnsfas_kv.Equals(""))
                    return "0.127"; // TODO 
             
                else
                    return tnsfas_kv;
            }
            //
            if (tnslnh_kv.Equals(""))
                return "0.220"; // TODO 

            else
                return tnsfas_kv;
           }

        public string GetNomeArq()
        {
            string strMes = AuxFunc.IntMes2strMes(_iMes);

            return _par._pathAlim + _par._alim + _cargaBT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
