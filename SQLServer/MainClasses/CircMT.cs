using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExportadorGeoPerdasDSS
{
    class CircMT
    {
        private Param _par;
        private static SqlConnectionStringBuilder _connBuilder;
        private readonly bool _modoReconf;
        private readonly bool _criaArqCoordenadas;
        private StrBoolElementosSDE _structElem;
        private int _iMes;
        private string _coordMT;

        private string _Alim_PAC = "";
        private string _TenNom_kV = "";
        private string _TenOpe_pu = "";
        private string _linhaEM = "";

        // strings de saida
        public string _stringMasterDSS_A; 
        public string _stringMasterDSS_B;
        public string _strCab;

        public CircMT(SqlConnectionStringBuilder connBuilder, Param par, bool modoReconfiguracao,
            bool criaArqCoordenadas, StrBoolElementosSDE structElem, int iMes, string coordMT)
        {
            _par = par;
            _connBuilder = connBuilder;
            _modoReconf = modoReconfiguracao;
            _criaArqCoordenadas = criaArqCoordenadas;
            _structElem = structElem;
            _iMes = iMes;
            _coordMT = coordMT;
        }

        private string GetNomeArqCoord()
        {
            return _par._alim + _coordMT;
        }

        /* // estrutura tipica do arquivo masterDSS criado
        clear

        new circuit.alimAFNU16 bus1=BMT155673172,basekv=13.8,pu=1.036

        ! Arquivo de curvas de carga normalizadas
        Redirect "..\PermRes\NovasCurvasTxt\CurvasDeCargaDU.dss"

        ! Arquivo de condutores
        Redirect "..\PermRes\Condutores.dss"

        Redirect AFNU16Transformadores.dss
        Redirect AFNU16Ramais.dss
        Redirect AFNU16ChavesMT.dss
        Redirect AFNU16SegmentosBT.dss
        Redirect AFNU16SegmentosMT.dss
        Redirect AFNU16Carga_BT.dss
        Redirect AFNU16Carga_MT.dss

        new energymeter.carga element=line.TR12301107,terminal=1

        ! Seta as tensoes de base do sistema
        Set voltagebases="13.8 0.24 0.22"

        ! Calcula as tensoes das linhas em pu
        CalcVoltageBases

        !Set mode=daily  hour=0 number=24 stepsize=1h

        Solve mode=daily  hour=0 number=24 stepsize=1h
        */
        public void ConsultaStoredCircMT()
        {
            // Obtem PAC inicio alim, tensoes base, operacao e trecho p/ alocar Energymeter
            Get_PAC_CTMT();

            // cria String MasterDSS Parte A
            _stringMasterDSS_A = CriaStringMasterDSS_ParteA(); 

            // cria String MasterDSS Parte B
            _stringMasterDSS_B = CriaStringMasterDSS_ParteB();

            // OBS: As partes A e B são adicionadas e posteriormente gravadas e um único arquivo 
            // com a seguinte nomenclaura XXXAnualA.dss, onde XXX eh o nome do alimentador
            // OBS: A parte B eh gravada com o nome XXXAnualB eh utilizada pela customização "ExecutorOpenDSS"
            _stringMasterDSS_A += _stringMasterDSS_B;

            // cria string arquivo MasterDSS que sera utilizada pela customizacao "ExecutorOpenDSS"
            // com a seguinte nomenclaura XXX.dss, onde XXX eh o nome do alimentador
            _strCab = CriaStrCabecalhoCustomizacao();
        }

        // Obtem PAC inicio alim, tensoes base, operacao e trecho p/ alocar Energymeter
        private void Get_PAC_CTMT()
        {
            string trechEM;

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    //Verifica se há mais de 1 trecho saindo do PelPrincipal
                    command.CommandText = "select [t].CodBase,[t].CodAlim,[t].CodPonAcopl,[t].TenNom_kV,[t].TenOpe_pu,[t].CodSegmMT," +
                        "[t].CodPonAcopl1,[t].CodPonAcopl2,count([t].CodSegmMT) over() as 'rowCount' " +
                        "from ( select [ct].CodBase,[ct].CodAlim,CodPonAcopl,TenNom_kV,TenOpe_pu,CodSegmMT,CodPonAcopl1,CodPonAcopl2 from " +
                        _par._DBschema + "StoredCircmt as [ct] " +
                        "inner join " + _par._DBschema + "StoredSegmentoMT as [seg] on [seg].CodPonAcopl1 = [ct].CodPonAcopl " +
                        "UNION select [ct].CodBase,[ct].CodAlim,CodPonAcopl,TenNom_kV,TenOpe_pu,CodSegmMT,CodPonAcopl1,CodPonAcopl2 from " +
                        _par._DBschema + "StoredCircmt as [ct] " +
                        "inner join " + _par._DBschema + "StoredSegmentoMT as [seg] on [seg].CodPonAcopl2 = [ct].CodPonAcopl ) as [t] ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "where [t].CodBase=@codbase and [t].CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "where [t].CodBase=@codbase and [t].CodAlim=@CodAlim";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                        command.Parameters.AddWithValue("@CodAlim", _par._alim);
                    }

                    using (var rs = command.ExecuteReader())
                    {
                        rs.Read();

                        // verifica se NAO tem linhas
                        if (!rs.HasRows)
                        {
                            Console.Write("1o PAC " + _par._alim + " não associado com SegmentoMT." + Environment.NewLine);
                        }
                        else
                        {
                            // TODO REVISAR !!
                            // exibe msg de erro.
                            /*
                            if (!_modoReconf)
                            {

                            }
                            if (_modoReconf)
                            {
                                _Alim_PAC = "CABFIC";
                            }*/

                            string pacCTME = rs["CodPonAcopl"].ToString();
                            string pac1 = rs["CodPonAcopl1"].ToString();
                            string pac2 = rs["CodPonAcopl2"].ToString();
                            string terminal="terminal=1";

                            // PAC ALIM
                            _Alim_PAC = "BMT" + pacCTME; //OBS: necessario adicionar prefixo BMT TODO
                            _TenNom_kV = rs["TenNom_kV"].ToString();
                            _TenOpe_pu = rs["TenOpe_pu"].ToString();
                            
                            // preenche trecho Energymeter
                            // verifica N de linhas
                            if (rs["rowCount"].ToString().Equals("1"))
                            {
                                // se pac CTMT != pac1 deve se alocar no terminal 1 
                                if (!pacCTME.Equals(pac1))
                                {
                                    terminal = "terminal=2";
                                }

                                trechEM = rs["CodSegmMT"].ToString();
                            }
                            // se ha mais de 1 trecho p/ o PAC de inicio do alim, cria trecho ficticio a montante
                            else
                            {
                                // trecho ficticio p/ correcao OU whole Substation 
                                trechEM = "FIC";

                                // keeps old PAC
                                string oldPac = _Alim_PAC; 

                                // new PAC 
                                _Alim_PAC = "CABFIC";

                                // if NOT substation mode
                                if (! _modoReconf)
                                {
                                    Console.Write("Há mais de 1 trecho para alocação do Energymeter! " + _par._alim + Environment.NewLine);

                                    // linha Energymeter
                                    _linhaEM += "new line." + "SMT_" + trechEM + " bus1=" + _Alim_PAC + ".1.2.3,bus2=" + oldPac + ".1.2.3,Phases=3,r1=0,x1=0.0001,Length=0.001,Units=km" + Environment.NewLine;
                                }
                            }

                            // cria linha de alocacao Energymeter //OBS: necessario adicionar prefixo "_SMT"
                            _linhaEM += "new energymeter.carga element=line." + "SMT_" + trechEM + "," + terminal + Environment.NewLine;                            
                        }
                    }
                }

                conn.Close();
            }
        }

        // Cria Str Energymeter e comandos adicionais
        private string CriaStringMasterDSS_ParteB()
        {
            string linha = "";

            //ENERGYMETER
            linha = _linhaEM;

            // OLD CODE DEL
            /*
            if (_modoReconf)
            {
                // OBS: aloca medidor no terminal 2
                linha = Environment.NewLine + "new energymeter.carga element=line." + "SMT_"+ _par._trEM + ",terminal=1" + Environment.NewLine;
            }
            else
            {
                linha = _linhaEM;
            }*/

            // voltage bases // TODO
            if (_par._dist.Equals("44"))
            {
                linha += Environment.NewLine + "Set voltagebases=[" + _TenNom_kV + " 0.38 0.44]" + Environment.NewLine;
            }
            else // TODO padrao tensao Cemig
            {
                linha += Environment.NewLine + "Set voltagebases=[" + _TenNom_kV + " 0.24 0.22]" + Environment.NewLine;
            }

            // CalcVoltageBases
            linha += "CalcVoltageBases" + Environment.NewLine;

            // Aumento do numero de iteracoes OpenDSS
            linha += "Set MaxIter = 1000" + Environment.NewLine;
            linha += "Set MaxControlIter = 1000" + Environment.NewLine;

            // 
            linha += Environment.NewLine + "! Solve mode=daily,hour=0,number=24,stepsize=1h" + Environment.NewLine;

            if (_criaArqCoordenadas)
            {
                // Linhas de coordenadas
                linha += Environment.NewLine + Environment.NewLine + "! BusCoords " + GetNomeArqCoord() + Environment.NewLine;
            }

            linha += "ClearBusMarkers" + Environment.NewLine +
            "! AddBusMarker Bus= ,code=12,color=Green,size=5 ! substation" + Environment.NewLine +
            "! AddBusMarker bus= ,code=27,color=Green,size=3 ! NC switch in Green" + Environment.NewLine +
            "! AddBusMarker bus= ,code=27,color=Yellow,size=3 ! Yellow circle = NO switch to be closed" + Environment.NewLine +
            "! plot circuit quantity=power,dots=n,labels=n,subs=y,showloops=n,C1=Blue,C2=Blue,C3=Red,R2=0.95,R3=0.90";
            return linha;
        }

        private string CriaStrCabecalhoCustomizacao()
        {
            string linha = "! manter este comentario para compatibilidade do C#" + Environment.NewLine;

            // add linha circuito
            linha += CriaStrCircuit();

            return linha;
        }

        private string CriaStrCircuit()
        {
            // modelo
            // new circuit.alimAFNU16 bus1=BMT155673172,basekv=13.8,pu=
            return "new circuit.alim" + _par._alim
            + " bus1=" + _Alim_PAC // OBS: + ".1.2.3"
            + ",basekv=" + _TenNom_kV
            + ",pu=" + _TenOpe_pu
            + ",r1=0,x1=0.0001"
            + Environment.NewLine + Environment.NewLine;
        }

        // funcao que cria a string do arquivo masterDSS
        private string CriaStringMasterDSS_ParteA()
        {
            // limpa
            string linha = "clear" + Environment.NewLine;

            // new circuit
            linha += CriaStrCircuit();

            // curva de carga 
            //TODO NO momento o programa nao gera arquivo de curvas de carga
            linha += "Redirect ..\\0PermRes\\NovasCurvasTxt\\CurvasDeCargaDU.dss" + Environment.NewLine;

            // arquivo condutores 
            //TODO No momento o programa na gera o arquivo de condutores.
            linha += "Redirect ..\\0PermRes\\Condutores.dss" + Environment.NewLine;

            if (_structElem._temSegmentoMT)
            {
                linha += "Redirect " + _par._alim + "SegmentosMT.dss" + Environment.NewLine;
            }
            if (_structElem._temChaveMT)
            {
                linha += "Redirect " + _par._alim + "ChavesMT.dss" + Environment.NewLine;
            }
            if (_structElem._temRegulador)
            {
                linha += "Redirect " + _par._alim + "Reguladores.dss" + Environment.NewLine;
            }
            if (_structElem._temTransformador)
            {
                linha += "Redirect " + _par._alim + "Transformadores.dss" + Environment.NewLine;
            }
            if (_structElem._temSegmentoBT)
            {
                linha += "Redirect " + _par._alim + "SegmentosBT.dss" + Environment.NewLine;
            }
            if (_structElem._temRamal)
            {
                linha += "Redirect " + _par._alim + "Ramais.dss" + Environment.NewLine;
            }

            if (_structElem._temCargaMT)
            {
                linha += "Redirect " + _par._alim + "CargaMT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }

            if (_structElem._temCargaBT)
            {
                linha += "Redirect " + _par._alim + "CargaBT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }

            if (_structElem._temCapacitorMT)
            {
                linha += "Redirect " + _par._alim + "CapacitorMT.dss" + Environment.NewLine;
            }

            if (_structElem._temGeradorMT)
            {
                linha += "Redirect " + _par._alim + "GeradorMT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }

            if (_structElem._temGeradorBT)
            {
                linha += "Redirect " + _par._alim + "GeradorBT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }
            return linha;
        }
    }
}
