using AuxClasses;
using ExportadorGeoPerdasDSS;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExportadorArqDSS
{
    class GeoPerdas2DSSFiles
    {
        // parametros banco de dados                                               
        public static string _dbms = "Postgre"; //"SQLServer" "Postgre"
        public static string _dataSource = "localhost"; // PostgreSQL 
        public static string _banco = "GEO_SIGR_PERDAS_CEMIG_2024"; // servidor SGBD
        public static string _schema = "GeoPerdas2023b."; // SQlServer: dbo. Postgre: GeoPerdas2023b.

        // path
        //public static readonly string _path = @"D:\CopiaDropbox\0_BDGDs\_CPFL\_DSS_2025\";
        //public static readonly string _path = @"D:\CopiaDropbox\0_BDGDs\2_Equatorial_MA\DSS\";
        public static readonly string _path = @"D:\CopiaDropbox\0_ERA\2026\07_RTP_Cemig\1_alim2024\";

        // codbase
        public static int _codBase = 2024124950; // CPFL:20231263 Cemig:2022124950 EquatorialMA:20241237 Energisa-MT: 202412405 MS:202312404
        public static string _dist = "4950"; //

        // mes e ano para a geracao dos arquivos de carga BT e MT
        public static int _iMes = 12;
        public static string _ano = "2024"; //

        // 
        public static bool _criaTodosOsMeses = false;  // flag p/ criar todos os meses de carga MT BT e geradores
        public static bool _criaArqCoordenadas = true; // flag p/ criar arq coordenadas

        /* // parametros PVSystem / generators
        Parameters order _modelPVSystems, _invControlModeMV, _varFollowInvMV, _PVPowerFactorMV

        _modelPVSystems     -> False = exports PVSystems as generator model=1
        _geraInvControl     _> true to generate invControl
        _invControlModeMV   -> Default: "VOLTVAR" "voltwatt"
        _varFollowInvMV     -> Default: false (Have to set False to enable Inverter night mode)
        _voltVarcurve =     -> Default: "voltvar_c"
        */
        public static readonly PVSystemPar _pvMV = new PVSystemPar(true, false); 
        public static readonly PVSystemPar _pvLV = new PVSystemPar(false, false);

        // 
        public static bool _criaDispProtecao = false; // flag p/ dispositivos de protecao (Recloser e Fuses) && taxas de falhas em lines
        public static bool _modelo4condutores = false; //modelo 4 condutor BT

        // cria arquivo DSS unificando alimentadores (e.g. estudos de reconfiguracao)
        public static bool _genAllSubstation = false; // Generates all substation feeders as one. Uses the first feeder for directory name.

        // arquivo txt com lista de alimentadores
        public static string _arqLstAlimentadores = "lstAlimentadores.m";

        // arquivo do Excel com curvas individuais dos clientes primarios
        public static readonly string _arqCurvaCargaCliMT = "curvasTipicasClientesMT_2018.xlsx";

        // prefixo arquivo .txt de feriados
        public static string _feriado = "Feriados"; //arquivo de feriado

        // TODO FIX this hard coded 
        // sub diretorio recursos permanentes
        public static string _permRes = "0PermRes\\";

        /*
        // parametros CPFL CELPE ALAGOAS        
        private readonly static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: false, modeloCarga: "PCONST",
           reatanciaTrafos: true);
        */
        
        // parametros CEMIG 
        private readonly static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: false, modeloCarga: "PCONST",
           reatanciaTrafos: true);        
        
        // FIM parametros configuraveis usuario

        //
        public static Param _par;

        private static NpgsqlConnectionStringBuilder _connBuilder;

        private static StrBoolElementosSDE _structElem;

        // arquivo de coordenadas
        private static readonly string _coordMT = "CoordMT.csv";

        // obtem Lista com numero de feriados mes X Mes
        private static List<List<int>> _numDiasFeriadoXMes;

        // utilizado por CargaMT e CargaBT
        private static Dictionary<string, double> _somaCurvaCargaDiariaPU;

        // OLD CODE
        // arquivo do Excel com somatorio em PU das curvas de carga
        public static readonly string _arqConsumoMensalPU = "somaCurvasCargaPU.xlsx";
        // utilizado por CargaMT e CargaBT
        private static Dictionary<string, List<string>> _curvasTipicasClientesMT;

        static void Main() //string[] args
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            _connBuilder = new NpgsqlConnectionStringBuilder();
            _connBuilder.Host = _dataSource;
            _connBuilder.Database = _banco;
            _connBuilder.Username = "postgres";
            _connBuilder.Password = "educ1986";

            // variaveis auxiliares
            CarregaVariaveisAux();

            // se modo reconfiguracao
            if (_genAllSubstation)
            {
                GenSubstationDSSFiles();
            }
            else
            {
                GeneratesFeedersDSSFiles();
            }

            Console.Write("Fim!");
            Console.ReadKey();
        }

        // cria arquivos DSS
        private static void CriaArquivosDSS()
        {
            //Cria o diretório do alimentador, caso não exista
            if (!System.IO.Directory.Exists(_par._pathAlim))
            {
                System.IO.Directory.CreateDirectory(_par._pathAlim);
            }

            // Create Condutores.dss file
            //Create_LineCodes_DSSFile();

            // UNDER CONSTRUCTION
            // Create LoadShape .dss files
            //Create_LoadShapes_DSSFile();            
            
            
            // Segmento MT
            CriaSegmentoMTDSS();
            
            // Se nao tem segmento, aborta 
            if (!_structElem._temSegmentoMT)
            {
                return;
            }
            
            // Regulador MT
            CriaReguladorMTDSS();
            
            // Chave MT
            CriaChaveMT();
            
            // Transformador MT
            CriaTransformadorMTMTMTBTDSS();
            
            // Capacitor
            if (_SDEE._incluirCapacitoresMT)
            {
                CriaCapacitorMTDSS();
            }
            
            CriaSegmentoBTDSS();
            
            // Ramais 
            CriaRamaisDSS();
            
            // Carga MT
            CriaCargaMTDSS();           

            // Gerador MT
            CriaGeradorMT();
            
            // Gerador BT
            CriaGeradorBT();            
            
            // Carga BT
            CriaCargaBTDSS();
            
            // arquivo cabecalho
            CriaCabecalhoDSS();
        }

        // Generates entire substation
        private static void GenSubstationDSSFiles()
        {
            // populates lstAlim with all feeders from txt file "Alimentadores.m"

            List<string> lstAlim = CemigFeeders.GetAllFeedersFromTxtFile(GetNomeArqLstAlimentadores());

            _par = new Param(_path, _permRes, _codBase, _modelo4condutores, _schema, "", _ano, _pvMV, _pvLV, _dist);

            // para cada SE da lista
            foreach (string alim in lstAlim)
            {
                // sets current feeder
                _par.SetCurrentAlim(alim);

                // gets substation // TODO refactory
                string substation = System.Text.RegularExpressions.Regex.Replace(alim, @"[\d-]", string.Empty);

                
                //removes 4 char from Alagoas Substation names
                if ( _dist.Equals("44") )
                {
                    substation = substation.Substring(0, 3);
                }

                // gets all feeders name from substation //TODO move to Param ?
                bool ret = CemigFeeders.GetAllFeedersFromSubstationString(substation, _connBuilder, _par);

                if (_par._conjAlim == null)
                {
                    Console.Write(substation + " não encontrado!\n");
                    continue;
                }

                // creates dss files
                CriaArquivosDSS();

                Console.Write("SE do " + alim + " .dss criados\n");
            }
        }

        // Generates entire substation
        private static void GeneratesFeedersDSSFiles()
        {
            // lista de alimentadores
            List<string> lstAlim = CemigFeeders.GetAllFeedersFromTxtFile(GetNomeArqLstAlimentadores());

            _par = new Param(_path, _permRes, _codBase, _modelo4condutores, _schema, "", _ano, _pvMV, _pvLV, _dist);

            // para cada alimentador da lista
            foreach (string alim in lstAlim)
            {
                // sets current feeder
                _par.SetCurrentAlim(alim);

                // creates dss files
                CriaArquivosDSS();

                Console.Write(alim + " .dss criados\n");

                // TODO
                //ClassificaCurvaBT curvaBT = new ClassificaCurvaBT(_connBuilder, _par);
                //Add_IMag_transformer obj = new Add_IMag_transformer(_par); 
            }
        }

        private static string GetNomeArqLstAlimentadores()
        {
            return _path + _permRes + _arqLstAlimentadores;
        }

        private static void CriaGeradorBT()
        {
            if (_criaTodosOsMeses)
            {
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    _iMes = i;

                    //
                    CriaGeradorBTPvt();
                }
            }
            else
            {
                //
                CriaGeradorBTPvt();
            }
        }

        private static void CriaGeradorBTPvt()
        {
            GeradorBT oGerBT = new GeradorBT(_connBuilder, _par, _iMes);

            // realiza consulta  
            _structElem._temGeradorBT = oGerBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temGeradorBT)
            {
                oGerBT.GravaEmArquivo();
            }
        }

        private static void CarregaVariaveisAux()
        {
            // obtem Lista com numero de feriados mes X Mes
            _numDiasFeriadoXMes = AuxFunc.Feriados(GetNomeArqFeriado());

            // preenche Dic de soma Carga Mensal - Utilizado por CargaMT e CargaBT
            _somaCurvaCargaDiariaPU = XLSXFile.XLSX2Dictionary(GetNomeArqConsumoMensalPU());

            // preenche Dic com curvas de carga INDIVIDUAIS da CargaMT
            if (_SDEE._utilizarCurvaDeCargaClienteMTIndividual)
            {
                _curvasTipicasClientesMT = XLSXFile.XLSX2DictString(GetNomeArqCurvaCargaCliMT());
            }
        }

        private static string GetNomeArqCurvaCargaCliMT()
        {
            return _path + _permRes + _arqCurvaCargaCliMT;
        }

        private static string GetNomeArqConsumoMensalPU()
        {
            return _path + _permRes + _arqConsumoMensalPU;
        }

        // nome arquivo feriado
        private static string GetNomeArqFeriado()
        {
            return _path + _permRes + _feriado + _ano + ".txt";
        }

        private static void CriaGeradorMT()
        {
            if (_criaTodosOsMeses)
            {
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    _iMes = i;

                    //
                    CriaGeradorMTPvt();
                }
            }
            else
            {
                //
                CriaGeradorMTPvt();
            }
        }

        //
        private static void CriaGeradorMTPvt()
        {
            GeradorMT oGerMT = new GeradorMT(_connBuilder, _par, _iMes);

            // realiza consulta  
            _structElem._temGeradorMT = oGerMT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temGeradorMT)
            {
                oGerMT.GravaEmArquivo();
            }
        }

        private static void CriaChaveMT()
        {
            ChaveMT oChaveMT = new ChaveMT(_connBuilder, _par, _criaDispProtecao);

            // realiza consulta 
            _structElem._temChaveMT = oChaveMT.ConsultaBanco(_genAllSubstation);

            // _temChaveMTno 
            if (_structElem._temChaveMT)
            {
                oChaveMT.GravaEmArquivo();
            }
        }

        private static void CriaReguladorMTDSS()
        {
            Regulador oRT = new Regulador(_connBuilder, _par);

            // realiza consulta StoredReguladorMT 
            _structElem._temRegulador = oRT.ConsultaStoredReguladorMT(_genAllSubstation);

            // _temRegulador
            if (_structElem._temRegulador)
            {
                oRT.GravaEmArquivo();
            }
        }

        // cria arquivo dss de segmentos de MT
        private static void CriaSegmentoMTDSS()
        {
            SegmentoMT oSegMT = new SegmentoMT(_connBuilder, _SDEE, _par, _criaDispProtecao);

            // realiza consulta StoredSegmentoMT 
            _structElem._temSegmentoMT = oSegMT.ConsultaStoredSegmentoMT(_genAllSubstation);

            // une cabeca alimentador
            oSegMT.UneSE(_genAllSubstation);

            // _temSegmentoMT
            if (_structElem._temSegmentoMT)
            {
                oSegMT.GravaEmArquivo();

                // se modo criar arq coordenadas
                if (_criaArqCoordenadas)
                {
                    //
                    oSegMT.ConsultaBusCoord(_genAllSubstation);

                    //
                    oSegMT.GravaArqCoord();
                }
            }
            // se alimentador nao tem segmento MT aborta
            else
            {
                Console.Write(_par._alim + ": sem segmento MT. Abortando!\n");
            }
        }

        private static void CriaTransformadorMTMTMTBTDSS()
        {
            Trafo oTrafo = new Trafo(_connBuilder, _par, _SDEE);

            // realiza consulta StoredReguladorMT 
            _structElem._temTransformador = oTrafo.ConsultaBanco(_genAllSubstation);

            if (_structElem._temTransformador)
            {
                oTrafo.GravaEmArquivo();
            }
        }

        private static void CriaSegmentoBTDSS()
        {
            SegmentoBT oSegBT = new SegmentoBT(_connBuilder, _par);

            // realiza consulta 
            _structElem._temSegmentoBT = oSegBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temSegmentoBT)
            {
                oSegBT.GravaEmArquivo();
            }
        }

        private static void CriaRamaisDSS()
        {
            RamalBT oRamal = new RamalBT(_connBuilder, _par);

            // realiza consulta 
            _structElem._temRamal = oRamal.ConsultaBanco(_genAllSubstation);

            if (_structElem._temRamal)
            {
                oRamal.GravaEmArquivo();
            }
        }

        private static void CriaCargaMTDSS()
        {
            if (_criaTodosOsMeses)
            {
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    // seta mes corrente
                    _iMes = i;

                    // 
                    CriaCargaMTDSSPvt();
                }
            }
            else
            {
                // 
                CriaCargaMTDSSPvt();
            }
        }

        private static void CriaCargaMTDSSPvt()
        {
            CargaMT oCargaMT = new CargaMT(_connBuilder, _iMes, _ano, _numDiasFeriadoXMes,
            _somaCurvaCargaDiariaPU, _SDEE, _par, _curvasTipicasClientesMT);

            // realiza consulta 
            _structElem._temCargaMT = oCargaMT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temCargaMT)
            {
                oCargaMT.GravaEmArquivo();
            }
        }

        private static void CriaCargaBTDSS()
        {
            if (_criaTodosOsMeses)
            {
                // TODO otimizar em uma consulta
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    // seta mes corrente
                    _iMes = i;

                    //
                    CriaCargaBTDSSPvt();
                }
            }
            else
            {
                //
                CriaCargaBTDSSPvt();
            }
        }

        private static void CriaCargaBTDSSPvt()
        {
            CargaBT oCargaBT = new CargaBT(_connBuilder, _iMes, _ano, _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU,
                _SDEE, _par);

            // realiza consulta 
            _structElem._temCargaBT = oCargaBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temCargaBT)
            {
                oCargaBT.GravaEmArquivo();
            }

        }

        // cria capacitorMT
        private static void CriaCapacitorMTDSS()
        {
            CapacitorMT oCap = new CapacitorMT(_connBuilder, _par);

            // realiza consulta 
            _structElem._temCapacitorMT = oCap.ConsultaBanco(_genAllSubstation);

            if (_structElem._temCapacitorMT)
            {
                oCap.GravaEmArquivo();
            }
        }

        // cria arquivos cabecalhos ou master. 3 arquivos sao criados com o seguinte modelo, onde Alim = nome do alimentador:
        // Alim.dss -> arquivo utilizado pelo projeto ExecutorOpenDSS, com informacoes iniciais do alimentador (eg: definicao do circuito)
        // AlimAnualA.dss -> arquivo p/ ser utilizado pelo usuario na GUI do OpenDSS. 
        // AlimAnulaB.dss -> arquivo utilizado pelo projeto ExecutorOpenDSS, com informacoes finais do alimentado (eg: solve).
        private static void CriaCabecalhoDSS()
        {
            CircMT oCircMT = new CircMT(_connBuilder, _par, _genAllSubstation, _criaArqCoordenadas,
                _structElem, _iMes, _coordMT);

            // realiza consulta 
            oCircMT.ConsultaStoredCircMT();

            // 
            if (! oCircMT._strCab.Equals(""))
            {
                // grava arquivo para ser utilizado pela OpenDSS 
                ArqManip.SafeDelete(GetNomeArq_Master());

                ArqManip.GravaEmArquivo(oCircMT._stringMasterDSS_A, GetNomeArq_Master()); 

                // arquivo para ser utilizado pela customizacao COM do OpenDSS
                ArqManip.SafeDelete(GetNomeArq_ForCustomCSharp());

                ArqManip.GravaEmArquivo(oCircMT._strCab, GetNomeArq_ForCustomCSharp());

                // arquivo para ser utilizado pela customizacao COM do OpenDSS
                ArqManip.SafeDelete(GetNomeArq_AnualB());

                ArqManip.GravaEmArquivo(oCircMT._stringMasterDSS_B, GetNomeArq_AnualB());
            }
        }

        private static string GetNomeArq_AnualB()
        {
            return _par._pathAlim + _par._alim + "AnualB.dss";
        }

        private static string GetNomeArq_ForCustomCSharp()
        {
            return _par._pathAlim + _par._alim + ".dss";
        }

        private static string GetNomeArq_Master()
        {
            return _par._pathAlim + _par._alim + "AnualA.dss";
        }
    }
}

