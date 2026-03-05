using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using Npgsql;

namespace ExportadorGeoPerdasDSS
{
    class SegmentoMT
    {        
        public static readonly string _arqDicReliability = "dictionayReliability.xlsx"; // arquivo do Excel com taxas de falhas (FaultRate) dos lineCodes
        private static readonly string _segmentosMT = "SegmentosMT.dss";
        private static readonly string _coordMT = "CoordMT.csv";
        private static NpgsqlConnectionStringBuilder _connBuilder;

        private readonly bool _criaDispProtecao;
        private StringBuilder _arqSegmentoMT;
        private StringBuilder _arqCoord;
        private readonly ModeloSDEE _SDEE;
        private readonly Dictionary<string, double> _dicLineCodeFaultRate;

        private readonly Param _par;

        public SegmentoMT(NpgsqlConnectionStringBuilder connBuilder, ModeloSDEE sdee, Param par, bool criaDispProtecao)
        {
            _par = par;
            _connBuilder = connBuilder;
            _SDEE = sdee;
            _criaDispProtecao = criaDispProtecao;

            // preenche dictionary lineCode X FaultRate
            if (_criaDispProtecao)
            {
                // preenche Dic de soma Carga Mensal - Utilizado por CargaMT e CargaBT
                _dicLineCodeFaultRate = XLSXFile.XLSX2Dictionary(GetNomeArqDicReliability());

            }
        }

        private string GetNomeArqDicReliability()
        {
            return _par._path + _par._permRes + _arqDicReliability;
        }

        //modelo
        //new line.TR1113 bus1=1575B.1.2.3,bus2=1568B.1.2.3,Phases=3,Linecode=CAB103_3_3,Length=0.038482,Units=km
        public bool ConsultaStoredSegmentoMT(bool _modoReconf)
        {
            _arqSegmentoMT = new StringBuilder();

            using (NpgsqlConnection conn = new NpgsqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                // TODO CoordPAC1_x,CoordPAC1_y,CoordPAC2_x,CoordPAC2_y 
                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodSegmMT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km"
                    + " from " + _par._DBschema + "StoredSegmentoMT";

                    if (_modoReconf)
                    {
                        command.CommandText += " where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += " where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string lineCode = rs["CodCond"].ToString();

                            string linha = "new line." + "SMT_" + rs["CodSegmMT"] //OBS1:
                                + " bus1=" + rs["CodPonAcopl1"] + fases //OBS1:
                                + ",bus2=" + rs["CodPonAcopl2"] + fases //OBS1:
                                + ",Phases=" + numFases
                                + ",Linecode=" + lineCode + "_" + numFases
                                + ",Length=" + rs["Comp_km"]
                                + ",Units=km";

                            if (_criaDispProtecao)
                            {
                                //obtem taxa de falha de acordo com o lineCode
                                string faultRate = GetFaultRate(lineCode);

                                linha += ",FaultRate=" + faultRate + ",Pctperm=20,Repair=3" + Environment.NewLine;
                            }
                            else
                            {
                                linha += Environment.NewLine;
                            }

                            _arqSegmentoMT.Append(linha);

                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        // 
        private string GetFaultRate(string lineCode)
        {
            if (_dicLineCodeFaultRate.ContainsKey(lineCode))
            {
                return _dicLineCodeFaultRate[lineCode].ToString();
            }
            else
            {
                return "0.06774";
            }
        }

        //modelo
        //new line.TR1113 bus1=1575B.1.2.3,bus2=1568B.1.2.3,Phases=3,Linecode=CAB103_3_3,Length=0.038482,Units=km
        public bool ConsultaBusCoord(bool _modoReconf)
        {
            _arqCoord = new StringBuilder();

            using (NpgsqlConnection conn = new NpgsqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "";

                    if (_modoReconf)
                    {
                        command.CommandText += "select CodPonAcopl1 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema
                            + "StoredSegmentoMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")" +
                        " union " +
                        "select CodPonAcopl2 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema + "StoredSegmentoMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";

                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "select CodPonAcopl1 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema
                            + "StoredSegmentoMT where CodBase=@codbase and CodAlim=@CodAlim"
                            + " union "
                            + "select CodPonAcopl2 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema
                            + "StoredSegmentoMT where CodBase=@codbase and CodAlim=@CodAlim";

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
                            string linhaPAC1 = rs["PAC"] + "," + rs["CoordPAC1_x"] + "," + rs["CoordPAC1_y"] + Environment.NewLine;

                            _arqCoord.Append(linhaPAC1);
                        }
                    }

                }

                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "";

                    if (_modoReconf)
                    {
                        command.CommandText += "select CodPonAcopl1 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema
                            + "StoredSegmentoMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")" +
                        " union " +
                        "select CodPonAcopl2 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema + "StoredSegmentoMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";

                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "select CodPonAcopl1 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema
                            + "StoredSegmentoBT where CodBase=@codbase and CodAlim=@CodAlim"
                            + " union "
                            + "select CodPonAcopl2 as 'PAC', CoordPAC1_x,CoordPAC1_y from " + _par._DBschema
                            + "StoredSegmentoBT where CodBase=@codbase and CodAlim=@CodAlim";

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
                            string linhaPAC1 = "BBT" + rs["PAC"] + "," + rs["CoordPAC1_x"] + "," + rs["CoordPAC1_y"] + Environment.NewLine;

                            _arqCoord.Append(linhaPAC1);
                        }
                    }

                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        public string GetNomeArq()
        {
            return _par._pathAlim + _par._alim + _segmentosMT;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoMT.ToString(), GetNomeArq());
        }

        internal void GravaArqCoord()
        {
            ArqManip.SafeDelete(GetNomeArqCoord());

            ArqManip.GravaEmArquivo(_arqCoord.ToString(), GetNomeArqCoord());
        }

        private string GetNomeArqCoord()
        {
            return _par._pathAlim + _par._alim + _coordMT;
        }

        // Une 
        internal bool UneSE(bool _modoReconfiguracao)
        {
            //condicao de saida 
            if (!_modoReconfiguracao)
                return false;

            List<string> dicCodPonAcopl = new List<string>();

            using (NpgsqlConnection conn = new NpgsqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodAlim,CodPonAcopl from " + _par._DBschema + "Storedcircmt where codbase = @codbase "
                                        + "and codalim in (" + _par._conjAlim + ")";
                    command.Parameters.AddWithValue("@codbase", _par._codBase);

                    using (var rs = command.ExecuteReader())
                    {
                        // verifica ocorrencia de elemento no banco
                        if (!rs.HasRows)
                        {
                            return false;
                        }

                        // Preenche lista com os 2 cod acoplamento
                        while (rs.Read())
                        {
                            dicCodPonAcopl.Add(rs["CodPonAcopl"].ToString());
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }

            string linha = "";

            // pega os pels 
            for (int i = 0; i < dicCodPonAcopl.Count - 1; i++)
            {
                //new line.SMT_FIC_SE8 bus1 = CABFIC.1.2.3,bus2 = FIC.1.2.3,Phases = 3,r1 = 0,x1 = 0.0001,Length = 0.001,Units = km

                string fic = "FIC_SE" + i.ToString();

                linha += "new line." + fic
                    + " bus1=" + dicCodPonAcopl[i] + ".1.2.3"
                    + ",bus2=" + dicCodPonAcopl[i + 1] + ".1.2.3"
                    + ",Phases=" + "3"
                    + ",r1=0,x1=0.0001"
                    + ",Length=" + "0.001"
                    + ",Units=km" + Environment.NewLine;
            }

            linha += "new line." + "FIC_SE" + dicCodPonAcopl.Count.ToString()
                + " bus1=" + dicCodPonAcopl.Last() + ".1.2.3"
                + ",bus2=" + "FIC" + ".1.2.3" // OBS: FIC
                + ",Phases=" + "3"
                + ",r1=0,x1=0.0001"
                + ",Length=" + "0.001"
                + ",Units=km" + Environment.NewLine;

            _par._trEM = "FIC";

            // Cria mais 1 trecho ficticio unido ao ultimo de dicCodPonAcopl
            linha += "new line." + "SMT_" + _par._trEM // OBS: if changing the name CABFIC must chang in CircMT class, as well.
                + " bus1=" + "CABFIC" +  ".1.2.3" // OBS: if changing the name CABFIC must chang in CircMT class, as well.
                + ",bus2=" + "FIC" + ".1.2.3" // OBS: FIC
                + ",Phases=" + "3"
                + ",r1=0,x1=0.0001"
                + ",Length=" + "0.001"
                + ",Units=km" + Environment.NewLine;

            _arqSegmentoMT.Append(linha);

            return true;
        }
    }
}
