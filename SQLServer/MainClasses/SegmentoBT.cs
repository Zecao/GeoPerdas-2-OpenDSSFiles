using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class SegmentoBT
    {
        // membros privados
        private static readonly string _segmentosBT = "SegmentosBT.dss";
        private static SqlConnectionStringBuilder _connBuilder;
        private readonly Param _par;
        private StringBuilder _arqSegmentoBT;

        public SegmentoBT(SqlConnectionStringBuilder connBuilder, Param par)
        {
            _par = par;
            _connBuilder = connBuilder;
        }

        //modelo
        // new line.TBT10260958 bus1=BBT2593020.1.2.3.0,bus2=BBT2593027.1.2.3.0,Phases=3,Linecode=BT107_BT107_7_3_1,Length=0.038957,Units=km
        // CodBase	CodSegmBT	CodAlim	CodTrafo	CodPonAcopl1	CodPonAcopl2	CodFas	CodCond	Comp_km	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	Ordm	De	Para
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText = "select CodSegmBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km" //,CoordPAC1_x,CoordPAC1_y,CoordPAC2_x,CoordPAC2_y
                            + " from " + _par._DBschema + "StoredSegmentoBT "
                            + "where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = "select CodSegmBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km" //,CoordPAC1_x,CoordPAC1_y,CoordPAC2_x,CoordPAC2_y
                            + " from " + _par._DBschema + "StoredSegmentoBT "
                            + "where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString(), _par._modelo4condutores);
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string comp_km = rs["Comp_km"].ToString();

                            string linha;

                            //NEW CODE
                            if (!_par._modelo4condutores)
                            {
                                linha = "new line.SBT_" + rs["CodSegmBT"].ToString()
                                + " bus1=" + rs["CodPonAcopl1"] + fases //+ "BBT"
                                + ",bus2=" + rs["CodPonAcopl2"] + fases //+ "BBT"
                                + ",Phases=" + numFases
                                + ",Linecode=" + rs["CodCond"].ToString() + "_" + numFases
                                + ",Length=" + rs["Comp_km"].ToString()
                                + ",Units=km" + Environment.NewLine;
                            }
                            else
                            //altera linecod 
                            {
                                double numFasesD = double.Parse(numFases);
                                numFasesD++;
                                numFases = numFasesD.ToString();

                                double comp_kmD = double.Parse(comp_km);

                                // se comp < 1metro
                                if (comp_kmD < 0.00049)
                                {
                                    string altLineCode = ",r1=0.001,r0=0.001,x1=0,x0=0,c1=0,c0=0,switch=T";

                                    linha = "new line.SBT_" + rs["CodSegmBT"].ToString()
                                    + " phases=" + numFases
                                    + ",bus1=" + rs["CodPonAcopl1"] + fases // + "BBT"
                                    + ",bus2=" + rs["CodPonAcopl2"] + fases // + "BBT"
                                    + altLineCode
                                    + ",length=0.001" //atribui 1metro 
                                    + ",units=km" + Environment.NewLine;
                                }
                                else
                                {

                                    linha = "new line.SBT_" + rs["CodSegmBT"].ToString()
                                    + " phases=" + numFases
                                    + ",bus1=" + rs["CodPonAcopl1"] + fases // + "BBT"
                                    + ",bus2=" + rs["CodPonAcopl2"] + fases // + "BBT"
                                    + ",linecode=" + rs["CodCond"].ToString() + "_" + numFases
                                    + ",length=" + comp_km
                                    + ",units=km" + Environment.NewLine;
                                }

                            }

                            _arqSegmentoBT.Append(linha);
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
            return _par._pathAlim + _par._alim + _segmentosBT;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
