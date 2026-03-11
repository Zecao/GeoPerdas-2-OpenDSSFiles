using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class RamalBT
    {
        // membros privados
        private static readonly string _ramais = "Ramais.dss";
        private static SqlConnectionStringBuilder _connBuilder;
        private Param _par;
        private readonly string _alim;
        private StringBuilder _arqSegmentoBT;

        public RamalBT(SqlConnectionStringBuilder connBuilder, Param par)
        {
            _par = par;
            _alim = par._alim;
            _connBuilder = connBuilder;
        }

        //modelo
        // new line.RML10041834 bus1=BBT2049281.1.0,bus2=R10041834.1.0,Phases=1,Linecode=CABBTG01_4_1,Length=0.017,Units=km
        //CodBase	CodRmlBT	CodAlim	CodTrafo	CodPonAcopl1	CodPonAcopl2	CodFas	CodCond	Comp_km	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	Ordm	De	Para
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
                        command.CommandText = "select CodRmlBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km,Descr from " + _par._DBschema + "StoredRamalBT "
                            + "where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = "select CodRmlBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km,Descr from " + _par._DBschema + "StoredRamalBT "
                            + "where CodBase=@codbase and CodAlim=@CodAlim";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                        command.Parameters.AddWithValue("@CodAlim", _alim);
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
                            // OLD CODE
                            // comprimento ramal interno
                            //string compRmlInt = "3";
                            /* // TODO: aumento de 3metros do ramal
                             double compRml_D = Double.Parse(compRml) + 0.003;

                             compRml = compRml_D.ToString();

                             // atribui compRml interno ao 
                             if (!compRmlInt.Equals(""))
                             {
                                 compRml = compRmlInt;
                             }
                             */

                            string compRml = rs["Comp_km"].ToString();
                            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString(), _par._modelo4condutores);
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string linha;

                            //NEW CODE
                            if (!_par._modelo4condutores)
                            {
                                linha = "new line.RBT_" + rs["CodRmlBT"].ToString()
                                    + " bus1=" + rs["CodPonAcopl1"] + fases //+ "BBT" 
                                    + ",bus2=" + rs["CodPonAcopl2"] + fases //+ "RML"
                                    + ",Phases=" + numFases
                                    + ",Linecode=" + rs["CodCond"].ToString() + "_" + numFases // OBS: altera linecode
                                    + ",Length=" + compRml
                                    + ",Units=km" + Environment.NewLine;
                            }
                            else
                            {
                                double numFasesD = double.Parse(numFases);
                                numFasesD++;
                                numFases = numFasesD.ToString();

                                linha = "new line.RBT_" + rs["CodRmlBT"].ToString()
                                + " bus1=" + rs["CodPonAcopl1"] + fases //+ "BBT"
                                + ",bus2=" + rs["CodPonAcopl2"] + fases //+ "RML"
                                + ",Phases=" + numFases
                                + ",Linecode=" + rs["CodCond"].ToString() + "_" + numFases // OBS: altera linecode
                                + ",Length=" + compRml
                                + ",Units=km" + Environment.NewLine;
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
            return _par._pathAlim + _alim + _ramais;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
