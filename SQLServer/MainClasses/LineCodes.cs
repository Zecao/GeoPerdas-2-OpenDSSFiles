using ExportadorGeoPerdasDSS;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.MainClasses
{
    internal class LineCodes
    {
        private static readonly string _arqCond = "Condutores.dss";
        private static SqlConnectionStringBuilder _connBuilder;
        private readonly Param _par;
        private StringBuilder _string_LineCodes;

        public LineCodes(SqlConnectionStringBuilder connBuilder, Param par)
        {
            _connBuilder = connBuilder;
            _par = par;

            ConsultaBanco();
        }

        public bool ConsultaBanco()
        {
            _string_LineCodes = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // 
                    command.CommandText = "select CodCond,Resis_ohms_km,Reat_ohms_km,CorrMax_A " +
                        "from " + _par._DBschema + "StoredCodCondutor where CodBase=@codbase ";

                    command.Parameters.AddWithValue("@codbase", _par._codBase);

                    using (var rs = command.ExecuteReader())
                    {
                        // verifica ocorrencia de elemento no banco
                        if (!rs.HasRows)
                        {
                            return false;
                        }

                        while (rs.Read())
                        {
                            // 
                            string CodCond = rs["CodCond"].ToString();
                            string Resis_ohms_km = rs["Resis_ohms_km"].ToString();
                            string Reat_ohms_km = rs["Reat_ohms_km"].ToString();
                            string CorrMax_A = rs["CorrMax_A"].ToString();

                            string linha = "";
                            linha += "new linecode." + CodCond + "_1 nPhases=1,r1=" + Resis_ohms_km + ",x1=" + Reat_ohms_km + ",units=km,normamps=" + CorrMax_A + Environment.NewLine;
                            linha += "new linecode." + CodCond + "_2 nPhases=2,r1=" + Resis_ohms_km + ",x1=" + Reat_ohms_km + ",units=km,normamps=" + CorrMax_A + Environment.NewLine;
                            linha += "new linecode." + CodCond + "_3 nPhases=3,r1=" + Resis_ohms_km + ",x1=" + Reat_ohms_km + ",units=km,normamps=" + CorrMax_A + Environment.NewLine;
                            linha += "new linecode." + CodCond + "_4 nPhases=4,r1=" + Resis_ohms_km + ",x1=" + Reat_ohms_km + ",units=km,normamps=" + CorrMax_A + Environment.NewLine;

                            _string_LineCodes.Append(linha);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        private string GetNomeArq()
        {
            return _par._pathAlim + _par._alim + _arqCond;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_string_LineCodes.ToString(), GetNomeArq());
        }
    }
}
