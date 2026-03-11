using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace ExportadorGeoPerdasDSS
{
    class CemigFeeders
    {
        // Reads the ".m" txt file containing the feeders names
        // Splits the lines on '%' comment char, mantaining only the feeder name.
        public static List<string> GetAllFeedersFromTxtFile(string arquivo)
        {
            //Variável que armazenará a lista com os alimentadores
            List<string> alimentadores = new List<string>();

            //Bloco que trata o arquivo, abrindo e fechando-o
            if (File.Exists(arquivo))
            {
                using (StreamReader sr = new StreamReader(arquivo))
                {
                    //Variável para armazenar a linha atual do arquivo
                    String linha;
                    //Lê a próxima linha até o fim do arquivo
                    while ((linha = sr.ReadLine()) != null)
                    {
                        //Caso haja uma linha em branco, linha[0] retornará um erro.
                        //O try/catch ignora o erro e passa para a próxima linha
                        try
                        {
                            //Se a linha começa com %, ignorar pois é comentário
                            if (!linha[0].Equals('%'))
                            {
                                //Adiciona a linha para a lista
                                alimentadores.Add(linha.Split('%')[0].Trim());
                            }

                        }
                        catch { }
                    }
                }
                return alimentadores;
            }
            else
            {
                throw new FileNotFoundException("Arquivo " + arquivo + " não encontrado.");
            }

        }

        //Transforms the feeder file string in a Substation list, removing the number after the name.
        public static List<string> GetAllSubstationFromTxtFile(string arquivo)
        {
            //Variável que armazenará a lista com os alimentadores
            List<string> substation = new List<string>();

            //Bloco que trata o arquivo, abrindo e fechando-o
            if (File.Exists(arquivo))
            {
                using (StreamReader sr = new StreamReader(arquivo))
                {
                    //Variável para armazenar a linha atual do arquivo
                    String linha;

                    //Lê a próxima linha até o fim do arquivo
                    while ((linha = sr.ReadLine()) != null)
                    {
                        //Caso haja uma linha em branco, linha[0] retornará um erro.
                        //O try/catch ignora o erro e passa para a próxima linha
                        try
                        {
                            //Se a linha começa com %, ignorar pois é comentário
                            if (!linha[0].Equals('%'))
                            {
                                // gets feeder name
                                linha = linha.Split('%')[0].Trim();

                                // 
                                linha = System.Text.RegularExpressions.Regex.Replace(linha, @"[\d-]", string.Empty);

                                //Adiciona a linha para a lista
                                substation.Add(linha);
                            }

                        }
                        catch { }
                    }
                }
                return substation;
            }
            else
            {
                throw new FileNotFoundException("Arquivo " + arquivo + " não encontrado.");
            }

        }

        //Get all feeders in a string separated by ',' from a substation name
        public static bool GetAllFeedersFromSubstationString(string sub, SqlConnectionStringBuilder con, Param par)
        {
            // OBS: a SE deve ter nome
            // obtem lstAlim da SE
            List<string> lstAlim = GetLstAlimSE(sub, con, par);

            if (lstAlim.Count == 0)
            {
                return false;
            }
            // cria string com a uniao dos alimentadores 
            string lstAlimSE = UneStringAlim(lstAlim);

            // adds lst feeders in _par object
            par._conjAlim = lstAlimSE;

            return true;
        }

        private static string UneStringAlim(List<string> lstAlim)
        {
            string conjAlims;

            // inicializacao 
            conjAlims = "'";

            // para cada alimentador da lista
            foreach (string alim in lstAlim)
            {
                conjAlims += alim;

                if (string.Equals(alim, lstAlim.Last()))
                {
                    conjAlims += "'";
                }
                else
                {
                    conjAlims += "','";
                }
            }
            return conjAlims;
        }

        private static List<string> GetLstAlimSE(string codSE, SqlConnectionStringBuilder _connBuilder, Param par)
        {
            List<string> lstAlim = new List<string>();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                //consulta a banco 
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodAlim from " + par._DBschema + "StoredCircMT "
                        + "where CodBase=@codbase and CodSub=@codSe";
                    command.Parameters.AddWithValue("@codbase", par._codBase);
                    command.Parameters.AddWithValue("@codSe", codSE);

                    using (var rs = command.ExecuteReader())
                    {
                        // verifica ocorrencia de elemento no banco
                        if (!rs.HasRows)
                        {
                            return lstAlim;
                        }

                        while (rs.Read())
                        {
                            lstAlim.Add(rs["CodAlim"].ToString());
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }

            return lstAlim;
        }

    }
}
