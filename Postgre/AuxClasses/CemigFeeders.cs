
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
    }
}
