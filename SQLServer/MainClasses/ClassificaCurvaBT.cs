using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class ClassificaCurvaBT
    {
        private static SqlConnectionStringBuilder _connBuilder;
        private readonly Param _par;
        private StringBuilder _arqSegmentoBT;
        private Dictionary<string, List<double>> _percentuaisCurva;
        public ClassificaCurvaBT(SqlConnectionStringBuilder connBuilder,
            Param par)
        {
            _par = par;
            _connBuilder = connBuilder;

            carregaExcelPercentuais();

            ConsultaBanco();

            GravaEmArquivo();
        }

        public string GetNomeArq()
        {
            return _par._path + "UCBT_novasCurvas.csv";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }

        public void carregaExcelPercentuais()
        {
            // TODO 
            // preenche Dic de soma Carga Mensal - Utilizado por CargaMT e CargaBT
            // _percentuaisCurva = XLSXFile.XLSX2DictString_ListDouble(GetNomeArqConsumoMensalPU());
        }

        private string GetNomeArqConsumoMensalPU()
        {
            return _par._path + _par._permRes + "CurvasPercentuaisClientes4.xlsx";
        }

        //modelo
        //new load.3001215463M1 bus1=R9772.1.3.0,Phases=2,kv=0.22,kw=1.29794758726823,pf=0.92,Vminpu=0.92,Vmaxpu=1.5,model=2,daily=arqCurvaNormRES4_11,status=variable
        //new load.3001215463M2 bus1=R9772.1.3.0,Phases=2,kv=0.22,kw=1.29794758726823,pf=0.92,Vminpu=0.92,Vmaxpu=1.5,model=3,daily=arqCurvaNormRES4_11,status=variable
        // CodBase	CodConsBT	CodAlim	CodTrafo	CodRmlBT	CodFas	CodPonAcopl	SemRedAssoc	TipMedi	TipCrvaCarga	EnerMedid01_MWh	EnerMedid02_MWh	EnerMedid03_MWh	EnerMedid04_MWh	EnerMedid05_MWh	EnerMedid06_MWh	EnerMedid07_MWh	EnerMedid08_MWh	EnerMedid09_MWh	EnerMedid10_MWh	EnerMedid11_MWh	EnerMedid12_MWh	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	TnsLnh_kV	TnsFas_kV
        public bool ConsultaBanco()
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // se modo reconfiguracao 
                    {
                        command.CommandText = "select CodConsBT,TipCrvaCarga,(EnerMedid01_MWh + EnerMedid02_MWh + EnerMedid03_MWh + EnerMedid04_MWh + EnerMedid05_MWh + EnerMedid06_MWh +"
                            + "EnerMedid07_MWh + EnerMedid08_MWh + EnerMedid09_MWh + EnerMedid10_MWh + EnerMedid11_MWh + EnerMedid12_MWh)*1000/12 as MediakWh from " +
                            _par._DBschema + "StoredCargaBT " + "where CodBase=@codbase and TipCrvaCarga<>'IP' order by CodConsBT"; //OLD CODE and CodAlim=@CodAlim";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                        //command.Parameters.AddWithValue("@CodAlim", _par._alim);
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
                            string tipCurv = rs["TipCrvaCarga"].ToString();

                            // condicao de saida
                            if (tipCurv.Contains("IP"))
                            {
                                continue;
                            }

                            // calculates mean
                            double media_kWh = double.Parse(rs["MediakWh"].ToString());

                            // rnd obj
                            Random rndObj = new Random();
                            double numSort = rndObj.NextDouble();

                            string faixa;
                            string chave;
                            string tipo;
                            string curva;
                            string ucbt = rs["CodConsBT"].ToString();
                            string ucbt_curva = ucbt + ",";

                            //obtem classe 
                            if (tipCurv.Contains("RES"))
                            {
                                faixa = getFaixa(media_kWh, "RES");

                                chave = "RES_" + faixa;

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("COM"))
                            {
                                faixa = getFaixa(media_kWh, "COM");

                                chave = "COM_" + faixa;

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("RUR"))
                            {
                                faixa = getFaixa(media_kWh, "RUR");

                                chave = "RUR_" + faixa;

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("IND"))
                            {
                                faixa = getFaixa(media_kWh, "IND");

                                chave = "IND_" + faixa;

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("SERVP"))
                            {
                                faixa = getFaixa(media_kWh, "SERVP");

                                chave = "SERVP_" + faixa;

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("UGBT-B1"))
                            {
                                //faixa = getFaixa(media, "UGBT-B1");

                                chave = "UGBT-B1_FX1";

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("UGBT-B2"))
                            {
                                //faixa = getFaixa(media, "UGBT-B2");

                                chave = "UGBT-B2_FX1";

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            else if (tipCurv.Contains("UGBT-B3"))
                            {
                                //faixa = getFaixa(media, "UGBT-B3");

                                chave = "UGBT-B3_FX1";

                                tipo = getTipo(chave, numSort);

                                curva = chave + "_" + tipo;

                                ucbt_curva += curva + Environment.NewLine;
                            }
                            _arqSegmentoBT.Append(ucbt_curva);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        private string getTipo(string chave, double numSort)
        {
            List<double> lista;

            //
            if (_percentuaisCurva.ContainsKey(chave))
            {
                lista = _percentuaisCurva[chave];
                int k = 0;
                double perAcc = lista[k];

                // varre a lista de percentuais acumulados.
                while (perAcc < numSort)
                {
                    k++;

                    //atualiza da 
                    perAcc = lista[k];
                }
                return (k + 1).ToString();
            }
            return "1";
        }

        private string getFaixa(double consumo, string classe)
        {
            string faixa = "FX1";

            if (classe.Equals("RES"))
            {
                /*  < 100 MWh
                    101 a 220 MWh
                    221 a 350 MWh
                    350 a 500 MWh
                    501 a 1000 MWh
                    > 1000 MWh
                 * */
                if (consumo > 1000) { return faixa = "FX6"; }
                if (consumo > 501) { return faixa = "FX5"; }
                if (consumo > 350) { return faixa = "FX4"; }
                if (consumo > 221) { return faixa = "FX3"; }
                if (consumo > 101) { return faixa = "FX2"; }
            }
            if (classe.Equals("RUR"))
            {
                /*  < 300 MWh
                301 a 1000 MWh
                1001 a 5000 MWh
                > 5000 MWh */
                if (consumo > 5000) { return faixa = "FX4"; }
                if (consumo > 1001) { return faixa = "FX3"; }
                if (consumo > 301) { return faixa = "FX2"; }
            }
            if (classe.Equals("COM"))
            {
                /*  < 500 MWh
                501 a 2000 MWh
                2001 a 5000 MWh
                > 5000 MWh
                 */
                if (consumo > 5000) { return faixa = "FX4"; }
                if (consumo > 2001) { return faixa = "FX3"; }
                if (consumo > 501) { return faixa = "FX2"; }
            }
            if (classe.Equals("IND"))
            {
                /*  < 1000 MWh
                1001 a 3000 MWh
                3001 a 7000 MWh
                > 7000 MWh
                 */
                if (consumo > 7000) { return faixa = "FX4"; }
                if (consumo > 3001) { return faixa = "FX3"; }
                if (consumo > 1001) { return faixa = "FX2"; }
            }
            if (classe.Equals("SERVP"))
            {
                /*  < 2000 MWh
                2001 a 5000 MWh
                5001 a 10000 MWh
                > 10000 MWh
                 */
                if (consumo > 10000) { return faixa = "FX4"; }
                if (consumo > 5001) { return faixa = "FX3"; }
                if (consumo > 2001) { return faixa = "FX2"; }
            }
            if (classe.Equals("UGBT-B1"))
            {
                /* 
                 */
                if (consumo > 2001) { return faixa = "FX2"; }
            }
            if (classe.Equals("UGBT-B2"))
            {
                /* 
                 */
                if (consumo > 2001) { return faixa = "FX2"; }
            }
            if (classe.Equals("UGBT-B3"))
            {
                /* 
                 */
                if (consumo > 10000) { return faixa = "FX4"; }
            }

            return faixa;
        }
    }
}
