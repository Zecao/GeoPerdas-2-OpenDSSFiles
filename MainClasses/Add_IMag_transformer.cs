using ExportadorGeoPerdasDSS;
using System;
using System.IO;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class Add_IMag_transformer
    {
        private readonly Param _par;
        private readonly StringBuilder _arqTrafoIMag = new StringBuilder();

        public Add_IMag_transformer(Param parGer)
        {
            _par = parGer;

            string nomeArq = GetNomeArq();

            if (! File.Exists(nomeArq) )
            {
                return;
            }

            FileStream arq = File.OpenRead(nomeArq);

            var enumLines = File.ReadLines(nomeArq, Encoding.UTF8);

            foreach (var line in enumLines)
            {
                string[] strTrafo = line.Split(',');

                string kvs="";
                string kva="";                


                // obtem kvs e kva 
                foreach (string par in strTrafo)
                {
                    if (par.Contains("kvs=["))
                    {
                        string[] split2 = par.Split(' ');

                        kvs = split2[0];
                        int found = kvs.IndexOf("kvs=[");
                        kvs = kvs.Substring(5);              
                    }
                    if (par.Contains("kvas=["))
                    {
                        string[] split2 = par.Split(' ');

                        kva = split2[0];
                        int found = kva.IndexOf("kvas=[");
                        kva = kva.Substring(6);
                    }
                }

                // get iMag
                double iMag = Get_iMag(double.Parse(kvs), double.Parse(kva));

                // cria string iMag
                string iMag_s = "";
                if (iMag != 0.0)
                {
                    // MAg
                    iMag_s = ",%imag=" + iMag.ToString();
                }

                // separa comando X comentario
                string[] split3 = line.Split('!');
                // 
                if (split3.Length > 1)
                {
                    // novo Trafo.
                    _arqTrafoIMag.Append(split3[0] + iMag_s + " !" + split3[1] + Environment.NewLine);
                }
                else // PT nao tem comentario comecanco com ! 
                {
                    // novo Trafo.
                    _arqTrafoIMag.Append(split3[0] + iMag_s + Environment.NewLine);
                }
            }

            // fecha stream
            arq.Close();

            // 
            GravaEmArquivo();
        }

        private string GetNomeArq()
        {
            return _par._pathAlim + _par._alim + "Transformadores.dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqTrafoIMag.ToString(), GetNomeArq());
        }

        // 
        private double Get_iMag(double kv, double kva)
        {
            double iMag = 0.0;

            switch (kv)
            {
                case 7.967:

                    switch (kva)
                    {
                        case 5:
                            iMag = 3.35;
                            break;
                        case 10:
                            iMag = 2.66;
                            break;
                        case 15:
                            iMag = 2.37;
                            break;
                        case 25:
                            iMag = 2.18;
                            break;
                        case 37.5:
                            iMag = 2.08;
                            break;
                        case 50:
                            iMag = 1.98;
                            break;
                        case 75:
                            iMag = 1.88;
                            break;
                        case 100:
                            iMag = 1.79;
                            break;
                        case 835:
                            iMag = 1.69;
                            break;
                        case 1670:
                            iMag = 1.70;
                            break;
                    }
                    break;

                case 12.7:

                    switch (kva)
                    {
                        case 5:
                            iMag = 3.73;
                            break;
                        case 10:
                            iMag = 3.26;
                            break;
                        case 15:
                            iMag = 2.96;
                            break;
                        case 25:
                            iMag = 2.78;
                            break;
                        case 37.5:
                            iMag = 2.68;
                            break;
                        case 50:
                            iMag = 2.58;
                            break;
                        case 75:
                            iMag = 1.98;
                            break;
                        case 100:
                            iMag = 1.38;
                            break;
                    }
                    break;

                case 19.98:

                    switch (kva)
                    {
                        case 5:
                            iMag = 4.02;
                            break;
                        case 10:
                            iMag = 3.46;
                            break;
                        case 15:
                            iMag = 3.16;
                            break;
                        case 25:
                            iMag = 2.98;
                            break;
                        case 37.5:
                            iMag = 2.78;
                            break;
                        case 50:
                            iMag = 2.57;
                            break;
                        case 75:
                            iMag = 1.98;
                            break;
                        case 100:
                            iMag = 1.38;
                            break;
                        case 835:
                            iMag = 1.69;
                            break;
                        case 1670:
                            iMag = 1.70;
                            break;
                    }
                    break;

                case 22.0:

                    switch (kva)
                    {
                        case 15:
                            iMag = 3.96;
                            break;
                        case 30:
                            iMag = 3.57;
                            break;
                        case 45:
                            iMag = 3.17;
                            break;
                        case 75:
                            iMag = 2.68;
                            break;
                        case 112.5:
                            iMag = 2.48;
                            break;
                        case 150:
                            iMag = 2.28;
                            break;
                        case 225:
                            iMag = 2.08;
                            break;
                        case 300:
                            iMag = 1.88;
                            break;
                        case 500:
                            iMag = 1.78;
                            break;
                        case 750:
                            iMag = 1.68;
                            break;
                        case 1000:
                            iMag = 1.38;
                            break;
                            /*default:
                                iMag = 3.0;*/
                    }
                    break;

                case 34:
                    switch (kva)
                    {
                        case 15:
                            iMag = 3.96;
                            break;
                        case 30:
                            iMag = 3.57;
                            break;
                        case 45:
                            iMag = 3.17;
                            break;
                        case 75:
                            iMag = 2.68;
                            break;
                        case 112.5:
                            iMag = 2.48;
                            break;
                        case 150:
                            iMag = 2.28;
                            break;
                        case 225:
                            iMag = 2.08;
                            break;
                        case 300:
                            iMag = 1.88;
                            break;
                        case 500:
                            iMag = 1.78;
                            break;
                        case 750:
                            iMag = 1.68;
                            break;
                        case 1000:
                            iMag = 1.38;
                            break;
                            /*default:
                                iMag = 3.0;*/
                    }
                    break;

                default: // 13.8

                    switch (kva)
                    {
                        case 15:
                            iMag = 3.97;
                            break;
                        case 30:
                            iMag = 3.57;
                            break;
                        case 45:
                            iMag = 3.18;
                            break;
                        case 75:
                            iMag = 2.68;
                            break;
                        case 112.5:
                            iMag = 2.48;
                            break;
                        case 150:
                            iMag = 2.28;
                            break;
                        case 225:
                            iMag = 2.09;
                            break;
                        case 300:
                            iMag = 1.89;
                            break;
                        case 500:
                            iMag = 1.58;
                            break;
                        case 750:
                            iMag = 1.28;
                            break;
                        case 1000:
                            iMag = 1.18;
                            break;
                        case 2500:
                            iMag = 1.69;
                            break;
                        case 5000:
                            iMag = 1.70;
                            break;
                    }
                    break;
            }

            return iMag;
        }


    }
}
