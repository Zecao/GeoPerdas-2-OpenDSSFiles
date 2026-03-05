using System;
using System.Data.SqlClient;
using System.Text;
using System.Text.RegularExpressions;
using Npgsql;

namespace ExportadorGeoPerdasDSS
{
    class ChaveMT
    {
        // membros privados
        private static readonly string _chave = "ChavesMT.dss";
        private static NpgsqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqChaveMT;
        private readonly Param _par;
        private readonly bool _criaDispProtecao;

        public ChaveMT(NpgsqlConnectionStringBuilder connBuilder, Param par, bool criaDispProtecao)
        {
            _par = par;
            _connBuilder = connBuilder;
            _criaDispProtecao = criaDispProtecao;
        }

        // new line.CTR100934 bus1=98417853.2,bus2=98417888.2,Phases=1,LineCode=tieSwitch1,Length=0.001,Units=km,switch=T
        // open line.CTR522670 term= 1
        // CodBase	CodChv	CodAlim	CodPonAcopl1	CodPonAcopl2	CodFas	EstChv	Descr	CodSubAtrib	CodAlimAtrib	Ordm	De	Para
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqChaveMT = new StringBuilder();

            using (NpgsqlConnection conn = new NpgsqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    /* TODO 
                    command.CommandText = "select CodChvMT,CodPonAcopl1,CodPonAcopl2,CodFas,EstChv,Descr,TipoDisp,EloFus "
                            + "from dbo.StoredChaveMT ";*/

                    command.CommandText = "select CodChvMT,CodPonAcopl1,CodPonAcopl2,CodFas,EstChv,Descr "
                     + "from " + _par._DBschema + "StoredChaveMT ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string fasesDSS = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string lineCode = "tieSwitch";
                            string codChave = rs["CodChvMT"].ToString();

                            //se chave mono
                            if (numFases.Equals("1"))
                            {
                                lineCode = "tieSwitch1";
                            }

                            string linha = "new line.CTR" + codChave //OBS1 adicionei prefixo CTR
                                + " bus1=" + rs["CodPonAcopl1"].ToString() + fasesDSS  //OBS1
                                + ",bus2=" + rs["CodPonAcopl2"].ToString() + fasesDSS  //OBS1
                                + ",Phases=" + numFases
                                + ",LineCode=" + lineCode
                                + ",Length=0.001,Units=km,switch=T"
                                + " !numEQ " + rs["Descr"].ToString() + Environment.NewLine;

                            // se chave fechada 
                            if (rs["EstChv"].ToString().Equals("1"))
                            {
                                linha += "open line.CTR" + codChave + " term=1" + Environment.NewLine; //OBS1 adicionei prefixo CTR
                            }

                            // creates protection devices (Recloser, Fuses)
                            if (_criaDispProtecao)
                            {
                                string tipoDisp = rs["TipoDisp"].ToString();
                                string eloFus = rs["EloFus"].ToString();
                                linha = CreateStrDispProtection(codChave, tipoDisp, eloFus, linha);
                            }

                            _arqChaveMT.Append(linha);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        // creates protection devices (Recloser, Fuses)
        private string CreateStrDispProtection(string codChave, string tipoDisp, string eloFus, string linha)
        {
            //
            if ((codChave.Contains("R")) || tipoDisp.Equals("R"))
            {
                linha += "new Recloser." + codChave + " monitoredobj='line.CTR" + codChave + "' monitoredterm=1,numfast=1,"
                    + "phasedelayed='VERY_INV',grounddelayed='VERY_INV',phasetrip=280,TDPhDelayed=1,GroundTrip=45,TDGrDelayed=14,"
                    + "shots=1,recloseintervals=(10, 20, 20)" + Environment.NewLine;
            }
            else if (tipoDisp.Equals("F"))
            {

                Regex rgxNumero = new Regex(@"[1-9]\d*");
                Regex rgxTexto = new Regex(@"\p{L}+");

                int corrente = int.Parse(rgxNumero.Match(eloFus).Value);
                string curva = rgxTexto.Match(eloFus).Value;

                if (curva.Equals(""))
                {
                    curva = "T";
                }

                /*
                int iCor = int.Parse(corrente); // remove zeros a esquerda ex: 040;s
                char curva = eloFus.Last();*/

                linha += "new Fuse." + codChave + " monitoredobj='line.CTR" + codChave + "',monitoredterm=1,"
                 + "Fusecurve='" + curva + "link',ratedcurrent=" + corrente.ToString() + Environment.NewLine;
            }
            return linha;
        }

        private string GetNomeArq()
        {
            return _par._pathAlim + _par._alim + _chave;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqChaveMT.ToString(), GetNomeArq());
        }
    }
}
