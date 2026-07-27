
using AuxClasses;
using Npgsql;

namespace ExportadorGeoPerdasDSS
{
    class Param
    {
        public string _DBschema; // 
        public string _path; // main path
        public string _permRes; // resources folder
        public int _codBase; // ANEEL company number
        public string _dist; // ANEEL company number
        public string _pathAlim; // feeder subdirectory
        public string _conjAlim; // string to concatenates more than one feeder separated by ','
        public string _trEM; //trecho energy meter
        public bool _modelo4condutores; // 4 model
        public string _alim;
        public string _ano;
        public PVSystemPar _pvMV;
        public PVSystemPar _pvLV;
        public Dictionary<string, string> _dicTipGer; // dic com tipo de geracao 

        public Param(string path, string permRes, int codBase, bool modelo4condutores, string schema,
             string alim, string ano, PVSystemPar pvMT, PVSystemPar pvBT, string dist, NpgsqlConnectionStringBuilder connBuilder)
        {
            if (path is null) throw new ArgumentNullException(nameof(path));
            if (permRes is null) throw new ArgumentNullException(nameof(permRes));
            if (schema is null) throw new ArgumentNullException(nameof(schema));
            if (alim is null) throw new ArgumentNullException(nameof(alim));
            if (ano is null) throw new ArgumentNullException(nameof(ano));
            if (dist is null) throw new ArgumentNullException(nameof(dist));
            if (pvMT is null) throw new ArgumentNullException(nameof(pvMT));
            if (pvBT is null) throw new ArgumentNullException(nameof(pvBT));
            if (connBuilder is null) throw new ArgumentNullException(nameof(connBuilder));

            _DBschema = schema;
            _permRes = permRes;
            _codBase = codBase;
            _path = path;
            _modelo4condutores = modelo4condutores;
            _alim = alim;
            _dist = dist;
            _ano = ano;
            _pvMV = pvMT;
            _pvLV = pvBT;
            _dicTipGer = GetTipoGeracaoDB(connBuilder);
            _conjAlim = string.Empty; 
            _trEM = string.Empty;
            _pathAlim = string.Empty;
        }

        public void SetCurrentAlim(string alim)
        {
            _alim = alim;
            _pathAlim = _path + alim + "\\";
        }


        //Get all feeders in a string separated by ',' from a substation name
        public bool GetAllFeedersFromSubstationString(string sub, NpgsqlConnectionStringBuilder con)
        {
            // OBS: a SE deve ter nome
            // obtem lstAlim da SE
            List<string> lstAlim = GetLstAlimSE(sub, con);


            // cria string com a uniao dos alimentadores 
            bool ret = UneStringAlim(lstAlim);

            return ret;
        }

        private List<string> GetLstAlimSE(string codSE, NpgsqlConnectionStringBuilder _connBuilder)
        {
            List<string> lstAlim = new List<string>();

            NpgsqlConnectionStringBuilder conn_old = new NpgsqlConnectionStringBuilder(_connBuilder.ToString());

            using (var conn = new NpgsqlConnection(_connBuilder.ToString()))
            {

                
                conn.Open();

                //consulta a banco 
                using (NpgsqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "SELECT CodAlim " +
                        "from " + _DBschema + "StoredCircMT " + 
                        "WHERE CodBase=@codbase AND CodSub=@codSe " +
                        "order by CodAlim";
                    command.Parameters.AddWithValue("@codbase", _codBase);
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
            }

            return lstAlim;
        }


        private bool UneStringAlim(List<string> lstAlim)
        {
            if (lstAlim.Count == 0)
            {
                return false;
            }

            string conjAlims = "'";

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

            // adds lst feeders in _par object
            _conjAlim = conjAlims;

            return true;
        }


        private Dictionary<string, string> GetTipoGeracaoDB(NpgsqlConnectionStringBuilder connBuilder)
        {
            Dictionary<string, string> _dicTipGer = new Dictionary<string, string>();

            using (var conn = new NpgsqlConnection(connBuilder.ToString()))
            {
                
                conn.Open();

                NpgsqlCommand command = conn.CreateCommand();

                // 
                command.CommandText = "select CodGera, CEG " +
                    "FROM " + _DBschema + "era_tipger "; // WHERE CodBase=@codbase AND CodAlim=@CodAlim";
                                                              //command.Parameters.AddWithValue("@codbase", _codBase);
                                                              //command.Parameters.AddWithValue("@CodAlim", _alim);               

                using (var rs = command.ExecuteReader())
                {
                    // verifica ocorrencia de elemento no banco
                    if (!rs.HasRows)
                    {
                        return _dicTipGer;
                    }

                    while (rs.Read())
                    {
                        var codGeraObj = rs["CodGera"];
                        var cegObj = rs["CEG"];

                        string codGera = codGeraObj == DBNull.Value || codGeraObj == null ? string.Empty : codGeraObj.ToString();
                        string ceg = cegObj == DBNull.Value || cegObj == null ? string.Empty : cegObj.ToString();

                        if (!string.IsNullOrEmpty(codGera) && !string.IsNullOrEmpty(ceg))
                        {
                            // sobrescreve caso já exista
                            _dicTipGer[codGera] = ceg;
                        }
                    }
                    return _dicTipGer;
                }
            }
        }
    }
}
