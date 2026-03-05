
using AuxClasses;

namespace ExportadorGeoPerdasDSS
{
    class Param
    {
        public string _DBschema; // database
        public string _path; // main path
        public string _permRes; // resources folder
        public string _codBase; // ANEEL company number
        public string _dist; // ANEEL company number
        public string _pathAlim; // feeder subdirectory
        public string _conjAlim; // string to concatenates more than one feeder separated by ','
        public string _trEM; //trecho energy meter
        public bool _modelo4condutores; // 4 model
        public string _alim;
        public string _ano;
        public PVSystemPar _pvMV;
        public PVSystemPar _pvLV;

        public Param(string path, string permRes, string codBase, bool modelo4condutores, string schema,
             string alim, string ano, PVSystemPar pvMT, PVSystemPar pvBT, string dist)
        {
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
        }

        public void SetCurrentAlim(string alim)
        {
            _alim = alim;
            _pathAlim = _path + alim + "\\";
        }
    }
}
