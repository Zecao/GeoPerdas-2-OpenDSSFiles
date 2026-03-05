namespace ExportadorGeoPerdasDSS
{
    // opcoes de modelos SDEE 
    class ModeloSDEE
    {
        public bool _usarCondutoresSeqZero = false;
        public bool _utilizarCurvaDeCargaClienteMTIndividual = true;
        public bool _incluirCapacitoresMT = false;
        public bool _reatanciaTrafos = false;
        public string _modeloCarga = "ANEEL"; // Modelos disponiveis: "ANEEL" "PCONST"
        //public bool _incluirReatanciaDispersaoTrafos = true;

        public ModeloSDEE(bool usarCondutoresSeqZero, bool utilizarCurvaDeCargaClienteMTIndividual, bool incluirCapacitoresMT, string modeloCarga, bool reatanciaTrafos)
        {
            _usarCondutoresSeqZero = usarCondutoresSeqZero;
            _utilizarCurvaDeCargaClienteMTIndividual = utilizarCurvaDeCargaClienteMTIndividual;
            _incluirCapacitoresMT = incluirCapacitoresMT;
            _modeloCarga = modeloCarga;
            _reatanciaTrafos = reatanciaTrafos;
        }
    }
}
