

namespace AuxClasses
{
    class PVSystemPar
    {
        public bool _modelPVSystems; // false = exports PVSystems as generator model=1
        public bool _geraInvControl;
        public string _invControlMode;
        public string _voltVarcurve;
        public string _varFollowInv; // set False to night mode (and if VOLTVAR)

        /*
        _modelPVSystems     -> False = exports PVSystems as generator model = 1
        _invControlModeMV   -> "PF=1" "VOLTVAR" "voltwatt"
        _varFollowInvMV     -> False to night mode(and if VOLTVAR)
        _voltVarcurve =     -> "voltvar_c";  ///uses voltvar_0 to no voltvar
        */

        /* VarFollowInverter: Boolean variable which indicates that the reactive power does not respect the inverter status.
        – When set to True, PVSystem’s reactive power will cease when the inverter status is OFF,
        due to the power from PV array dropping below %cutout. The reactive power will begin
        again when the power from PV array is above %cutin;
        – When set to False, PVSystem will provide/absorb reactive power regardless of the status
        of the inverter.*/

        public PVSystemPar(bool modelPVSystems, bool gerIC, string invControlModeLV = "VOLTVAR", string varFollowInvLV = "False", string vvarCurve = "voltvar_c")
        {
            // se InvControl was choose, modelPVSystem must be true 
            if (gerIC)
            {
                modelPVSystems = true;
            }

            _modelPVSystems = modelPVSystems;
            _geraInvControl = gerIC;

            // Optional parameters 
            _invControlMode = invControlModeLV;
            _varFollowInv = varFollowInvLV;
            _voltVarcurve = vvarCurve;
        }

        public bool GeraInvControl()
        {
            return _geraInvControl;
            /*
            if ( _modelPVSystems) //&& ! _invControlMode.Equals("")
                return true;
            else
                return false;
                */
        }
    }
}
