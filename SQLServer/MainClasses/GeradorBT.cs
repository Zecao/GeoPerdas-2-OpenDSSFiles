using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class GeradorBT
    {
        private static readonly string _geradorBT = "GeradorBT_";
        private static SqlConnectionStringBuilder _connBuilder;
        private int _iMes;
        private StringBuilder _arqGeradorBT;
        private readonly Param _par;

        public GeradorBT(SqlConnectionStringBuilder connBuilder, Param par, int iMes)
        {
            _iMes = iMes;
            _par = par;
            _connBuilder = connBuilder;
        }

        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqGeradorBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // 
                    command.CommandText = "select CodGeraBT,CodAlim,CodFas,CodPonAcopl,TnsLnh_kV," +
                        "EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh," +
                        "EnerMedid07_MWh,EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh,Descr,TipGer ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "from " + _par._DBschema + "StoredGeradorBT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);

                    }
                    else
                    {
                        command.CommandText += "from " + _par._DBschema + "StoredGeradorBT where CodBase=@codbase and CodAlim=@CodAlim";
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

                        _arqGeradorBT.Append(AuxFunc.GetLoadShape("UFV", "curvaUGBT"));

                        // has any PVSystem in feeder
                        bool hasAnyPVSystem = false;

                        while (rs.Read())
                        {
                            string fasesDSS = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string CodGeraBT = rs["CodGeraBT"].ToString();

                            // Obtem a geracao de acordo com o mes
                            string injecaoMes_MWh = AuxFunc.GetConsumoMesCorrente(rs, _iMes);

                            // se consumo = 0, nao gera string do gerador 
                            if (double.Parse(injecaoMes_MWh) == 0.0)
                                continue;

                            string tipGer = rs["TipGer"].ToString();
                            string tensaoLinha = rs["TnsLnh_kV"].ToString(); // TODO corrigir
                            string codFas = rs["CodFas"].ToString();

                            string phasesCom = AuxFunc.GetFasesDSS(codFas);
                            string numFases = AuxFunc.GetNumFases(codFas);
                            string tensaoKV = AuxFunc.Get_kVbase_BT(numFases);

                            // Gets k factor (transforms Energy into a power or demand value) 
                            string fatorKdiario = AuxFunc.GetFatorK(tipGer);

                            //Calculates mean Injected power
                            string meanInjectedPower_hour = AuxFunc.CalcDemandaPorFatorKdiario(injecaoMes_MWh, _iMes, _par._ano, fatorKdiario);
                            double meanInjectedPower_hour_d = double.Parse(meanInjectedPower_hour);

                            // flags if has any PVSystem
                            if ( tipGer.Equals("UFV") || tipGer.Equals("GD.") )
                            {
                                hasAnyPVSystem = true;
                            }

                            // TODO caso queir aindividualizar curvas
                            //defiir laodshape conforme tipo usina. 
                            // curva PU
                            //string linha = AuxFunc.GetLoadShape(tipGer,curvaUGBT);                                                         
                            string linha = "";

                            // PVSystem
                            if (_par._pvLV._geraInvControl && hasAnyPVSystem )
                            {
                                string kVA = (meanInjectedPower_hour_d * 1.2).ToString();

                                linha += "new PVSystem." + CodGeraBT
                                + " bus1=" + rs["CodPonAcopl"] + phasesCom //TODO "100"
                                + ",Phases=" + numFases
                                + ",kv=" + tensaoKV
                                + ",kVA=" + kVA
                                + ",Pmpp=" + meanInjectedPower_hour
                                + ",irradiance=1"
                                + ",%cutin=0.1,%cutout=0.1"
                                + ",VarFollowInverter=True" //modo diurno
                                + ",daily=" + "ccurvaUGBT" + Environment.NewLine;
                            }
                            else
                            {
                                linha += "new generator." + CodGeraBT
                                + " bus1=" + rs["CodPonAcopl"] + phasesCom //TODO 
                                + ",Phases=" + numFases
                                + ",kv=" + tensaoKV
                                + ",kW=" + meanInjectedPower_hour
                                + ",pf=1"
                                + ",model=1"
                                + ",daily=" + "ccurvaUGBT"
                                + ",status=Variable" + Environment.NewLine;
                            }

                            _arqGeradorBT.Append(linha);
                        }

                        //_par._pvLV._invControlMode.Equals("VOLTVAR") NAO USADO
                        // adds InvControl if invControlMode = voltvar and has any PVSystem
                        if (_par._pvLV._geraInvControl && hasAnyPVSystem)
                        {
                            string linha2 = "New InvControl.InvPVCtrl"
                                + " mode=VOLTVAR"
                                + ",voltage_curvex_ref=rated"
                                + ",vvc_curve1=voltvar_c"
                                + ",VoltageChangeTolerance=0.001,VarChangeTolerance=0.05"
                                  + Environment.NewLine;
                            /*  
                                !VoltageChangeTolerance=0.0001 !default
                                !VarChangeTolerance=0.025 !default
                                !RateofChangeMode=RISEFALL
                                !RiseFallLimit=1 VARMAX
                                VarFollowInverter: Boolean variable which indicates that the reactive power does not respect
                                the inverter status.
                                – When set to True, PVSystem’s reactive power will cease when the inverter status is OFF,
                                due to the power from PV array dropping below %cutout. The reactive power will begin
                                again when the power from PV array is above %cutin;
                                – When set to False, PVSystem will provide/absorb reactive power regardless of the status
                                of the inverter.
                            */
                            linha2 += "!,VoltageChangeTolerance=0.0001,VarChangeTolerance=0.025,EventLog=yes,VarFollowInverter=True";
                            _arqGeradorBT.Append(linha2);
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
            string strMes = AuxFunc.IntMes2strMes(_iMes);

            return _par._pathAlim + _par._alim + _geradorBT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqGeradorBT.ToString(), GetNomeArq());
        }
    }
}
