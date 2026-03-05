using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class GeradorMT
    {
        // membros privador
        private static readonly string _geradorMT = "GeradorMT_";
        private static SqlConnectionStringBuilder _connBuilder;
        private int _iMes;
        private StringBuilder _arqGeradorMT;
        private readonly Param _par;

        public GeradorMT(SqlConnectionStringBuilder connBuilder, Param par, int iMes)
        {
            _iMes = iMes;
            _par = par;
            _connBuilder = connBuilder;
        }

        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqGeradorMT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // 
                    command.CommandText = "select CodGeraMT,CodAlim,CodFas,CodPonAcopl,TnsLnh_kV," +
                        "EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh," +
                        "EnerMedid07_MWh,EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh,Descr,TipGer ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "from " + _par._DBschema + "StoredGeradorMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);

                    }
                    else
                    {
                        command.CommandText += "from " + _par._DBschema + "StoredGeradorMT where CodBase=@codbase and CodAlim=@CodAlim";
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

                        // has any PVSystem in feeder
                        bool hasAnyPVSystem = false;

                        while (rs.Read())
                        {
                            string fasesDSS = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string CodGeraMT = rs["CodGeraMT"].ToString();

                            // Obtem a geracao de acordo com o mes
                            string injecaoMes_MWh = AuxFunc.GetConsumoMesCorrente(rs, _iMes);

                            // se consumo = 0, nao gera string do gerador 
                            if (double.Parse(injecaoMes_MWh) == 0.0)
                                continue;

                            string tipGer = rs["TipGer"].ToString();

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

                            // sets loadshape by gnerator type 
                            string linha = AuxFunc.GetLoadShape(tipGer, CodGeraMT);

                            // PVSystem
                            if (_par._pvMV._geraInvControl && hasAnyPVSystem)
                            {
                                string kVA = (meanInjectedPower_hour_d * 1.2).ToString();

                                linha += "new PVSystem." + CodGeraMT
                                + " bus1=" + "BMT" + rs["CodPonAcopl"] + ".1.2.3"
                                + ",Phases=3"
                                + ",kv=" + rs["TnsLnh_kV"].ToString()
                                + ",kVA=" + kVA
                                + ",Pmpp=" + meanInjectedPower_hour
                                + ",irradiance=1"
                                + ",%cutin=0.1,%cutout=0.1"
                                + ",VarFollowInverter=True" //modo diurno
                                + ",daily=c" + CodGeraMT + Environment.NewLine;
                            }
                            else
                            {
                                linha += "new generator." + CodGeraMT
                                + " bus1=" + "BMT" + rs["CodPonAcopl"] + ".1.2.3"
                                + ",Phases=3"
                                + ",kv=" + rs["TnsLnh_kV"].ToString()
                                + ",kW=" + meanInjectedPower_hour
                                + ",pf=1"
                                + ",model=1"
                                + ",daily=c" + CodGeraMT
                                + ",status=Variable" + Environment.NewLine;
                            }

                            _arqGeradorMT.Append(linha);
                        }

                        // adds InvControl if invControlMode = voltvar and has any PVSystem
                        if (_par._pvMV._geraInvControl && _par._pvMV._invControlMode.Equals("VOLTVAR") && hasAnyPVSystem)
                        {
                            string linha = "New InvControl.InvPVCtrl"
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
                            linha += "!,VoltageChangeTolerance=0.0001,VarChangeTolerance=0.025,EventLog=yes,VarFollowInverter=True";
                            _arqGeradorMT.Append(linha);
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

            return _par._pathAlim + _par._alim + _geradorMT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqGeradorMT.ToString(), GetNomeArq());
        }
    }
}
