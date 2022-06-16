using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Keithley.Ke26XXA.Interop;
using System.Globalization;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Diagnostics;
using System.Linq;

namespace Transistor
{
    public partial class Form1 : Form
    {

        private IKe26XXA Kdriver;
        string KeithleyID;

        bool Keithley_connected = false, isSaved = true, nextVg = false, isLifetime = false, Cdrain = true, Cgate = true, overflow = false, isTransfer = false, backup=false;

        double GateCur, DrainCur, GateVolt, DrainVolt, time=0, setVg, setVd, Rdrain, Rgate;
        int Step, saveNumber=10, nVals = 0;
        double[] Level, Duration, Delay;

        double[] range = { 100e-9, 1e-6, 10e-6, 100e-6, 1e-3, 10e-3, 100e-3, 1, 1.5 };

        List<double> IdsList = new List<double>();
        List<double> IgsList = new List<double>();
        List<double> VdList = new List<double>();
        List<double> VgList = new List<double>();
        List<double> TimeList = new List<double>();
        List<string> TimeIRLList = new List<string>();

        // Prepare chart series
        Series IdsT = new Series("IdsT");
        Series IgsT = new Series("IgsT");
        Series VdT = new Series("VdT");
        Series VgT = new Series("VgT");

        Stopwatch stopwatch = new Stopwatch();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txt_KeithleyID.Text = Properties.Settings.Default.KeithleyID;

            Nud_drain_start.Value = Properties.Settings.Default.Drain_Start_O;
            Nud_drain_step.Value = Properties.Settings.Default.Drain_Step_O;
            Nud_drain_end.Value = Properties.Settings.Default.Drain_End_O;
            cb_drain_range.SelectedIndex = Properties.Settings.Default.Drain_Range_O;
            chk_drain_sweep.Checked = Properties.Settings.Default.Drain_Sweep_O;

            Nud_gate_start.Value = Properties.Settings.Default.Gate_Start_O;
            Nud_gate_step.Value = Properties.Settings.Default.Gate_Step_O;
            Nud_gate_end.Value = Properties.Settings.Default.Gate_End_O;
            cb_gate_range.SelectedIndex = Properties.Settings.Default.Gate_Range_O;
            chk_gate_sweep.Checked = Properties.Settings.Default.Gate_Sweep_O;

            Nud_delay.Value = Properties.Settings.Default.Delay_O;
            txt_name.Text = Properties.Settings.Default.Name;
            txt_path.Text = Properties.Settings.Default.Path;
            chk_autosave.Checked = Properties.Settings.Default.Autosave;

            Nud_time_drainSet.Value = Properties.Settings.Default.Time_Drain_Voltage;
            cb_time_drain_range.SelectedIndex = Properties.Settings.Default.Time_Drain_Range;
            Nud_time_gateSet.Value = Properties.Settings.Default.Time_Gate_Voltage;
            cb_time_gate_range.SelectedIndex = Properties.Settings.Default.Time_Gate_Range;
            Nud_time_interval.Value = Properties.Settings.Default.Time_Interval;
            Nud_time_duration.Value = Properties.Settings.Default.Time_Duration;

            Checkboxes();

            if (txt_path.Text == "" || txt_path.Text == " ")
            {
                txt_path.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
        }

        private void Btt_path_Click(object sender, EventArgs e)
        {
            fbd_path.SelectedPath = txt_path.Text;
            if (fbd_path.ShowDialog() == DialogResult.OK)
            {
                txt_path.Text = fbd_path.SelectedPath;
            }
        }


        private void PrepareMeasurement()
        {
            SendCommand("smua.reset()");
            SendCommand("display.screen = display.SMUA");
            SendCommand("format.data = format.ASCII");
            SendCommand("smua.nvbuffer1.clear()");
            SendCommand("display.smua.measure.func = display.MEASURE_DCAMPS");
            SendCommand("smua.source.levelv = 0");

            SendCommand("smub.reset()");
            SendCommand("display.screen = display.SMUA_SMUB");
            SendCommand("smub.nvbuffer1.clear()");
            SendCommand("display.smub.measure.func = display.MEASURE_DCAMPS");
            SendCommand("smub.source.levelv = 0");

            


            // Set Measurement Range
            if (cb_drain_range.SelectedIndex == 0)
            {
                SendCommand("smua.measure.autorangei = smua.AUTORANGE_ON");
            }
            else
            {
                Rdrain = range[cb_drain_range.SelectedIndex - 1];
                SendCommand(String.Format("smua.measure.rangei = {0}", range[cb_drain_range.SelectedIndex - 1]));
            }

            if (cb_gate_range.SelectedIndex == 0)
            {
                SendCommand("smub.measure.autorangei = smub.AUTORANGE_ON");
            }
            else
            {
                Rgate = range[cb_gate_range.SelectedIndex - 1];
                SendCommand(String.Format("smub.measure.rangei = {0}", range[cb_drain_range.SelectedIndex - 1]));
            }

            if (!isSaved)
            {
                DialogResult res = MessageBox.Show("Last measurement not saved. Save old data?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                switch (res)
                {
                    case DialogResult.Cancel:
                        return;
                    case DialogResult.Yes:
                        SaveFile(false);
                        break;
                    case DialogResult.No:
                        isSaved = true;
                        break;
                }
            }

            btt_cancel.Enabled = true;
            btt_start.Enabled = false;

            ClearData();

            isSaved = false;

            // Prepare chart series
            IdsT.IsVisibleInLegend = false;
            IdsT.ChartType = SeriesChartType.Line;
            IdsT.MarkerStyle = MarkerStyle.Circle;
            IdsT.Color = Color.Blue;
            IgsT.IsVisibleInLegend = false;
            IgsT.ChartType = SeriesChartType.Line;
            IgsT.MarkerStyle = MarkerStyle.Circle;
            IgsT.Color = Color.Blue;

            //Add Series to graphs
            ct_DS.Series.Clear();
            ct_GS.Series.Clear();
            ct_DS.Series.Add(IdsT);
            ct_GS.Series.Add(IgsT);

            pB_progress.Maximum = 100;
            pB_progress.Value = 0;
            pB_progress.Style = ProgressBarStyle.Blocks;
        }

        private void PrepareLifetime()
        {
            SendCommand("smua.reset()");
            SendCommand("display.screen = display.SMUA");
            SendCommand("format.data = format.ASCII");
            SendCommand("smua.nvbuffer1.clear()");
            SendCommand("display.smua.measure.func = display.MEASURE_DCAMPS");
            SendCommand("smua.source.levelv = 0");

            SendCommand("smub.reset()");
            SendCommand("display.screen = display.SMUA_SMUB");
            SendCommand("smub.nvbuffer1.clear()");
            SendCommand("display.smub.measure.func = display.MEASURE_DCAMPS");
            SendCommand("smub.source.levelv = 0");


            // Set Measurement Range
            if (cb_time_drain_range.SelectedIndex == 0)
            {
                SendCommand("smua.measure.autorangei = smua.AUTORANGE_ON");
            }
            else
            {
                Rdrain = range[cb_time_drain_range.SelectedIndex - 1];
                SendCommand(String.Format("smua.measure.rangei = {0}", range[cb_time_drain_range.SelectedIndex - 1]));
            }

            if (cb_time_gate_range.SelectedIndex == 0)
            {
                SendCommand("smub.measure.autorangei = smub.AUTORANGE_ON");
            }
            else
            {
                Rgate = range[cb_time_gate_range.SelectedIndex - 1];
                SendCommand(String.Format("smub.measure.rangei = {0}", range[cb_time_drain_range.SelectedIndex - 1]));
            }

            if (!isSaved)
            {
                DialogResult res = MessageBox.Show("Last measurement not saved. Save old data?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                switch (res)
                {
                    case DialogResult.Cancel:
                        return;
                    case DialogResult.Yes:
                        SaveFile(false);
                        break;
                    case DialogResult.No:
                        isSaved = true;
                        break;
                }
            }

            foreach (ListViewItem itm in LV_sweep.Items)
            {
                itm.BackColor = txt_KeithleyID.BackColor;
            }

            btt_cancel.Enabled = true;
            btt_start.Enabled = false;
            gb_display.Enabled = false;

            ClearData();

            isSaved = false;

            // Prepare chart series
            VdT.IsVisibleInLegend = false;
            VdT.ChartType = SeriesChartType.Line;
            VdT.MarkerStyle = MarkerStyle.Circle;
            VdT.Color = Color.Blue;
            VgT.IsVisibleInLegend = false;
            VgT.ChartType = SeriesChartType.Line;
            VgT.MarkerStyle = MarkerStyle.Circle;
            VgT.Color = Color.Blue;

            IdsT.IsVisibleInLegend = false;
            IdsT.ChartType = SeriesChartType.Line;
            IdsT.MarkerStyle = MarkerStyle.Circle;
            IdsT.Color = Color.Blue;
            IgsT.IsVisibleInLegend = false;
            IgsT.ChartType = SeriesChartType.Line;
            IgsT.MarkerStyle = MarkerStyle.Circle;
            IgsT.Color = Color.Blue;

            //Add Series to graphs
            ct_time_DS.Series.Clear();
            ct_time_GS.Series.Clear();

            if (Cdrain)
            {
                ct_time_DS.Series.Add(IdsT);
            }
            else
            {
                ct_time_DS.Series.Add(VdT);
            }

            if (Cgate)
            {
                ct_time_GS.Series.Add(IgsT);
            }
            else
            {
                ct_time_GS.Series.Add(VgT);
            }

            if (Nud_time_duration.Value > 3)
            {
                ct_time_DS.ChartAreas[0].AxisX.Title = "Time (min)";
                ct_time_GS.ChartAreas[0].AxisX.Title = "Time (min)";
            }
            else
            {
                ct_time_DS.ChartAreas[0].AxisX.Title = "Time (s)";
                ct_time_GS.ChartAreas[0].AxisX.Title = "Time (s)";
            }

            pB_progress.Maximum = 100;
            pB_progress.Value = 0;
            pB_progress.Style = ProgressBarStyle.Marquee;
            pB_progress.MarqueeAnimationSpeed = 100;
        }

        /// <summary>
        /// Send a text command to the Keithley
        /// </summary>
        /// <param name="cmd">Command to send</param>
        private void SendCommand(string cmd)
        {
            // use . as decimalSeperator
            NumberFormatInfo info = (NumberFormatInfo)NumberFormatInfo.CurrentInfo.Clone();
            info.NumberDecimalSeparator = ".";
            CultureInfo culture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            culture.NumberFormat = info;
            Thread.CurrentThread.CurrentCulture = culture;

            if (!Keithley_connected)
            {
                MessageBox.Show("Keithley not Connected", "Ke26XX Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cmd.Trim() == "")
            {
                return;
            }

            try
            {
                Kdriver.System.DirectIO.WriteString(cmd, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:  " + ex.Message, "Ke26XX Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// Set the voltage on the Keithley
        /// </summary>
        /// <param name="V">Voltage level</param>
        private void SetVoltage(double Vd, double Vg)
        {
            try
            {
                SendCommand("smua.source.levelv = " + Vd.ToString());
                SendCommand("smub.source.levelv = " + Vg.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:  " + ex.Message, "Ke26XX Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// Set the gate Voltage
        /// </summary>
        /// <param name="Vgate">Voltage level for gate</param>
        private void SetVoltageG(double Vgate)
        {
            try
            {
                SendCommand("smub.source.levelv = " + Vgate.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:  " + ex.Message, "Ke26XX Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// Set the drain Voltage
        /// </summary>
        /// <param name="Vgate">Voltage level for drain</param>
        private void SetVoltageD(double Vdrain)
        {
            try
            {
                SendCommand("smua.source.levelv = " + Vdrain.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:  " + ex.Message, "Ke26XX Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }



        private void MeasureIV()
        {
            double[] result = { -1, -1 };
            try
            {
                Thread.Sleep(10);
                SendCommand("smua.measure.i(smua.nvbuffer1)");
                Thread.Sleep(10);
                SendCommand("printbuffer(1, 1, smua.nvbuffer1)");
                //read results
                DrainCur = Convert.ToDouble(Kdriver.System.DirectIO.ReadString());
                if (DrainCur > Rdrain) { throw new OverflowException("Drain-Source current out of measurement range"); }
                Thread.Sleep(10);

                SendCommand("smub.measure.i(smub.nvbuffer1)");
                Thread.Sleep(10);
                SendCommand("printbuffer(1, 1, smub.nvbuffer1)");
                //read results
                GateCur = Convert.ToDouble(Kdriver.System.DirectIO.ReadString());
                if (GateCur > Rgate) { throw new OverflowException("Gate-Source current out of measurement range"); }
                Thread.Sleep(10);

                SendCommand("smua.measure.v(smua.nvbuffer1)");
                Thread.Sleep(10);
                SendCommand("printbuffer(1, 1, smua.nvbuffer1)");
                //read results
                DrainVolt = Convert.ToDouble(Kdriver.System.DirectIO.ReadString());

                SendCommand("smub.measure.v(smub.nvbuffer1)");
                Thread.Sleep(10);
                SendCommand("printbuffer(1, 1, smub.nvbuffer1)");
                //read results
                GateVolt = Convert.ToDouble(Kdriver.System.DirectIO.ReadString());

            }
            catch (OverflowException ovEx)
            {
                SendCommand("smua.source.output=0");
                SendCommand("smub.source.output=0");

                overflow = true;

                MessageBox.Show(ovEx.Message, "Overflow Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (isLifetime)
                {
                    bgW_LT.CancelAsync();
                }
                else
                {
                    bgW.CancelAsync();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:  " + ex.Message, "Ke26XX Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ;
            }
        }



        private void Btt_start_Click(object sender, EventArgs e)
        {
            if (!Keithley_connected)
            {
                MessageBox.Show("Keithley not connected!", "Connection error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tc_mode.SelectedIndex < 2)
            {
                PrepareMeasurement();
                isLifetime = false;
                bgW.RunWorkerAsync();
            }
            else
            {
                PrepareLifetime();
                isLifetime = true;
                if (Chk_Pulse.Checked)
                {
                    ConvertData();
                    bgWPulse.RunWorkerAsync();
                }
                else
                {
                    bgW_LT.RunWorkerAsync();
                }
                
            }
            
        }

        private void Btt_cancel_Click(object sender, EventArgs e)
        {
            if (isLifetime)
            {
                if (Chk_Pulse.Checked)
                {
                    bgWPulse.CancelAsync();
                }
                bgW_LT.CancelAsync();
            }
            else
            {
                bgW.CancelAsync();
            }
        }

        private void Btt_connect_Click_1(object sender, EventArgs e)
        {
            if (Keithley_connected)
            {
                if (Kdriver.Initialized)
                {
                    Kdriver.Close();
                    Keithley_connected = false;
                    lbl_Connection.Text = "";
                    btt_connect.Text = "Connect";
                    return;
                }
            }
            try
            {
                Kdriver = new Ke26XXA();
                KeithleyID = String.Format("USB0::0x05E6::0x2612::{0}::INSTR",txt_KeithleyID.Text.Trim());

                Kdriver.Initialize(KeithleyID, true, true, "QueryInstrStatus=True");

                // Clear all of the applicable stimulus settings
                Kdriver.System.DirectIO.WriteString("smua.trigger.arm.stimulus = 0", true);
                Kdriver.System.DirectIO.WriteString("smua.trigger.source.stimulus = 0", true);
                Kdriver.System.DirectIO.WriteString("smua.trigger.measure.stimulus = 0", true);
                Kdriver.System.DirectIO.WriteString("smua.trigger.endpulse.stimulus = 0", true);

                Kdriver.Sweep.set_SourceChangesEnabled("A", true);
                Kdriver.Sweep.set_MeasurementsEnabled("A", true);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error\n" + ex.Message);
                return;
            }

            if (Kdriver.Initialized)
            {
                lbl_Connection.Text = "Keithley connected";
                Keithley_connected = true;
                btt_connect.Text = "Disconnect";

                SendCommand("smua.reset()");
                SendCommand("display.screen = display.SMUA");
                SendCommand("format.data = format.ASCII");
                SendCommand("smua.nvbuffer1.clear()");
                SendCommand("display.smua.measure.func = display.MEASURE_DCAMPS");

                SendCommand("smub.reset()");
                SendCommand("display.screen = display.SMUA_SMUB");
                SendCommand("smub.nvbuffer1.clear()");
                SendCommand("display.smub.measure.func = display.MEASURE_DCAMPS");
            }
        }

        private void BgW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Check if user wants to cancel
            if (bgW.CancellationPending)
            { return; }

            pB_progress.Value = e.ProgressPercentage;
            pB_progress.Update();

            ct_DS.Titles[0].Text = String.Format("Drain-Source Current (Vd = {0} V)", setVd);
            ct_GS.Titles[0].Text = String.Format("Gate-Source Current (Vg = {0} V)", setVg);

            if (DrainCur == 0)
            {
                IdsT.Points.AddXY(tscb_xAxis.SelectedIndex == 0 ? setVd : setVg, Math.Abs(DrainCur + 1e-9));
            }
            else
            {
                IdsT.Points.AddXY(tscb_xAxis.SelectedIndex == 0 ? setVd : setVg, Math.Abs(DrainCur));
            }

            if (GateCur == 0)
            {
                IgsT.Points.AddXY(setVg, Math.Abs(GateCur + 1e-9));
            }
            else
            {
                IgsT.Points.AddXY(setVg, Math.Abs(GateCur));
            }
            

            ct_DS.Update();
            ct_GS.Update();

            if (nextVg)
            {
                nextVg = false;
                ClearCharts();

                //Add Series to graphs
                ct_DS.Series.Add(IdsT);
                ct_GS.Series.Add(IgsT);
            }
        }

        private void BgW_LT_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Check if user wants to cancel
            if (bgW.CancellationPending){ return; }

            double t = time / 1000;

            pB_progress.Value = e.ProgressPercentage;
            pB_progress.Update();

            TimeSpan span = new TimeSpan(0, 0, 0, 0, (int)time);
            lbl_timeElapsed.Text = span.ToString(@"hh\:mm\:ss");

            if ((int)Nud_time_duration.Value > 3)
            {
                t /= 60;
                IdsT.XValueType = ChartValueType.Int32;
                IgsT.XValueType = ChartValueType.Int32;
            }

            IdsT.Points.AddXY(t, Math.Abs(DrainCur));
            IgsT.Points.AddXY(t, Math.Abs(GateCur));

            VdT.Points.AddXY(t, Math.Abs(DrainVolt));
            VgT.Points.AddXY(t, Math.Abs(GateVolt));

            ct_time_DS.Update();
            ct_time_GS.Update();

            if (backup)
            {
                SaveFile(true);
                backup = false;
            }
        }

        private void BgWPulse_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Check if user wants to cancel
            if (bgW.CancellationPending){ return; }

            pB_progress.Value = e.ProgressPercentage;
            pB_progress.Update();

            LV_sweep.Items[e.ProgressPercentage].BackColor = Color.LightGreen;
            LV_sweep.Update();
        }

        private void BgW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SendCommand("smua.source.levelv = 0");
            SendCommand("smua.source.output=0");
            SendCommand("smub.source.levelv = 0");
            SendCommand("smub.source.output=0");

            pB_progress.Style = ProgressBarStyle.Blocks;
            pB_progress.Value = 0;
            pB_progress.MarqueeAnimationSpeed = 0;
            btt_cancel.Enabled = false;
            btt_start.Enabled = true;
            gb_display.Enabled = true;

            if (e.Cancelled)
            {
                MessageBox.Show("Measurement cancelled " + (overflow ? "due to overflow" : "by user"), "Measurement cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Thread.Sleep(500);
                overflow = false;
                return;
            }
            else
            {
                MessageBox.Show("Measurement finished ", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //System.IO.Directory.Delete(txt_path.Text + "/Backup");
            }

            if (chk_autosave.Checked)
            {
                SaveFile(false);
            }

        }

        private void BgW_DoWork(object sender, DoWorkEventArgs e)
        {
            double VstartD = Convert.ToDouble(Nud_drain_start.Value);
            double VendD = Convert.ToDouble(Nud_drain_end.Value);
            double VincD = Convert.ToDouble(Nud_drain_step.Value);

            double VstartG = Convert.ToDouble(Nud_gate_start.Value);
            double VendG = Convert.ToDouble(Nud_gate_end.Value);
            double VincG = Convert.ToDouble(Nud_gate_step.Value);

            double delay = Convert.ToDouble(Nud_delay.Value) * 1000;

            bool gateSweep = chk_gate_sweep.Checked;
            bool drainSweep = chk_drain_sweep.Checked;
            bool bothSweep = gateSweep & drainSweep;

            if (!chk_drain_sweep.Checked)
            {
                VendD = VstartD;
            }
            if (!chk_gate_sweep.Checked)
            {
                VendG = VstartG;
            }

            if (VstartD > VendD && (VincD > 0))
            {
                VincD = -VincD;
            }
            if (VstartG > VendG && (VincG > 0))
            {
                VincG = -VincG;
            }

            int nPointsD = (int)Math.Ceiling(Math.Abs((VendD - VstartD) / VincD)) + 1;
            int nPointsG = (int)Math.Ceiling(Math.Abs((VendG - VstartG) / VincG)) + 1;
            int nPoints = nPointsD * nPointsG;
            int count = 0;
            Step = isTransfer ? nPointsG : nPointsD;

            //Activate outputs
            SendCommand("smua.source.output=1");
            SendCommand("smub.source.output=1");

            stopwatch.Restart();

            if (!isTransfer) // OUTPUT MODE
            {
                double vg = VstartG;
                for (int i = 0; i < nPointsG; i++)
                {
                    double vd = VstartD;
                    for (int j = 0; j < nPointsD; j++)
                    {
                        // Check if user wants to cancel
                        if (bgW.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                        setVd = vd;
                        setVg = vg;
                        SetVoltage(vd, vg);
                        MeasureIV();
                        VdList.Add(DrainVolt);
                        VgList.Add(GateVolt);
                        IdsList.Add(DrainCur);
                        IgsList.Add(GateCur);
                        //time = count * delay;
                        TimeList.Add(stopwatch.ElapsedMilliseconds);
                        TimeIRLList.Add(DateTime.Now.TimeOfDay.ToString());
                        bgW.ReportProgress(100 / nPoints * (++count));
                        Thread.Sleep((int)delay);
                        vd += VincD;
                    }
                    vg += VincG;
                    if (i + 1 < nPointsG)
                    {
                        nextVg = true;
                    }
                }
            }
            else // TRANSFER MODE
            {
                double vd = VstartD;
                for (int i = 0; i < nPointsD; i++)
                {
                    double vg = VstartG;
                    for (int j = 0; j < nPointsG; j++)
                    {
                        // Check if user wants to cancel
                        if (bgW.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                        setVd = vd;
                        setVg = vg;
                        SetVoltage(vd, vg);
                        MeasureIV();
                        VdList.Add(DrainVolt);
                        VgList.Add(GateVolt);
                        IdsList.Add(DrainCur);
                        IgsList.Add(GateCur);
                        //time = count * delay;
                        TimeList.Add(stopwatch.ElapsedMilliseconds);
                        TimeIRLList.Add(DateTime.Now.TimeOfDay.ToString());
                        bgW.ReportProgress(100 / nPoints * (++count));
                        Thread.Sleep((int)delay);
                        vg += VincG;
                    }
                    vd += VincD;
                    if (i + 1 < nPointsD)
                    {
                        nextVg = true;
                    }
                }
            } 
        }

        private void BgW_LT_DoWork(object sender, DoWorkEventArgs e)
        {           
            setVd = Convert.ToDouble(Nud_time_drainSet.Value);
            setVg = Convert.ToDouble(Nud_time_gateSet.Value);
            long limitMS = Convert.ToInt64(Nud_time_duration.Value * 60000);
            double delay = Convert.ToDouble(Nud_time_interval.Value) * 1000;
            int counter = 0;

            SetVoltage(setVd, setVg);

            //Activate outputs
            SendCommand("smua.source.output=1");
            SendCommand("smub.source.output=1");

            stopwatch.Restart();
            while (stopwatch.ElapsedMilliseconds <= limitMS)
            {
                // Check if user wants to cancel
                if (bgW_LT.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                SetVoltage(Convert.ToDouble(Nud_time_drainSet.Value), Convert.ToDouble(Nud_time_gateSet.Value));
                limitMS = Convert.ToInt64(Nud_time_duration.Value * 60000);
                delay = Convert.ToDouble(Nud_time_interval.Value) * 1000;

                MeasureIV();

                VdList.Add(DrainVolt);
                VgList.Add(GateVolt);
                IdsList.Add(DrainCur);
                IgsList.Add(GateCur);
                time = stopwatch.ElapsedMilliseconds;
                TimeList.Add(stopwatch.ElapsedMilliseconds);
                TimeIRLList.Add(DateTime.Now.TimeOfDay.ToString());

                if (++counter == saveNumber)
                {
                    counter = 0;
                    backup = true;
                }

                bgW_LT.ReportProgress(50);
                Thread.Sleep((int)delay);
            }
        }

        private void bgWPulse_DoWork(object sender, DoWorkEventArgs e)
        {
            setVd = Convert.ToDouble(Nud_time_drainSet.Value);
            long limitMS = Convert.ToInt64(Nud_time_duration.Value * 60000);
            double delay = Convert.ToDouble(Nud_time_interval.Value) * 1000;
            int nitm = LV_sweep.Items.Count;

            if (LV_sweep.Items.Count == 0)
            {
                e.Cancel = true;
                MessageBox.Show("No voltages to sweep in list.");
                return;
            }

            SetVoltageD(setVd);

            //Activate outputs
            SendCommand("smua.source.output=1");

            for (int i = 0; i < nitm; i++)
            {
                // Check if user wants to cancel
                if (bgW_LT.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                SetVoltageG(Level[i]);
                SendCommand("smub.source.output=1");
                Thread.Sleep((int)Duration[i]);
                SendCommand("smub.source.output=0");
                Thread.Sleep((int)Delay[i]);
                bgWPulse.ReportProgress(i);
            }
        }

        private void ConvertData()
        {
            int nitm = LV_sweep.Items.Count;
            Level = new double[nitm];
            Duration = new double[nitm];           
            Delay = new double[nitm];
            
            foreach (ListViewItem itm in LV_sweep.Items)
            {
                int idx = itm.Index;
                Level[idx] = Convert.ToDouble(itm.SubItems[1].Text);
                Duration[idx] = Convert.ToDouble(itm.SubItems[2].Text);
                Delay[idx] = Convert.ToDouble(itm.SubItems[3].Text);

            }
        }

        private void Chk_Pulse_CheckedChanged(object sender, EventArgs e)
        {
            if (Chk_Pulse.Checked)
            {
                sc_list.Enabled = true;
                Nud_time_gateSet.Enabled = false;
                Gb_timing.Enabled = false;
            }
            else
            {
                sc_list.Enabled = false;
                Nud_time_gateSet.Enabled = true;
                Gb_timing.Enabled = true;
            }
        }

        private void Btt_pulse_clear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Delete all elements in the table?", "Delete all", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                LV_sweep.Items.Clear();
            }
        }

        private void Btt_pulse_duplicate_Click(object sender, EventArgs e)
        {
            if (LV_sweep.SelectedItems.Count == 0)
            {
                return;
            }
            ListViewItem lvi = LV_sweep.SelectedItems[0];
            int index = lvi.Index + 1;
            LV_sweep.Items.Insert(index, (ListViewItem)lvi.Clone());
            lvi.Selected = true;
            RestoreIdx();
            LV_sweep.Focus();
        }

        private void RestoreIdx()
        {
            foreach (ListViewItem itm in LV_sweep.Items)
            {
                itm.SubItems[0].Text = (itm.Index + 1).ToString();
            }
        }

        private void Btt_pulse_save_Click(object sender, EventArgs e)
        {
            if (LV_sweep.Items.Count == 0)
            {
                return;
            }

            string sep, filename = "TransistorPulseList_";
            NumberFormatInfo info = (NumberFormatInfo)NumberFormatInfo.CurrentInfo.Clone();
            CultureInfo culture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            // use . as decimal Seperator and , as data seperator
            info.NumberDecimalSeparator = ".";
            sep = ",";
            culture.NumberFormat = info;
            Thread.CurrentThread.CurrentCulture = culture;

            filename += DateTime.Now.Year.ToString();
            if (DateTime.Now.Month < 10)
            {
                filename += "0";
            }
            filename += DateTime.Now.Month.ToString();
            if (DateTime.Now.Day < 10)
            {
                filename += "0";
            }
            filename += DateTime.Now.Day.ToString() + "_";
            filename += DateTime.Now.ToLongTimeString().Replace(":", "");

            Sfd_list.AddExtension = false;
            Sfd_list.InitialDirectory = txt_path.Text;
            Sfd_list.FileName = filename;
            Sfd_list.Filter = "Comma-separated values (*.csv)| *.csv| Plain Text(*.txt) | *.txt| Plain Data(*.dat) | *.dat";

            if (Sfd_list.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter writer = new StreamWriter(Sfd_list.FileName))
                {
                    writer.WriteLine(String.Format("{0}{8}{1}{8}{2}{8}{3}{8}{4}{8}{5}{8}{6}{8}{7}", "#", "Duration (ms)", "Type (A)", "Level (A)", "Limit (A)", "Type (B)", "Level (B)", "Limit (B)", sep));
                    foreach (ListViewItem itm in LV_sweep.Items)
                    {
                        writer.WriteLine(String.Format("{0}{8}{1}{8}{2}{8}{3}{8}{4}{8}{5}{8}{6}{8}{7}", itm.SubItems[0].Text, itm.SubItems[1].Text, itm.SubItems[2].Text, itm.SubItems[3].Text, itm.SubItems[4].Text, itm.SubItems[5].Text, itm.SubItems[6].Text, itm.SubItems[7].Text, sep));
                    }
                }

                tsl_saved.Visible = true;
                tsl_link.Visible = true;
                tsl_link.Text = txt_path.Text;
            }
        }

        private void Btt_pulse_load_Click(object sender, EventArgs e)
        {
            NumberFormatInfo info = (NumberFormatInfo)NumberFormatInfo.CurrentInfo.Clone();
            CultureInfo culture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            // use . as decimal Seperator and , as data seperator
            info.NumberDecimalSeparator = ".";
            culture.NumberFormat = info;
            Thread.CurrentThread.CurrentCulture = culture;

            string path;
            string[] tmp, readText;
            int nPoints;

            Ofd_list.Title = "Select file containing program data";
            Ofd_list.InitialDirectory = txt_path.Text;

            if (Ofd_list.ShowDialog() == DialogResult.OK)
            {
                path = Ofd_list.FileName;
                readText = File.ReadAllLines(path);
                nPoints = readText.Length - 1;

                for (int i = 1; i <= nPoints; i++)
                {
                    tmp = readText[i].Split(',');
                    LV_sweep.Items.Add(new ListViewItem(tmp));
                    LV_sweep.Items[LV_sweep.Items.Count - 1].Checked = true;
                }
                RestoreIdx();
            }
        }

        private void Btt_pulse_remove_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem itm in LV_sweep.Items)
            {
                if (itm.Selected)
                {
                    itm.Remove();
                }
            }
            RestoreIdx();
        }

        private void Nud_pulse_level_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown tmp = (NumericUpDown)sender;
            tmp.BackColor = (tmp.Value > 10 ? Color.Orange : SystemColors.Window);
        }

        private void Btt_pulse_up_Click(object sender, EventArgs e)
        {
            if (LV_sweep.SelectedItems.Count == 0)
            {
                return;
            }
            ListViewItem lvi = LV_sweep.SelectedItems[0];
            if (lvi.Index > 0)
            {
                int index = lvi.Index - 1;
                LV_sweep.Items.RemoveAt(lvi.Index);
                LV_sweep.Items.Insert(index, lvi);
                lvi.Selected = true;
            }
            RestoreIdx();
            LV_sweep.Focus();
        }

        private void Btt_pulse_down_Click(object sender, EventArgs e)
        {
            if (LV_sweep.SelectedItems.Count == 0)
            {
                return;
            }
            ListViewItem lvi = LV_sweep.SelectedItems[0];
            if (lvi.Index < LV_sweep.Items.Count - 1)
            {
                int index = lvi.Index + 1;
                LV_sweep.Items.RemoveAt(lvi.Index);
                LV_sweep.Items.Insert(index, lvi);
                lvi.Selected = true;
            }
            RestoreIdx();
            LV_sweep.Focus();
        }

        

        private void Btt_pulse_add_Click(object sender, EventArgs e)
        {
            string[] itm = new string[4];
            itm[0] = (++nVals).ToString();
            itm[1] = Nud_pulse_level.Value.ToString();
            itm[2] = Nud_pulse_duration.Value.ToString();
            itm[3] = Nud_pulse_delay.Value.ToString();
            LV_sweep.Items.Add(new ListViewItem(itm));
            LV_sweep.Items[LV_sweep.Items.Count - 1].Checked = true;
            RestoreIdx();
        }

        private void Btt_Clear_Click(object sender, EventArgs e)
        {
            ClearData();
        }

        private void Tscb_xAxis_SelectedIndexChanged(object sender, EventArgs e)
        {
            ct_DS.ChartAreas[0].AxisX.Title = tscb_xAxis.SelectedIndex == 0 ? "Drain Voltage(V)" : "Gate Voltage (V)";

            if (VdList.Count != 0)
            {
                IdsT.Points.Clear();
                if (tscb_xAxis.SelectedIndex == 0)
                {
                    for (int i = 0; i < VdList.Count; i++)
                    {
                        IdsT.Points.AddXY(VdList[i], Math.Abs(IdsList[i]));
                    }
                }
                else
                {
                    for (int i = 0; i < VdList.Count; i++)
                    {
                        IdsT.Points.AddXY(VgList[i], Math.Abs(IdsList[i]));
                    }
                }              
                ct_DS.Update();
            }
        }

        private void Tsl_link_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(txt_path.Text);
        }

        private void Tscb_axis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tscb_DrainAxis.SelectedIndex == 0)
            {
                ct_DS.ChartAreas[0].AxisY.IsLogarithmic = false;
                ct_time_DS.ChartAreas[0].AxisY.IsLogarithmic = false;
            }
            else
            {
                ct_DS.ChartAreas[0].AxisY.IsLogarithmic = true;
                ct_time_DS.ChartAreas[0].AxisY.IsLogarithmic = true;
            }
        }

        private void Tscb_gateAxis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tscb_DrainAxis.SelectedIndex == 0)
            {
                ct_GS.ChartAreas[0].AxisY.IsLogarithmic = false;
                ct_time_GS.ChartAreas[0].AxisY.IsLogarithmic = false;
            }
            else
            {
                ct_GS.ChartAreas[0].AxisY.IsLogarithmic = true;
                ct_time_GS.ChartAreas[0].AxisY.IsLogarithmic = true;
            }
        }

        private void ToolStripButton1_Click(object sender, EventArgs e)
        {
            SendCommand("smua.reset()");
            SendCommand("display.screen = display.SMUA");
            SendCommand("format.data = format.ASCII");
            SendCommand("smua.nvbuffer1.clear()");
            SendCommand("display.smua.measure.func = display.MEASURE_DCAMPS");
            SendCommand("smua.source.levelv = 0");

            SendCommand("smub.reset()");
            SendCommand("display.screen = display.SMUA_SMUB");
            SendCommand("smub.nvbuffer1.clear()");
            SendCommand("display.smub.measure.func = display.MEASURE_DCAMPS");
            SendCommand("smub.source.levelv = 0");

            // Set Measurement Range
            if (cb_drain_range.SelectedIndex == 0)
            {
                SendCommand("smua.measure.autorangei = smua.AUTORANGE_ON");
            }
            else
            {
                SendCommand(String.Format("smua.measure.rangei = {0}", range[cb_drain_range.SelectedIndex - 1]));
            }

            if (cb_gate_range.SelectedIndex == 0)
            {
                SendCommand("smub.measure.autorangei = smub.AUTORANGE_ON");
            }
            else
            {
                SendCommand(String.Format("smub.measure.rangei = {0}", range[cb_drain_range.SelectedIndex - 1]));
            }

            //Activate outputs
            SendCommand("smua.source.output=1");
            SendCommand("smub.source.output=1");

            SetVoltage(5, 2);
            MeasureIV();

            SendCommand("smua.source.output=0");
            SendCommand("smub.source.output=0");

            MessageBox.Show(String.Format("Vd: {0}\nIds: {1}\nVg: {2}\nIgs: {3}\n", DrainVolt, DrainCur, GateVolt, GateCur));
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tc_mode.SelectedIndex < 2)
            {
                tc_mode.TabPages[tc_mode.SelectedIndex].Controls.Add(sc_iv_charts);
                tscb_xAxis.SelectedIndex = tc_mode.SelectedIndex;

                if (tc_mode.SelectedIndex == 0) // output
                {
                    isTransfer = false;
                    

                    Properties.Settings.Default.Drain_Start_T = Nud_drain_start.Value;
                    Properties.Settings.Default.Drain_Step_T = Nud_drain_step.Value;
                    Properties.Settings.Default.Drain_End_T = Nud_drain_end.Value;
                    Properties.Settings.Default.Drain_Range_T = cb_drain_range.SelectedIndex;
                    Properties.Settings.Default.Drain_Sweep_T = chk_drain_sweep.Checked;

                    Properties.Settings.Default.Gate_Start_T = Nud_gate_start.Value;
                    Properties.Settings.Default.Gate_Step_T = Nud_gate_step.Value;
                    Properties.Settings.Default.Gate_End_T = Nud_gate_end.Value;
                    Properties.Settings.Default.Gate_Range_T = cb_gate_range.SelectedIndex;
                    Properties.Settings.Default.Gate_Sweep_T = chk_gate_sweep.Checked;

                    Properties.Settings.Default.Delay_T = Nud_delay.Value;
                    Properties.Settings.Default.Save();


                    Nud_drain_start.Value = Properties.Settings.Default.Drain_Start_O;
                    Nud_drain_step.Value = Properties.Settings.Default.Drain_Step_O;
                    Nud_drain_end.Value = Properties.Settings.Default.Drain_End_O;
                    cb_drain_range.SelectedIndex = Properties.Settings.Default.Drain_Range_O;
                    chk_drain_sweep.Checked = Properties.Settings.Default.Drain_Sweep_O;

                    Nud_gate_start.Value = Properties.Settings.Default.Gate_Start_O;
                    Nud_gate_step.Value = Properties.Settings.Default.Gate_Step_O;
                    Nud_gate_end.Value = Properties.Settings.Default.Gate_End_O;
                    cb_gate_range.SelectedIndex = Properties.Settings.Default.Gate_Range_O;
                    chk_gate_sweep.Checked = Properties.Settings.Default.Gate_Sweep_O;

                    Nud_delay.Value = Properties.Settings.Default.Delay_O;

                    chk_drain_sweep.Checked = true;
                    chk_drain_sweep.Enabled = false;
                    chk_gate_sweep.Enabled = true;

                }
                else //transfer
                {
                    isTransfer = true;

                    Properties.Settings.Default.Drain_Start_O = Nud_drain_start.Value;
                    Properties.Settings.Default.Drain_Step_O = Nud_drain_step.Value;
                    Properties.Settings.Default.Drain_End_O = Nud_drain_end.Value;
                    Properties.Settings.Default.Drain_Range_O = cb_drain_range.SelectedIndex;
                    Properties.Settings.Default.Drain_Sweep_O = chk_drain_sweep.Checked;

                    Properties.Settings.Default.Gate_Start_O = Nud_gate_start.Value;
                    Properties.Settings.Default.Gate_Step_O = Nud_gate_step.Value;
                    Properties.Settings.Default.Gate_End_O = Nud_gate_end.Value;
                    Properties.Settings.Default.Gate_Range_O = cb_gate_range.SelectedIndex;
                    Properties.Settings.Default.Gate_Sweep_O = chk_gate_sweep.Checked;

                    Properties.Settings.Default.Delay_O = Nud_delay.Value;
                    Properties.Settings.Default.Save();

                    Nud_drain_start.Value = Properties.Settings.Default.Drain_Start_T;
                    Nud_drain_step.Value = Properties.Settings.Default.Drain_Step_T;
                    Nud_drain_end.Value = Properties.Settings.Default.Drain_End_T;
                    cb_drain_range.SelectedIndex = Properties.Settings.Default.Drain_Range_T;
                    chk_drain_sweep.Checked = Properties.Settings.Default.Drain_Sweep_T;

                    Nud_gate_start.Value = Properties.Settings.Default.Gate_Start_T;
                    Nud_gate_step.Value = Properties.Settings.Default.Gate_Step_T;
                    Nud_gate_end.Value = Properties.Settings.Default.Gate_End_T;
                    cb_gate_range.SelectedIndex = Properties.Settings.Default.Gate_Range_T;
                    chk_gate_sweep.Checked = Properties.Settings.Default.Gate_Sweep_T;

                    Nud_delay.Value = Properties.Settings.Default.Delay_T;

                    chk_gate_sweep.Checked = true;
                    chk_gate_sweep.Enabled = false;
                    chk_drain_sweep.Enabled = true;
                }

                gb_timed.Visible = false;
                gb_drain.Visible = true;
                gb_gate.Visible = true;
                gb_general.Visible = true;
            } else
            {
                gb_drain.Visible = false;
                gb_gate.Visible = false;
                gb_general.Visible = false;
                gb_timed.Visible = true;
            }

        }

        private void Rb_drain_current_CheckedChanged(object sender, EventArgs e)
        {
            ct_time_DS.Titles[0].Text = "Drain" + (rb_drain_current.Checked ? "-Source Current" : " Voltage");
            ct_time_DS.ChartAreas[0].AxisY.Title = (rb_drain_current.Checked ? "Current (A)" : "Voltage (V)");
            Cdrain = rb_drain_current.Checked;
        }

        

        private void Rb_gate_current_CheckedChanged(object sender, EventArgs e)
        {
            ct_time_GS.Titles[0].Text = "Gate" + (rb_gate_current.Checked ? "-Source Current" : " Voltage");
            ct_time_GS.ChartAreas[0].AxisY.Title = (rb_gate_current.Checked ? "Current (A)" : "Voltage (V)");
            Cgate = rb_gate_current.Checked;
        }

        

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Keithley_connected)
            {
                Kdriver.Close();
            }

            Properties.Settings.Default.KeithleyID = txt_KeithleyID.Text;

            Properties.Settings.Default.Drain_Start_O = Nud_drain_start.Value;
            Properties.Settings.Default.Drain_Step_O = Nud_drain_step.Value;
            Properties.Settings.Default.Drain_End_O = Nud_drain_end.Value;
            Properties.Settings.Default.Drain_Range_O = cb_drain_range.SelectedIndex;
            Properties.Settings.Default.Drain_Sweep_O = chk_drain_sweep.Checked;

            Properties.Settings.Default.Gate_Start_O = Nud_gate_start.Value;
            Properties.Settings.Default.Gate_Step_O = Nud_gate_step.Value;
            Properties.Settings.Default.Gate_End_O = Nud_gate_end.Value;
            Properties.Settings.Default.Gate_Range_O = cb_gate_range.SelectedIndex;
            Properties.Settings.Default.Gate_Sweep_O = chk_gate_sweep.Checked;

            Properties.Settings.Default.Delay_O = Nud_delay.Value;
            Properties.Settings.Default.Name = txt_name.Text;
            Properties.Settings.Default.Path = txt_path.Text;
            Properties.Settings.Default.Autosave = chk_autosave.Checked;

            Properties.Settings.Default.Time_Drain_Voltage = Nud_time_drainSet.Value;
            Properties.Settings.Default.Time_Drain_Range = cb_time_drain_range.SelectedIndex;
            Properties.Settings.Default.Time_Gate_Voltage = Nud_time_gateSet.Value;
            Properties.Settings.Default.Time_Gate_Range = cb_time_gate_range.SelectedIndex;
            Properties.Settings.Default.Time_Interval = Nud_time_interval.Value;
            Properties.Settings.Default.Time_Duration = Nud_time_duration.Value;

            Properties.Settings.Default.Save();

            if (!isSaved)
            {
                if (MessageBox.Show("Save data before closing?", "Unsaved data", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SaveFile(false);
                }
            }
        }

        private void Chk_drain_sweep_CheckedChanged(object sender, EventArgs e)
        {
            Checkboxes();
        }

        private void Checkboxes()
        {
            if (!chk_drain_sweep.Checked)
            {
                Nud_drain_step.Enabled = false;
                Nud_drain_end.Enabled = false;
                lbl_drain.Text = "Value:";
            }
            else
            {
                Nud_drain_step.Enabled = true;
                Nud_drain_end.Enabled = true;
                lbl_drain.Text = "Start:";
            }

            if (!chk_gate_sweep.Checked)
            {
                Nud_gate_step.Enabled = false;
                Nud_gate_end.Enabled = false;
                lbl_gate.Text = "Value:";
            }
            else
            {
                Nud_gate_step.Enabled = true;
                Nud_gate_end.Enabled = true;
                lbl_gate.Text = "Start:";
            }
        }

        private void Btt_save_Click(object sender, EventArgs e)
        {
            SaveFile(false);
        }

        private void Tsb_zero_Click(object sender, EventArgs e)
        {
            SendCommand("smua.source.output=0");
            SendCommand("smub.source.output=0");
            SendCommand("smua.source.levelv = 0");
            SendCommand("smub.source.levelv = 0");
        }

        private void SaveFile(bool isBackup)
        {
            string sep, path, filename = "";
            NumberFormatInfo info = (NumberFormatInfo)NumberFormatInfo.CurrentInfo.Clone();
            CultureInfo culture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            // use . as decimal Seperator and , as data seperator
            info.NumberDecimalSeparator = ".";
            sep = ",";
            culture.NumberFormat = info;
            Thread.CurrentThread.CurrentCulture = culture;

            double[] Times = TimeList.ToArray();
            string[] IRLTimes = TimeIRLList.ToArray();
            double[] IDS = IdsList.ToArray();
            double[] IGS = IgsList.ToArray();
            double[] VD = VdList.ToArray();
            double[] VG = VgList.ToArray();

            int nPoints = VgList.Count;

            if (IdsList.Count == 0) { return; }

            //Add date time string
            filename += DateTime.Now.Year.ToString().Replace("20", "");
            if (DateTime.Now.Month < 10)
            {
                filename += "0";
            }
            filename += DateTime.Now.Month.ToString();
            if (DateTime.Now.Day < 10)
            {
                filename += "0";
            }
            filename += DateTime.Now.Day.ToString() + "_";
            filename += DateTime.Now.ToLongTimeString().Replace(":", "") + "_";

            // Add type label
            if (isLifetime) {
                filename += "timed_";
            }
            else
            {
                if (isTransfer)
                {
                    filename += "tsf_";
                }
                else
                {
                    filename += "out_";
                }
            }

            // Add name from textbox
            filename += txt_name.Text;

            //if Backup then create backup directory and empty it
            if (isBackup)
            {
                System.IO.Directory.CreateDirectory(txt_path.Text + "/Backup");
                var files = new DirectoryInfo(txt_path.Text + "/Backup").EnumerateFiles().OrderByDescending(f => f.CreationTime).Skip(5).ToList();
                files.ForEach(f => f.Delete());
            }           
            

            path = txt_path.Text + (isBackup ? "/Backup/" : "/") + filename + ".csv";           

            bool gateSweep = chk_gate_sweep.Checked;
            bool drainSweep = chk_drain_sweep.Checked;
            bool bothSweep = gateSweep & drainSweep;

            if (isLifetime)
            {
                using (StreamWriter writer = new StreamWriter(path))
                {
                    writer.WriteLine(String.Format("Date:{0}{1}", sep,DateTime.Now.Date.ToShortDateString()));
                    writer.WriteLine(String.Format("Measurement startet at:{0}{1}", sep, IRLTimes[0].Split('.')[0]));
                    writer.WriteLine(String.Format("Measurement Mode:{0}{1}", sep, "Time"));
                    writer.WriteLine(String.Format("Drain Current Range:{0}{1}", sep, cb_drain_range.Text));
                    writer.WriteLine(String.Format("Gate Current Range:{0}{1}", sep, cb_gate_range.Text));
                    writer.WriteLine();

                    writer.WriteLine(String.Format("Real_Time{0}Measurement_Time{0}V_Drain{0}I_Drain{0}V_Gate{0}I_Gate", sep));
                    writer.WriteLine(String.Format("{0}min{0}V{0}A{0}V{0}A", sep));
                    for (int i = 0; i < nPoints; i++)
                    {
                        writer.WriteLine(String.Format("{0}{6}{1}{6}{2}{6}{3}{6}{4}{6}{5}", IRLTimes[i].Split('.')[0], Times[i] / 60000.0, VD[i], IDS[i], VG[i], IGS[i], sep));
                    }
                }
            }
            else
            {             
                if (drainSweep || bothSweep)
                {
                    if (bothSweep && !isTransfer) // Output Mode
                    {
                        int Dstep = nPoints / Step;
                        using (StreamWriter writer = new StreamWriter(path))
                        {
                            writer.WriteLine(String.Format("Date:{0}{1}", sep, DateTime.Now.Date.ToShortDateString()));
                            writer.WriteLine(String.Format("Measurement startet at:{0}{1}", sep, IRLTimes[0].Split('.')[0]));
                            writer.WriteLine(String.Format("Measurement Mode:{0}{1}", sep, isTransfer ? "Transfer" : "Output"));
                            writer.WriteLine(String.Format("Drain Current Range:{0}{1}", sep, cb_drain_range.Text));
                            writer.WriteLine(String.Format("Gate Current Range:{0}{1}", sep, cb_gate_range.Text));
                            writer.WriteLine();

                            for (int i = 0; i < Dstep; i++)
                            {
                                writer.WriteLine(String.Format("Real_Time{0}Measurement_Time{0}V_Drain{0}I_Drain{0}V_Gate{0}I_Gate", sep));
                                writer.WriteLine(String.Format("{0}s{0}V{0}A{0}V{0}A", sep));
                                for (int j = 0; j < Step; j++)
                                {
                                    writer.WriteLine(String.Format("{0}{6}{1}{6}{2}{6}{3}{6}{4}{6}{5}", IRLTimes[i].Split('.')[0], Times[i * Step + j] / 1000.0, VD[i * Step + j], IDS[i * Step + j], VG[i * Step + j], IGS[i * Step + j], sep));
                                }
                                writer.WriteLine();
                            }
                        } 
                    }
                    else  // Transfer Mode
                    {
                        int Gstep = nPoints / Step;
                        using (StreamWriter writer = new StreamWriter(path))
                        {
                            writer.WriteLine(String.Format("Date:{0}{1}", sep, DateTime.Now.Date.ToShortDateString()));
                            writer.WriteLine(String.Format("Measurement startet at:{0}{1}", sep, IRLTimes[0].Split('.')[0]));
                            writer.WriteLine(String.Format("Measurement Mode:{0}{1}", sep, isTransfer ? "Transfer" : "Output"));
                            writer.WriteLine(String.Format("Drain Current Range:{0}{1}", sep, cb_drain_range.Text));
                            writer.WriteLine(String.Format("Gate Current Range:{0}{1}", sep, cb_gate_range.Text));
                            writer.WriteLine();

                            for (int i = 0; i < Gstep; i++)
                            {
                                writer.WriteLine(String.Format("Real_Time{0}Measurement_Time{0}V_Drain{0}I_Drain{0}V_Gate{0}I_Gate", sep));
                                writer.WriteLine(String.Format("{0}s{0}V{0}A{0}V{0}A", sep));
                                for (int j = 0; j < Step; j++)
                                {
                                    writer.WriteLine(String.Format("{0}{6}{1}{6}{2}{6}{3}{6}{4}{6}{5}", IRLTimes[i].Split('.')[0], Times[i * Step + j] / 1000.0, VD[i * Step + j], IDS[i * Step + j], VG[i * Step + j], IGS[i * Step + j], sep));
                                }
                                writer.WriteLine();
                            }
                        }
                    }
                }
                if (gateSweep && !bothSweep)
                {
                    using (StreamWriter writer = new StreamWriter(path))
                    {
                        writer.WriteLine(String.Format("Date:{0}{1}", sep, DateTime.Now.Date.ToShortDateString()));
                        writer.WriteLine(String.Format("Measurement startet at:{0}{1}", sep, IRLTimes[0].Split('.')[0]));
                        writer.WriteLine(String.Format("Measurement Mode:{0}{1}", sep, isTransfer ? "Transfer" : "Output"));
                        writer.WriteLine(String.Format("Drain Current Range:{0}{1}", sep, cb_drain_range.Text));
                        writer.WriteLine(String.Format("Gate Current Range:{0}{1}", sep, cb_gate_range.Text));
                        writer.WriteLine();

                        writer.WriteLine(String.Format("Real_Time{0}Measurement_Time{0}V_Drain{0}I_Drain{0}V_Gate{0}I_Gate", sep));
                        writer.WriteLine(String.Format("{0}s{0}V{0}A{0}V{0}A", sep));
                        for (int j = 0; j < nPoints; j++)
                        {
                            writer.WriteLine(String.Format("{0}{6}{1}{6}{2}{6}{3}{6}{4}{6}{5}", IRLTimes[j].Split('.')[0], Times[j] / 1000.0, VD[j], IDS[j], VG[j], IGS[j], sep));
                        }
                        writer.WriteLine();
                    }
                }
            }

            if (!isBackup)
            {
                tsl_saved.Visible = true;
                tsl_link.Visible = true;
                tsl_link.Text = filename;
                isSaved = true;
            }

        }

        /// <summary>
        /// Clears the data variables for a new measurement
        /// </summary>
        private void ClearData()
        {
            if (!isSaved)
            {
                if (MessageBox.Show("Data not saved. Clear anyway?", "Unsaved data", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
            }

            isSaved = true;

            

            /*Ids = null;
            Igs = null;
            Vg = null;
            Vd = null;
            time = null;*/

            IdsList.Clear();
            IgsList.Clear();
            VdList.Clear();
            VgList.Clear();
            TimeList.Clear();

            tsl_saved.Visible = false;
            tsl_link.Visible = false;

            pB_progress.Value = 0;

            ClearCharts();
        }

        /// <summary>
        /// Clears the charts for a new measurement
        /// </summary>
        private void ClearCharts()
        {
            // Clear Charts
            ct_DS.Series.Clear();
            ct_GS.Series.Clear();
            ct_time_DS.Series.Clear();
            ct_time_GS.Series.Clear();
            IdsT.Points.Clear();
            IgsT.Points.Clear();

            ct_DS.Titles[0].Text = "Drain-Source Current";
            ct_GS.Titles[0].Text = "Gate-Source Current";
        }

        

    }
}
