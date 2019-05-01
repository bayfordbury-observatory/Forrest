/*    ----------------------------------------------------
 * R.W. Forrest Telescope Controller
 *   University of Hertfordshire
 *   David Campbell
 * 
 * 
 *   v0.9.0.0 ---------------------
 *   - Initial release
 *   v0.9.0.1 ---------------------
 *   - Redesigned for 1920x1080
 *   v0.9.1.0 ---------------------
 *   - Added full list of faults from the antenna controller
 *   - Added start/stop buttons
 * 

 * 
*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Threading;

using ShadowEngine;
using ShadowEngine.OpenGL;
using Tao.OpenGl;

using AstroLib;

namespace Forrest {
    public partial class MainForm : Form {

        private SerialPort RxSerialPort = new SerialPort();
        private SerialPort ScopeSerialPort = new SerialPort();

        public Location location = new Location(51.774849.ToRad(), 0.095656.ToRad(), 60);

        public static double ScopetoTarget;

        public HorizontalCoords ScopeHor = new HorizontalCoords();
        public EquatorialCoords ScopeEq = new EquatorialCoords();
        public GalacticCoords ScopeGal = new GalacticCoords();

        public CoordsString CoordString = new CoordsString();

        public HorizontalCoords TargetHor = new HorizontalCoords(82d.ToRad(), 180d.ToRad());
        public EquatorialCoords TargetEq = new EquatorialCoords();
        public GalacticCoords TargetGal = new GalacticCoords();

        public HorizontalCoords MoonHor = new HorizontalCoords();
        public HorizontalCoords SunHor = new HorizontalCoords();

        EquatorialCoords SunEq = new EquatorialCoords();
        EquatorialCoords MoonEq = new EquatorialCoords();

        bool RxInUse;
        bool ScopeInUse;
        int i, n, k;

        int inbyte;
        bool stopscan = false;
        int max, min;
        int steps = 1;

        string saveName;

        public string[,] allFaults = new string[4, 32];
        int[] faultNums = new int[4];

        static int NumSources = 11;

        [DllImport("GDI32.dll")]
        public static extern void SwapBuffers(uint hdc);

        Series ContSeries = new Series();
        int[] ContArray = new int[600];

        Series sChanSeries = new Series();
        int[] sChanArray = new int[300];

        double[] SpectrumBVel = new double[441];
        double[] SpectrumRVel = new double[441];

        int[] SpectrumArrayBSum = new int[441];
        int[] SpectrumArrayBN = new int[441];
        double[] SpectrumArrayBAv = new double[441];

        int[] SpectrumArrayRSum = new int[441];
        int[] SpectrumArrayRN = new int[441];
        double[] SpectrumArrayRAv = new double[441];

        Series SpectrumSeriesR = new Series();
        Series SpectrumSeriesB = new Series();

        public double[] sourceRAarray = new double[NumSources];
        public double[] sourceDecarray = new double[NumSources];

        public Boolean tracking = false;

        int spectrumLoops;


        bool bitSet(byte b, int pos) {
            return (b & (1 << pos)) != 0;
        }

        bool bitSet(Int32 b, int pos) {
            return (b & (1 << pos)) != 0;
        }

        public void LoadLoadingScreen() {
            Application.Run(new LoadScreen());
        }

        public void CommPortSetup() {

            if(!RxSerialPort.IsOpen) {

                var _with1 = RxSerialPort;
                _with1.PortName = "COM2";//"COM2";
                _with1.BaudRate = 2400;
                _with1.DataBits = 8;
                _with1.Parity = Parity.None;
                _with1.StopBits = StopBits.One;
                _with1.Handshake = Handshake.None;
                _with1.WriteTimeout = 250;
                _with1.ReadTimeout = 250;

            }

            try {
                RxSerialPort.Open();
            } catch(Exception ex) {
                MessageBox.Show(ex.Message);
            }

            if(!ScopeSerialPort.IsOpen) {

                var _with1 = ScopeSerialPort;
                _with1.PortName = "COM4";
                _with1.BaudRate = 9600;
                _with1.DataBits = 8;
                _with1.Parity = Parity.None;
                _with1.StopBits = StopBits.One;
                _with1.Handshake = Handshake.None;
                _with1.WriteTimeout = 250;
                _with1.ReadTimeout = 250;
                _with1.NewLine = ">";

            }

            try {
                //ScopeSerialPort.Open();
            } catch(Exception ex) {
                MessageBox.Show(ex.Message);
            }

        }

        public MainForm() {

            //Start loading screen thread
            Thread ThreadT = new Thread(new ThreadStart(LoadLoadingScreen));
            ThreadT.Start();

            InitializeComponent();

            //enable serial ports
            RxInUse = true;
            ScopeInUse = true;

            CommPortSetup();

            RxInUse = false;
            ScopeInUse = false;

            //Default image save format
            SaveChart.Filter = "PNG image | *.png";
            SaveChart.DefaultExt = "png";

            //default data save format
            SaveData.Filter = "CSV file | *.csv";
            SaveData.DefaultExt = "csv";

            Globals.LoadProgress = 10;
            Globals.Status = "Initialising charts";

            //intialise charts
            ContChart.Series.Clear();
            ContSeries.ChartType = SeriesChartType.Line;
            ContSeries.Color = Color.Blue;

            sChanChart.Series.Clear();
            sChanSeries.ChartType = SeriesChartType.Line;
            sChanSeries.Color = Color.Blue;

            Spectrum.Series.Clear();
            SpectrumSeriesR.ChartType = SeriesChartType.Line;
            SpectrumSeriesR.Color = Color.Red;
            SpectrumSeriesB.ChartType = SeriesChartType.Line;
            SpectrumSeriesB.Color = Color.Blue;

            Globals.Status = "Initialising more charts";
            Globals.LoadProgress = 20;

            //Continuum chart setup
            ContChart.Legends[0].Enabled = true;
            ContChart.ChartAreas[0].AxisY.Minimum = 0;
            ContChart.ChartAreas[0].AxisY.Maximum = 4096;
            ContChart.ChartAreas[0].AxisY.Interval = 100;
            ContChart.ChartAreas[0].AxisY.Title = "Millivolts";
            ContChart.ChartAreas[0].AxisX.Minimum = 0;
            ContChart.ChartAreas[0].AxisX.Maximum = 600;
            ContChart.ChartAreas[0].AxisX.IsReversed = true;
            ContChart.ChartAreas[0].AxisX.Interval = 60;
            ContChart.ChartAreas[0].AxisX.Title = "time";
            ContChart.Legends[0].Enabled = false;

            //Single channel
            sChanChart.Legends[0].Enabled = true;
            sChanChart.ChartAreas[0].AxisY.Minimum = 0;
            sChanChart.ChartAreas[0].AxisY.Maximum = 4096;
            sChanChart.ChartAreas[0].AxisY.Interval = 100;
            sChanChart.ChartAreas[0].AxisY.Title = "Millivolts";
            sChanChart.ChartAreas[0].AxisX.Minimum = 0;
            sChanChart.ChartAreas[0].AxisX.Maximum = 300;
            sChanChart.ChartAreas[0].AxisX.IsReversed = true;
            sChanChart.ChartAreas[0].AxisX.Interval = 60;
            sChanChart.ChartAreas[0].AxisX.Title = "time";
            sChanChart.Legends[0].Enabled = false;

            //Spectrograph
            Spectrum.Legends[0].Enabled = true;
            Spectrum.ChartAreas[0].AxisY.Minimum = 0;
            Spectrum.ChartAreas[0].AxisY.Maximum = 4096;
            Spectrum.ChartAreas[0].AxisY.Interval = 100;
            Spectrum.ChartAreas[0].AxisY.Title = "Millivolts";
            Spectrum.ChartAreas[0].AxisX.Minimum = -450;
            Spectrum.ChartAreas[0].AxisX.Maximum = 450;
            Spectrum.ChartAreas[0].AxisX.IsReversed = true;
            Spectrum.ChartAreas[0].AxisX.Interval = 50;
            Spectrum.ChartAreas[0].AxisX.Title = "Velocity (km/s)";
            Spectrum.Legends[0].Enabled = false;

            ContChart.Series.Add(ContSeries);

            Globals.LoadProgress = 30;
            Globals.Status = "Initialising some arrays";

            //initialise arrays

            //continuum
            for(n = 0; n < 600; n++) {
                ContArray[n] = 4097;
            }

            //single channel
            for(n = 0; n < 300; n++) {
                sChanArray[n] = 4097;
            }

            //channel>velocity
            for(n = 0; n < 441; n++) {
                SpectrumRVel[n] = (440 - Convert.ToDouble(n)) * 1.055332251 + 0.5;
                SpectrumBVel[n] = (440 - Convert.ToDouble(n + 440)) * 1.055332251 + 0.5;
            }

            Globals.Status = "Loading settings";
            Globals.LoadProgress = 40;

            /*
             if (Properties.Settings.Default.NeedsUpgrading) {
                 Properties.Settings.Default.Upgrade();
                 Properties.Settings.Default.NeedsUpgrading = false;
             }
             */

            //default settings
            ContInt.Text = "0.3s";
            ContGain.Text = "x5";
            ContOffset.Value = 1400;

            SpecInt.Text = "0.3s";
            SpecGain.Text = "x5";
            SpecOffset.Value = 800;

            Gain70.Value = 10;

            //------------------   DEBUGGING
            // readSettings();
            //-------------------
            
            //steps = Convert.ToInt16(NumSteps.Value);
            //label15.Text = (900.0 / steps).ToString("F3") + " steps";

            //default timer interval
            ContTimer.Interval = Convert.ToInt32(ContGraphInt.Value * 1000);

            Globals.Status = "Loading more settings";
            Globals.LoadProgress = 50;

            //Target source list

            //M31
            sourceRAarray[2] = 10.68470833d.ToRad();
            sourceDecarray[2] = 41.26875d.ToRad();
            //Cassiopea A
            sourceRAarray[3] = 350.85d.ToRad();
            sourceDecarray[3] = 58.815d.ToRad();
            //Cygnus A
            sourceRAarray[4] = 299.8681542d.ToRad();
            sourceDecarray[4] = 40.73391667d.ToRad();
            //Hercules A
            sourceRAarray[5] = 252.7841667d.ToRad();
            sourceDecarray[5] = 4.9925d.ToRad();
            //Virgo A
            sourceRAarray[6] = 187.7059292d.ToRad();
            sourceDecarray[6] = 12.39112222d.ToRad();
            //Taurus A
            sourceRAarray[7] = 83.63308333d.ToRad();
            sourceDecarray[7] = 22.0145d.ToRad();
            //M42
            sourceRAarray[8] = 83.82208333d.ToRad();
            sourceDecarray[8] = -5.391111111d.ToRad();
            //PSR B1931+24
            sourceRAarray[9] = 293.407625d.ToRad();
            sourceDecarray[9] = 24.611d.ToRad();
            //M33
            sourceRAarray[10] = 23.4584166666d.ToRad();
            sourceDecarray[10] = 30.6601944444d.ToRad();

            //list of faults

            faultNums[0] = 17;
            faultNums[1] = 24;
            faultNums[2] = 22;
            faultNums[3] = 30;

            allFaults[0, 0] = "No power at drive cabinet";
            allFaults[0, 1] = "Emergency stop at drive cabinet";
            allFaults[0, 2] = "Maintenance override at drive cabinet";
            allFaults[0, 3] = "Travel limit switch (summary)";
            allFaults[0, 4] = "Azimuth drive fault";
            allFaults[0, 5] = "Elevation drive fault";
            allFaults[0, 6] = "Azimuth CW limit switch";
            allFaults[0, 6] = "Azimuth CW limit switch";
            allFaults[0, 7] = "Azimuth CCW limit switch";
            allFaults[0, 8] = "Elevation upper limit switch";
            allFaults[0, 9] = "Elevation lower limit switch";
            allFaults[0, 10] = "Polarization2 CW limit switch";
            allFaults[0, 11] = "Polarization 2 CCW limit switch";
            allFaults[0, 12] = "Resetting drive cabinet";
            allFaults[0, 13] = "Drives disabled at console";
            allFaults[0, 14] = "ACU Offline";
            allFaults[0, 15] = "Polarization #2 CW limit switch";
            allFaults[0, 16] = "Polarization #2 CCW limit switch";

            allFaults[1, 0] = "Azimuth CW soft limit";
            allFaults[1, 1] = "Azimuth CCW soft limit";
            allFaults[1, 2] = "Elevation upper soft limit";
            allFaults[1, 3] = "Elevation lower soft limit";
            allFaults[1, 4] = "Polarization2 CW soft limit";
            allFaults[1, 5] = "Polarization 2 CCW soft limit";
            allFaults[1, 6] = "West box limit violation";
            allFaults[1, 7] = "East box limit violation";
            allFaults[1, 8] = "North box limit violation";
            allFaults[1, 9] = "South box limit violation";
            allFaults[1, 10] = "Polarization 2 CW box limit violation";
            allFaults[1, 11] = "Polarization 2 CCW box limit violation";
            allFaults[1, 12] = "Azimuth immobile";
            allFaults[1, 13] = "Azimuth reversed";
            allFaults[1, 14] = "Azimuth runaway";
            allFaults[1, 15] = "Elevation immobile";
            allFaults[1, 16] = "Elevation reversed";
            allFaults[1, 17] = "Elevation runaway";
            allFaults[1, 18] = "Polarization2 immobile";
            allFaults[1, 19] = "Polarization2 reversed";
            allFaults[1, 20] = "Polarization2 runaway";
            allFaults[1, 21] = "Keyboard stop";
            allFaults[1, 22] = "Polarization #2 CW soft limit";
            allFaults[1, 23] = "Polarization #2 CCW soft limit";

            allFaults[2, 0] = "Target outside of soft limits";
            allFaults[2, 1] = "Low tracking signal level";
            allFaults[2, 2] = "Excessive tracking signal noise";
            allFaults[2, 3] = "Intelsat data expired";
            allFaults[2, 4] = "Intelsat pre-epoch prediction";
            allFaults[2, 5] = "Intelsat data invalid - cannot track";
            allFaults[2, 6] = "OPT cannot track";
            allFaults[2, 7] = "Tracking signal input saturated";
            allFaults[2, 8] = "Invalid target in target schedule";
            allFaults[2, 9] = "Sun outage; Steptrack inhibited";
            allFaults[2, 10] = "Tracking delay in effect";
            allFaults[2, 11] = "Reserved; unused";
            allFaults[2, 12] = "Reserved; unused";
            allFaults[2, 13] = "Reserved; unused";
            allFaults[2, 14] = "Standby (no tracking in progress)";
            allFaults[2, 15] = "Reserved; unused";
            allFaults[2, 16] = "Tracking receiver serial link failure";
            allFaults[2, 17] = "Tracking receiver in LOCAL control";
            allFaults[2, 18] = "Tracking receiver out of band";
            allFaults[2, 19] = "Tracking receiver fault";
            allFaults[2, 20] = "Table Track data expired";
            allFaults[2, 21] = "Orbital elements invalid - - cannot track";

            allFaults[3, 0] = "LB PROM checksum failure";
            allFaults[3, 1] = "HB PROM checksum failure";
            allFaults[3, 2] = "Azimuth (coarse) LOS";
            allFaults[3, 3] = "Azimuth (coarse) BIT failure";
            allFaults[3, 4] = "Azimuth (fine) LOS";
            allFaults[3, 5] = "Azimuth (fine) BIT failure";
            allFaults[3, 6] = "Elevation (coarse) LOS";
            allFaults[3, 7] = "Elevation (coarse) BIT failure";
            allFaults[3, 8] = "Elevation (fine) LOS";
            allFaults[3, 9] = "Elevation (fine) BIT failure";
            allFaults[3, 10] = "Polarization1 LOS";
            allFaults[3, 11] = "Polarization1 BIT failure";
            allFaults[3, 12] = "A/D 1 failure";
            allFaults[3, 13] = "A/D 2 failure";
            allFaults[3, 14] = "Unexpected exception";
            allFaults[3, 15] = "Sanity check failed";
            allFaults[3, 16] = "Non-volatile RAM corrupted";
            allFaults[3, 17] = "Watchdog timeout";
            allFaults[3, 18] = "Simulation";
            allFaults[3, 19] = "Azimuth encoder error";
            allFaults[3, 20] = "Elevation encoder error";
            allFaults[3, 21] = "Polarization1 encoder error";
            allFaults[3, 22] = "OUINTF task aborted";
            allFaults[3, 23] = "POSITIONER task aborted";
            allFaults[3, 24] = "TARGETER task aborted";
            allFaults[3, 25] = "SCHEDULER task aborted";
            allFaults[3, 26] = "Bus error on boot up";
            allFaults[3, 27] = "SYSFAIL line - timeout";
            allFaults[3, 28] = "SIMULATOR task aborted";
            allFaults[3, 29] = "Remote control panel link failure";
            allFaults[3, 30] = "System date/time invalid";

            //Globals.TargetAz = 180;
            //Globals.TargetAlt = 80;

            Globals.Status = "Connecting to the telescope";
            Globals.LoadProgress = 60;
            /*
            //try to connect to telescope
            while(ScopeInUse) { }
            ScopeInUse = true;
            ScopeSerialPort.DiscardInBuffer();
            ScopeSerialPort.Write("R0\r");

            string linein;

            try {
                linein = ScopeSerialPort.ReadLine();
                ScopeLinkLbl.Text = "Link: OK";
                ScopeLinkLbl.ForeColor = Color.Green;
            } catch {
                // uh oh
                ScopeLinkLbl.Text = "Link: Not found";
                ScopeLinkLbl.ForeColor = Color.Red;
            }
            ScopeInUse = false;
            */
            //try to connect to receiever
            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!D000");

            try {
                inbyte = RxSerialPort.ReadChar();
                string hexVal = "";
                for(i = 0; i < 3; i++) {
                    inbyte = RxSerialPort.ReadByte();
                    hexVal += Convert.ToChar(inbyte);
                }
                RxLinkLbl.Text = "Link: OK";
                RxLinkLbl.ForeColor = Color.Green;
            } catch {
                //uh oh
                RxLinkLbl.Text = "Link: Not found";
                RxLinkLbl.ForeColor = Color.Red;
            }

            RxInUse = false;

            //update 70MHz gain
            Gain70Update();

            Globals.Status = "Initialising sky map";
            Globals.LoadProgress = 70;

            //Initialise sky map
            Globals.LoadProgress += 5;

            thetaC = Math.PI;
            phiC = Math.PI / 2;

            hdc = (uint)pnlViewPort.Handle;

            string error = "";
            //OpenGLControl.OpenGLInit(ref hdc, pnlViewPort.Width, pnlViewPort.Height, ref error);

            if(error != "") {
                MessageBox.Show(error);
            }

            Globals.Status = "Loading sky map textures";
            Globals.LoadProgress = 80;

            //Load textures
            ContentManager.SetTextureList("textures/");
            ContentManager.LoadTextures();

            Gl.glClearColor(0, 0, 0, 1);

            //Setup celestial sphere
            Glu.GLUquadric quadratic = Glu.gluNewQuadric();
            Glu.gluQuadricNormals(quadratic, Glu.GLU_SMOOTH);
            Glu.gluQuadricTexture(quadratic, Gl.GL_TRUE);

            Globals.Status = "Reticulating splines";
            Globals.LoadProgress = 90;

            //various rotations and transformations of the celestial sphere
            list = Gl.glGenLists(1);
            Gl.glNewList(list, Gl.GL_COMPILE);
            Gl.glPushMatrix();
            Gl.glRotated(90, 1, 0, 0);
            Globals.LoadProgress += 5;
            Glu.gluSphere(quadratic, 1.01, 256, 256);

            Gl.glPopMatrix();
            Gl.glEndList();

            Globals.LoadProgress = 100;
            Globals.Status = "Done!";

            //select continuum tab
            tabControl1.SelectedIndex = 0;

            //enable sky map
            SkyMapTimer.Enabled = true;

            //make sure window is maximised
            this.WindowState = FormWindowState.Maximized;
        }

        private void ResetRx_Click(object sender, EventArgs e) {
            //Resets RX to default values. Does not then display them on the screen(!)
            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!R000");
            RxInUse = false;
        }

        void Gain70Update() {
            //change 70MHz gain

            string Gain = Convert.ToInt16((Gain70.Value - 10) * 4).ToString("X2");

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!A0" + Gain);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;
            //Properties.Settings.Default.Gain70 = Gain70.Value;
            //Properties.Settings.Default.Save();
        }

        void ContOffsetUpdate() {
            //update continuum offset

            string hexVal = Convert.ToInt16(ContOffset.Value).ToString("X3");
            int decValue = Convert.ToInt16(hexVal, 16);

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!O" + hexVal);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;
            //Properties.Settings.Default.ContOffset = ContOffset.Value;
            //Properties.Settings.Default.Save();

        }

        void ContGainUpdate() {
            //update continuum gain

            string text = ContGain.Text.ToString();
            string outVal = "0";

            //label3.Text = text;
            if(text == "x1") {
                outVal = "0";
            } else if(text == "x5") {
                outVal = "1";
            } else if(text == "x10") {
                outVal = "2";
            } else if(text == "x20") {
                outVal = "3";
            } else if(text == "x50") {
                outVal = "4";
            } else if(text == "x60") {
                outVal = "5";
            }

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!G00" + outVal);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;

            //Properties.Settings.Default.ContGain = text;
            //Properties.Settings.Default.Save();

        }

        void ContIntUpdate() {
            //update continuum integration time

            string text = ContInt.Text.ToString();
            string outVal = "0";

            if(text == "0.3s") {
                ContTimer.Interval = 300;
                outVal = "0";
            } else if(text == "1.0s") {
                ContTimer.Interval = 1000;
                outVal = "1";
            } else if(text == "10s") {
                ContTimer.Interval = 10000;
                outVal = "2";
            }

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!I00" + outVal);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;

            //Properties.Settings.Default.ContInt = text;
            //Properties.Settings.Default.Save();
        }
        
        public void SpecOffsetUpdate() {
            //label8.Text = Convert.ToInt16(SpecOffset.Value).ToString("X3");
            //int decValue = Convert.ToInt16(label8.Text, 16);
            //textBox1.Text = decValue.ToString();

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!J" + Convert.ToInt16(SpecOffset.Value).ToString("X3"));
            System.Threading.Thread.Sleep(50);
            RxInUse = false;
            //Properties.Settings.Default.SpecOffset2 = SpecOffset.Value;
            //Properties.Settings.Default.Save();
        }

        void SpecGainUpdate() {
            string text = SpecGain.Text.ToString();
            // label10.Text = text;

            string textout = "0";

            if(text == "x1") {
                textout = "0";
            } else if(text == "x5") {
                textout = "1";
            } else if(text == "x10") {
                textout = "2";
            } else if(text == "x20") {
                textout = "3";
            } else if(text == "x50") {
                textout = "4";
            } else if(text == "x60") {
                textout = "5";
            }

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!K00" + textout);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;

            //Properties.Settings.Default.SpecGain = text;
            // Properties.Settings.Default.Save();
        }

        void SpecIntUpdate() {

            string text = SpecInt.Text.ToString();
            // label9.Text = text;
            string textout = "0";
            if(text == "0.3s") {
                textout = "0";
                sChanTimer.Interval = 350;
                SpectrumTimer.Interval = 350;

            } else if(text == "0.5s") {
                textout = "1";
                sChanTimer.Interval = 550;
                SpectrumTimer.Interval = 550;
            } else if(text == "1.0s") {
                textout = "2";
                sChanTimer.Interval = 1050;
                SpectrumTimer.Interval = 1050;
            }

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!L00" + textout);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;

            //Properties.Settings.Default.SpecInt = text;
            //Properties.Settings.Default.Save();
        }

        int ContRead() {
            //Take a continuum reading

            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!D000");

            inbyte = RxSerialPort.ReadByte();
            string hexVal = "";

            for(i = 0; i < 3; i++) {
                inbyte = RxSerialPort.ReadByte();
                hexVal += Convert.ToChar(inbyte);
            }

            RxInUse = false;

            // 0-4095
            return Convert.ToInt16(hexVal, 16);
        }

        int SpecRead() {
            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!D001");
            //SpecReading.Text = "";
            inbyte = RxSerialPort.ReadByte();
            string hexVal = "";
            for(i = 0; i < 3; i++) {
                inbyte = RxSerialPort.ReadByte();
                hexVal += Convert.ToChar(inbyte);
                //SpecReading.AppendText(Convert.ToChar(inbyte).ToString());
            }

            RxInUse = false;

            return Convert.ToInt16(hexVal, 16);
        }

        void SpecChannelUpdate() {
            //change spectrograph channel

            int nn = Convert.ToInt16(chanUpDown.Value);
            string chan = (nn + 80).ToString("X3");
            //label13.Text = chan;
            while(RxInUse) { }
            RxInUse = true;
            RxSerialPort.DiscardInBuffer();
            RxSerialPort.Write("!F" + chan);
            System.Threading.Thread.Sleep(50);
            RxInUse = false;

            double vel = (440 - Convert.ToDouble(nn)) * 1.055332251 + 0.5;

            label7.Text = vel.ToString("F3") + " km/s";

        }

        private void NoiseSource_CheckedChanged(object sender, EventArgs e) {
            //Toggle noise source

            if(NoiseSource.Checked) {
                while(RxInUse) { }
                RxInUse = true;
                RxSerialPort.DiscardInBuffer();
                RxSerialPort.Write("!N001");
                RxInUse = false;
            } else {
                while(RxInUse) { }
                RxInUse = true;
                RxSerialPort.DiscardInBuffer();
                RxSerialPort.Write("!N000");
                RxInUse = false;
            }
        }

        private void ContTimer_Tick(object sender, EventArgs e) {
            //Continuum timer

            //get a reading
            int reading = ContRead();

            //set meters
            ContBar.Value = reading;
            //ContTrack.Value = reading;

            label1.Text = reading.ToString() + " mV";

            //reet max/min
            max = 0;
            min = 4096;

            //shift array and get max min value
            for(n = 0; n < 599; n++) {
                ContArray[n] = ContArray[n + 1];
                if(ContArray[n] > max && ContArray[n] < 4097) { max = ContArray[n]; }
                if(ContArray[n] < min) { min = ContArray[n]; }
            }

            //add last reading to end of array
            ContArray[599] = reading;

            //check if last reading is min/max
            if(ContArray[599] > max) { max = ContArray[599]; }
            if(ContArray[599] < min) { min = ContArray[599]; }

            //clear chart and series
            ContChart.Series.Remove(ContSeries);
            ContSeries.Points.Clear();

            //add array to series
            for(n = 0; n < 600; n++) {
                ContSeries.Points.AddXY(600 - n, ContArray[n]);
            }

            //set chart axis
            ContChart.ChartAreas[0].AxisY.Minimum = Math.Floor((double)min / 10) * 10 - 10;
            if(ContChart.ChartAreas[0].AxisY.Minimum < 0) { ContChart.ChartAreas[0].AxisY.Minimum = 0; }
            ContChart.ChartAreas[0].AxisY.Maximum = Math.Ceiling((double)max / 10) * 10 + 10;
            double range = ContChart.ChartAreas[0].AxisY.Maximum - ContChart.ChartAreas[0].AxisY.Minimum;
            ContChart.ChartAreas[0].AxisY.Interval = range / 10;

            //add series to chart
            ContChart.Series.Add(ContSeries);

            //save to file if needed
            if(saveName != "") {
                File.AppendAllText(saveName, DateTime.Now.ToString("HH:mm:ss") + "," + reading.ToString() + Environment.NewLine);
            }

        }

        private void sChanTimer_Tick(object sender, EventArgs e) {

            int reading = SpecRead();
            //SpecReading.AppendText(" - " + reading.ToString() + Environment.NewLine);
            SpecBar.Value = reading;
            label2.Text = reading.ToString() + " mV";

            int max = 0;
            int min = 4096;

            for(n = 0; n < 299; n++) {
                sChanArray[n] = sChanArray[n + 1];
                if(sChanArray[n] > max && sChanArray[n] < 4096) { max = sChanArray[n]; }
                if(sChanArray[n] < min) { min = sChanArray[n]; }
            }

            sChanArray[299] = reading;
            if(sChanArray[299] > max) { max = sChanArray[299]; }
            if(sChanArray[299] < min) { min = sChanArray[299]; }

            sChanChart.Series.Remove(sChanSeries);
            sChanSeries.Points.Clear();

            for(n = 0; n < 300; n++) {
                sChanSeries.Points.AddXY(300 - n, sChanArray[n]);
            }
            sChanChart.ChartAreas[0].AxisY.Minimum = Math.Floor((double)min / 10) * 10 - 10;
            if(sChanChart.ChartAreas[0].AxisY.Minimum < 0) { sChanChart.ChartAreas[0].AxisY.Minimum = 0; }
            sChanChart.ChartAreas[0].AxisY.Maximum = Math.Ceiling((double)max / 10) * 10 + 10;
            double range = sChanChart.ChartAreas[0].AxisY.Maximum - sChanChart.ChartAreas[0].AxisY.Minimum;

            sChanChart.ChartAreas[0].AxisY.Interval = range / 10;
            sChanChart.Series.Add(sChanSeries);

            if(saveName != "") {
                File.AppendAllText(saveName, DateTime.Now.ToString("HH:mm:ss") + "," + reading.ToString() + Environment.NewLine);
            }

        }
        
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) {
            //tab has been changed


            if(tabControl1.SelectedIndex == 1) {
                //single channel

                //stop continuum graph
                ContTimer.Enabled = false;
                ContStartStop.Text = "Start";

                //send settings to rx
                SpecIntUpdate();
                SpecGainUpdate();
                SpecOffsetUpdate();
                SpecChannelUpdate();

                //Flush integration capacitor
                Thread.Sleep(50);
                SpecRead();
                Thread.Sleep(1000);
                SpecRead();
                Thread.Sleep(50);

                //SpecTimer.Enabled = true;

            } else if(tabControl1.SelectedIndex == 2) {
                //spectrum

                //stop continuum and single channel timers
                ContTimer.Enabled = false;
                sChanTimer.Enabled = false;
                sChanStartStop.Text = "Start";
                ContStartStop.Text = "Start";

                //update spectrum settings
                SpecIntUpdate();
                SpecGainUpdate();
                SpecOffsetUpdate();

                //update chart axis
                SpecAxisUpdate();


            } else if(tabControl1.SelectedIndex == 0) {
                //continuum

                //stop single channel timer
                sChanTimer.Enabled = false;
                sChanStartStop.Text = "Start";

                //update continuum settings
                ContIntUpdate();
                ContGainUpdate();
                ContOffsetUpdate();

                //Flush integration capacitor
                Thread.Sleep(50);
                ContRead();
                Thread.Sleep(1000);
                ContRead();
                Thread.Sleep(50);

                //ContTimer.Enabled = true;
            }

        }

        private void b15k_CheckedChanged(object sender, EventArgs e) {
            //Alledgedly changed the bandwidth, doesn't actually seem to do anything

            if(b15k.Checked == true) {
                while(RxInUse) { }
                RxInUse = true;
                RxSerialPort.DiscardInBuffer();
                RxSerialPort.Write("!B000");
                System.Threading.Thread.Sleep(50);
                RxInUse = false;

                //Properties.Settings.Default.bwidth = 15;
                // Properties.Settings.Default.Save();
            } else {
                while(RxInUse) { }
                RxInUse = true;
                RxSerialPort.DiscardInBuffer();
                RxSerialPort.Write("!B001");
                System.Threading.Thread.Sleep(50);
                RxInUse = false;

                //Properties.Settings.Default.bwidth = 30;
                // Properties.Settings.Default.Save();
            }
        }

        private void SpectrumScan_Click(object sender, EventArgs e) {
            //Start or stop a spectrum scan


            if(SpectrumScan.Text == "Stop") {
                //set stop flag
                stopscan = true;

            } else {
                //start a scan

                //ask for a name to save data too
                saveName = "";
                SaveData.FileName = "l" + ScopeGal.l.ToDeg().ToString("F0") + "b" + ScopeGal.b.ToDeg().ToString("F0") + ".csv";
                SaveData.ShowDialog();

                //set initial info
                if(saveName != "") {
                    File.AppendAllText(saveName, DateTime.Now.ToString() + Environment.NewLine);
                }

                stopscan = false;

                //wait for rx
                while(RxInUse) { }
                RxInUse = true;
                SpectrumScan.Text = "Stop";

                //change to start channel
                n = Convert.ToInt16(start.Value);
                string chan = (n + 80).ToString("X3");
                RxSerialPort.Write("!F" + chan);

                //wait for capacitor
                System.Threading.Thread.Sleep(2000);

                //clear all points
                SpectrumSeriesR.Points.Clear();
                SpectrumSeriesB.Points.Clear();

                //reset max/min
                max = 0;
                min = 4096;

                //repeat count to 1
                spectrumLoops = 1;

                //reset arrays
                for(i = 0; i < 441; i++) {
                    SpectrumArrayRN[i] = 0;
                    SpectrumArrayRSum[i] = 0;
                    SpectrumArrayBN[i] = 0;
                    SpectrumArrayBSum[i] = 0;
                }

                //set the timer going
                SpectrumTimer.Enabled = true;

            }
        }

        private void SpectrumTimer_Tick(object sender, EventArgs e) {
            //spectrum scan timer

            //convert channel to velocity
            double vel = (440 - Convert.ToDouble(n)) * 1.055332251 + 0.5;

            //calculate channel number from n
            string chan = (n + 80).ToString("X3");

            //set repeat and current channel labels
            repeatLbl.Text = spectrumLoops.ToString() + "/" + spectrumRepeats.Value.ToString();
            currentChanLbl.Text = n.ToString() + " - " + vel.ToString("F3") + " km/s";

            //change channel
            RxSerialPort.Write("!F" + chan);
            //calculate wait time
            int t = SpectrumTimer.Interval - 50;
            //minimum wait 50ms
            if(t < 50) { t = 50; }
            //wait for capacitor
            System.Threading.Thread.Sleep(t);
            RxSerialPort.DiscardInBuffer();
            //request reading
            RxSerialPort.Write("!D001");
            //SpecReading.Text = "";
            //read
            inbyte = RxSerialPort.ReadByte();
            string hexVal = "";
            //convert to proper numbers
            for(i = 0; i < 3; i++) {
                inbyte = RxSerialPort.ReadByte();
                hexVal += Convert.ToChar(inbyte);
            }
            int reading = Convert.ToInt16(hexVal, 16);

            //save to file if name is set
            if(saveName != "") {
                File.AppendAllText(saveName, n + "," + vel.ToString("F2") + "," + DateTime.Now.ToString("HH:mm:ss") + "," + reading.ToString() + Environment.NewLine);
            }

            //get max/min
            if(reading > max) { max = reading; }
            if(reading < min) { min = reading; }

            //set axis minimum
            Spectrum.ChartAreas[0].AxisY.Minimum = Math.Floor((double)min / 10) * 10 - 10;

            //make sure it's not <0
            if(Spectrum.ChartAreas[0].AxisY.Minimum < 0) {
                Spectrum.ChartAreas[0].AxisY.Minimum = 0;
            }
            //set axis maximum
            Spectrum.ChartAreas[0].AxisY.Maximum = Math.Ceiling((double)max / 10) * 10 + 10;

            //set tick interval
            double range = Spectrum.ChartAreas[0].AxisY.Maximum - Spectrum.ChartAreas[0].AxisY.Minimum;
            Spectrum.ChartAreas[0].AxisY.Interval = range / 10;

            if(n == 440) {
                //special case where red shift and blue shift join

                SpectrumArrayRN[n]++;
                SpectrumArrayRSum[n] += reading;

                SpectrumArrayBN[n - 440]++;
                SpectrumArrayBSum[n - 440] += reading;

            } else if(n < 440) {
                //red shifted

                SpectrumArrayRN[n]++;
                SpectrumArrayRSum[n] += reading;

            } else {
                //blue shifted

                SpectrumArrayBN[n - 440]++;
                SpectrumArrayBSum[n - 440] += reading;
            }

            //clear both points
            SpectrumSeriesR.Points.Clear();
            SpectrumSeriesB.Points.Clear();

            //add points from arrays
            for(i = 0; i < 441; i++) {
                if(SpectrumArrayRN[i] != 0) {
                    SpectrumArrayRAv[i] = Convert.ToDouble(SpectrumArrayRSum[i]) / Convert.ToDouble(SpectrumArrayRN[i]);
                    SpectrumSeriesR.Points.AddXY(SpectrumRVel[i], SpectrumArrayRAv[i]);
                }

                if(SpectrumArrayBN[i] != 0) {
                    SpectrumArrayBAv[i] = Convert.ToDouble(SpectrumArrayBSum[i]) / Convert.ToDouble(SpectrumArrayBN[i]);
                    SpectrumSeriesB.Points.AddXY(SpectrumBVel[i], SpectrumArrayBAv[i]);
                }
            }

            //add points to series
            Spectrum.Series.Clear();
            Spectrum.Series.Add(SpectrumSeriesR);
            Spectrum.Series.Add(SpectrumSeriesB);

            //next channel
            n = n + steps;

            //if at last channel, add repeat counter
            if(n > Convert.ToInt16(stop.Value)) {
                n = Convert.ToInt16(start.Value);
                spectrumLoops++;
            }

            //if done all repeats, or asked to stop... stop
            if(spectrumLoops > Convert.ToInt16(spectrumRepeats.Value) || stopscan) {

                SpectrumTimer.Enabled = false;

                RxInUse = false;

                SpecChannelUpdate();

                SpectrumScan.Text = "Scan";

            }
        }

        void SpecAxisUpdate() {
            //update spectrograph x and y axes

            //convert start channel to velocity
            double startvel = (440 - Convert.ToDouble(start.Value)) * 1.055332251 + 0.5;

            //convert stop channel to velocity
            double stopvel = (440 - Convert.ToDouble(stop.Value)) * 1.055332251 + 0.5;

            steps = Convert.ToInt16(NumSteps.Value);

            label15.Text = ((stop.Value - start.Value) / steps).ToString("F1") + " steps";

            //display start/stop velocity
            label18.Text = startvel.ToString("F3") + " km/s";
            label19.Text = stopvel.ToString("F3") + " km/s";

            /*
            if(startvel < 0) {
                Spectrum.ChartAreas[0].AxisX.Minimum = 0;
            } else {

                if(Math.Abs(startvel) > Math.Abs(stopvel)) {
                    Spectrum.ChartAreas[0].AxisX.Maximum = Math.Floor(Math.Abs(startvel) / 50) * 50 + 50;
                    Spectrum.ChartAreas[0].AxisX.Minimum = Math.Floor((0 - Math.Abs(startvel)) / 50) * 50;
                } else {
                    Spectrum.ChartAreas[0].AxisX.Maximum = Math.Floor(Math.Abs(stopvel) / 50) * 50 + 50;
                    Spectrum.ChartAreas[0].AxisX.Minimum = Math.Floor((0 - Math.Abs(stopvel)) / 50) * 50;
                }
            }

            if(stopvel > 0) {
                Spectrum.ChartAreas[0].AxisX.Maximum = 0;
            } else {
                if(Math.Abs(startvel) > Math.Abs(stopvel)) {
                    Spectrum.ChartAreas[0].AxisX.Maximum = Math.Floor(Math.Abs(startvel) / 50) * 50 + 50;
                    Spectrum.ChartAreas[0].AxisX.Minimum = Math.Floor((0 - Math.Abs(startvel)) / 50) * 50;
                } else {
                    Spectrum.ChartAreas[0].AxisX.Maximum = Math.Floor(Math.Abs(stopvel) / 50) * 50 + 50;
                    Spectrum.ChartAreas[0].AxisX.Minimum = Math.Floor((0 - Math.Abs(stopvel)) / 50) * 50;
                }
            }
             */

            Spectrum.ChartAreas[0].AxisX.Maximum = Math.Floor(startvel / 50) * 50 + 50;
            Spectrum.ChartAreas[0].AxisX.Minimum = Math.Floor(stopvel / 50) * 50;

        }

        //----------------------------------------------------------------------------------------------------------
        //Telescope control region

        private void PositionRead_Tick(object sender, EventArgs e) {
            //ScopePos();
        }

        private void ScopePos() {
            //Read the telescope position

            double JD = Utils.jd();
            double T = Utils.JDtoT(JD);

            if(AltAzRB.Checked) {
                //Get alt az input position

                TargetHor.az = Convert.ToDouble(AzUD.Value).ToRad();
                TargetHor.alt = Convert.ToDouble(AltUD.Value).ToRad();

            } else if(eqRB.Checked) {

                //get equatorial input position

                double ra = Convert.ToDouble((RAhUD.Value + (RAmUD.Value / 60) + (RAsUD.Value / 3600)) * 15).ToRad();
                double dec = Convert.ToDouble(DecdUD.Value + (DecmUD.Value / 60) + (DecsUD.Value / 3600)).ToRad();

                EquatorialCoords eqcoords = new EquatorialCoords(ra, dec);

                EquatorialCoords eqnow = eqcoords.J2000ToNow(JD);

                //convert to alt az
                TargetHor = eqnow.ToHorizontal(location, T);

            } else if(GalRB.Checked) {

                //get galactic input position

                double l = Convert.ToDouble(ldUD.Value + (lmUD.Value / 60)).ToRad();
                double b = Convert.ToDouble(bdUD.Value + (bmUD.Value / 60)).ToRad();

                GalacticCoords galcoords = new GalacticCoords(l, b);

                //convert to b1950 equatorial
                EquatorialCoords b1950 = galcoords.ToB1950();

                //convert to j2000 equatorial
                EquatorialCoords eqnow = b1950.B1950ToNow(JD);

                //convert to alt az
                TargetHor = eqnow.ToHorizontal(location, T);

            } else if(sourceRB.Checked) {

                //named source from drop down box

                //-1 none selected
                //0 = sun
                //1 = moon

                if(sourceBox.SelectedIndex > 1 && sourceBox.SelectedIndex < NumSources) {

                    //outside solar system
                    EquatorialCoords eqcoords = new EquatorialCoords(sourceRAarray[sourceBox.SelectedIndex], sourceDecarray[sourceBox.SelectedIndex]);

                    EquatorialCoords eqnow = eqcoords.J2000ToNow(JD);

                    TargetHor = eqnow.ToHorizontal(location, T);

                } else if(sourceBox.SelectedIndex == 0) {

                    //sun
                    TargetHor = SunHor;

                } else if(sourceBox.SelectedIndex == 1) {

                    //moon
                    TargetHor = MoonHor;

                }

            }

            //output target alt az
            TargetAzLbl.Text = "Azimuth: " + TargetHor.az.ToDeg().ToString("F");
            TargetAltLbl.Text = "Elevation: " + TargetHor.alt.ToDeg().ToString("F");

            //don't allow to goto if alt <10 degrees
            if(TargetHor.alt.ToDeg() < 10) {
                TargetAltLbl.ForeColor = Color.Red;
                if(gotoBtn.Enabled) {
                    gotoBtn.Enabled = false;
                }
                if(trackBtn.Enabled) {
                    trackBtn.Enabled = false;
                }
            } else {
                TargetAltLbl.ForeColor = Color.Black;
                if(!gotoBtn.Enabled) {
                    gotoBtn.Enabled = true;
                }
                if(!trackBtn.Enabled) {
                    trackBtn.Enabled = true;
                }
            }

            while(ScopeInUse) { }
            ScopeInUse = true;

            double t = Utils.JDtot(JD);
            NutObl nutobl = new NutObl(T);

            ScopeSerialPort.DiscardInBuffer();
            //general status request
            ScopeSerialPort.Write("R7 0\r");

            string linein;

            try {
                linein = ScopeSerialPort.ReadLine();
            } catch {
                linein = "TIMEOUT";
            }

            ScopeInUse = false;

            //0 1 date time  01/19/2015 14:47:05 
            //2 target (R300) 0 
            //3 target mode (R11) 1 
            //4 target submode (RC) 3 
            //5 summary fault (R2) 002 (bit 0 = drive cabinet fault, 1 = motion fault, 2 = tracking fault, 3 = system error)
            //6 signal level (R10) -16.0 
            //7 8 9 position (R1 0) 179.99 82.00 152.8
            //  \r000

            try {
                string[] words = linein.Split(' ');
                ScopeHor.az = Convert.ToDouble(words[7]).ToRad();
                ScopeHor.alt = Convert.ToDouble(words[8]).ToRad();

                ScopeAzLbl.Text = "Azimuth: " + ScopeHor.az.ToDeg().ToString("F");
                ScopeAltLbl.Text = "Elevation: " + ScopeHor.alt.ToDeg().ToString("F");

                int faults = Convert.ToInt16(words[5]);

                if(faults > 0) {

                    while(ScopeInUse) { }
                    ScopeInUse = true;

                    ScopeSerialPort.DiscardInBuffer();

                    ScopeSerialPort.Write("R3 0\r");
                    string errors = ScopeSerialPort.ReadLine();

                    string[] uhohs = errors.Split(' ');

                    Int32[] faultList = new Int32[4];

                    int n;

                    errorsLbl.Text = "Errors: ";

                    for(k = 0; k < 4; k++) {
                        n = 1;
                        for(i = 7; i >= 0; i--) {
                            // richTextBox1.AppendText(uhohs[k] + " ");
                            faultList[k] += Convert.ToInt32(Regex.Replace(uhohs[k], @"\s+", "").Substring(i, 1)) * n;
                            n = n * 16;
                        }

                        for(i = 0; i < faultNums[k]; i++) {
                            if(bitSet(faultList[k], i)) {
                                errorsLbl.Text += allFaults[k, i] + Environment.NewLine;
                            }
                        }

                    }

                } else {
                    errorsLbl.Text = "No errors";
                }

            } catch(Exception ex) {

                ScopeAzLbl.Text = "Azimuth: " + ex.Message;
                ScopeAltLbl.Text = "Elevation: " + ex.Message;

            }

            ScopeInUse = false;

            ScopeEq = ScopeHor.ToEquatorial(location, T);

            EquatorialCoords ScopeEqb1950 = ScopeEq.J2000ToB1950();

            ScopeEq.NowToJ2000(JD);

            ScopeGal = ScopeEqb1950.B1950ToGalactic();

            CoordsString EqString = ScopeEq.ToDMS();
            CoordsString GalString = ScopeGal.ToDMS();

            ScopeRALbl.Text = " RA: " + EqString.ra;
            ScopeDecLbl.Text = "Dec: " + EqString.dec;

            galacticLlbl.Text = "l: " + GalString.l;
            galacticBlbl.Text = "b: " + GalString.b;

            //distance between scope and target
            ScopetoTarget = (Math.Acos(Math.Sin(ScopeHor.alt) * Math.Sin(TargetHor.alt) + Math.Cos(ScopeHor.alt) * Math.Cos(TargetHor.alt) * Math.Cos(ScopeHor.az - TargetHor.az))).ToDeg();

            //estimate time to sloew
            double aztime;
            double alttime;

            //individual distances
            double AzDist = Math.Abs(TargetHor.az - ScopeHor.az).ToDeg();
            double AltDist = Math.Abs(TargetHor.alt - ScopeHor.alt).ToDeg();

            if(AzDist > 2) {
                aztime = AzDist * 1.36 + 9.28;
            } else {
                aztime = AzDist * 6;
            }

            if(AltDist > 2) {
                alttime = AltDist * 1.266 + 7.4;
            } else {
                alttime = AltDist * 4.966;
            }

            //slew time is the longest one
            double slewtime;
            if(alttime > aztime) {
                slewtime = alttime;
            } else {
                slewtime = aztime;
            }

            string slewtimestr = "";
            if(ScopetoTarget > 0.03) {
                slewtimestr = " (" + Math.Round(slewtime, 0).ToString() + "s)";

                if(ScopetoTarget > 0.03 && tracking) {
                    SlewScope();
                }
            }

            Distlbl.Text = "Distance: " + ScopetoTarget.ToString("F") + slewtimestr;

        }

        private bool SlewScope() {
            //Ask the telescope to move

            //check if it's within limits, ought to have some notification here really.

            if(TargetHor.az.ToDeg() < 0.5 || TargetHor.az.ToDeg() > 359.5 || TargetHor.alt.ToDeg() < 10.0 || TargetHor.alt.ToDeg() > 82.0) {
                return false;
            } else {

                //calc distance to target
                ScopetoTarget = (Math.Acos(Math.Sin(ScopeHor.alt) * Math.Sin(TargetHor.alt) + Math.Cos(ScopeHor.alt) * Math.Cos(TargetHor.alt) * Math.Cos(ScopeHor.az - TargetHor.az))).ToDeg();

                //set target alt az labels
                TargetAzLbl.Text = "Azimuth: " + TargetHor.az.ToDeg().ToString("F");
                TargetAltLbl.Text = "Elevation: " + TargetHor.alt.ToDeg().ToString("F");

                //if it's close, don't bother
                if(ScopetoTarget > 0.02) {
                    while(ScopeInUse) { }
                    ScopeInUse = true;
                    ScopeSerialPort.DiscardInBuffer();
                    ScopeSerialPort.Write("C2 " + TargetHor.az.ToDeg().ToString("F") + " " + TargetHor.alt.ToDeg().ToString("F") + " 0.0\r");
                    ScopeInUse = false;
                }
                return true;
            }
        }

        void calcTarget(object sender, EventArgs e) {

            if(eqRB.Checked) {

                if(RAsUD.Value == 60) {
                    RAsUD.Value = 0;
                    RAmUD.Value++;
                } else if(RAsUD.Value == -1) {
                    RAsUD.Value = 59;
                    RAmUD.Value--;
                }

                if(RAmUD.Value == 60) {
                    RAmUD.Value = 0;
                    RAhUD.Value++;
                } else if(RAmUD.Value == -1) {
                    RAmUD.Value = 59;
                    RAhUD.Value--;
                }

                if(RAhUD.Value == 24) {
                    RAhUD.Value = 0;
                } else if(RAhUD.Value == -1) {
                    RAhUD.Value = 23;
                }

                if(DecsUD.Value == 60) {
                    DecsUD.Value = 0;
                    DecmUD.Value++;
                } else if(DecsUD.Value == -1) {
                    DecsUD.Value = 59;
                    DecmUD.Value--;
                }

                if(DecmUD.Value == 60) {
                    DecmUD.Value = 0;
                    DecdUD.Value++;
                } else if(DecmUD.Value == -1) {
                    DecmUD.Value = 59;
                    DecdUD.Value--;
                }

            } else if(GalRB.Checked) {

                if(lmUD.Value == 60) {
                    lmUD.Value = 0;
                    ldUD.Value++;
                } else if(lmUD.Value == -1) {
                    lmUD.Value = 59;
                    ldUD.Value--;
                }

                if(ldUD.Value == 360) {
                    ldUD.Value = 0;
                } else if(ldUD.Value == -1) {
                    ldUD.Value = 359;
                }

                if(bmUD.Value == 60) {
                    bmUD.Value = 0;
                    bdUD.Value++;
                } else if(bmUD.Value == -1) {
                    bmUD.Value = 59;
                    bdUD.Value--;
                }
            }
        }

        void targetRBchanged(object sender, EventArgs e) {

            if(tracking) {
                tracking = false;
                trackBtn.Text = "Track";
            }

            if(AltAzRB.Checked) {

                RAhUD.Enabled = false;
                RAmUD.Enabled = false;
                RAsUD.Enabled = false;
                DecdUD.Enabled = false;
                DecmUD.Enabled = false;
                DecsUD.Enabled = false;

                ldUD.Enabled = false;
                lmUD.Enabled = false;
                bdUD.Enabled = false;
                bmUD.Enabled = false;

                sourceBox.Enabled = false;

                AzUD.Enabled = true;
                AltUD.Enabled = true;

            } else if(eqRB.Checked) {

                RAhUD.Enabled = true;
                RAmUD.Enabled = true;
                RAsUD.Enabled = true;
                DecdUD.Enabled = true;
                DecmUD.Enabled = true;
                DecsUD.Enabled = true;

                ldUD.Enabled = false;
                lmUD.Enabled = false;
                bdUD.Enabled = false;
                bmUD.Enabled = false;

                sourceBox.Enabled = false;

                AzUD.Enabled = false;
                AltUD.Enabled = false;

            } else if(GalRB.Checked) {

                RAhUD.Enabled = false;
                RAmUD.Enabled = false;
                RAsUD.Enabled = false;
                DecdUD.Enabled = false;
                DecmUD.Enabled = false;
                DecsUD.Enabled = false;

                ldUD.Enabled = true;
                lmUD.Enabled = true;
                bdUD.Enabled = true;
                bmUD.Enabled = true;

                sourceBox.Enabled = false;

                AzUD.Enabled = false;
                AltUD.Enabled = false;

            } else if(sourceRB.Checked) {

                RAhUD.Enabled = false;
                RAmUD.Enabled = false;
                RAsUD.Enabled = false;
                DecdUD.Enabled = false;
                DecmUD.Enabled = false;
                DecsUD.Enabled = false;

                ldUD.Enabled = false;
                lmUD.Enabled = false;
                bdUD.Enabled = false;
                bmUD.Enabled = false;

                sourceBox.Enabled = true;

                AzUD.Enabled = false;
                AltUD.Enabled = false;

            }

        }

        //This got a bit messy so I disabled it for now
        /*
  private void Scan_Click(object sender, EventArgs e) {

      PositionRead.Enabled = false;
      ContTimer.Enabled = false;

      for (Globals.TargetAz = 308; Globals.TargetAz >= 298; Globals.TargetAz = Globals.TargetAz - 0.5) {
          if ((Math.Floor(Globals.TargetAz) % 1) == 0) {
              for (Globals.TargetAlt = 70; Globals.TargetAlt >= 60; Globals.TargetAlt = Globals.TargetAlt - 0.5) {

                  TargetAzLbl.Text = "Azimuth: " + Globals.TargetAz.ToString("F");
                  TargetAltLbl.Text = "Elevation: " + Globals.TargetAlt.ToString("F");

                  double AzR1 = Globals.TargetAz);
                  double AzR2 = ScopeAz);
                  double ElR1 = Globals.TargetAlt);
                  double ElR2 = ScopeAlt);
                  ScopetoTarget = rad2deg(Math.Acos(Math.Sin(ElR1) * Math.Sin(ElR2) + Math.Cos(ElR1) * Math.Cos(ElR2) * Math.Cos(AzR1 - AzR2)));
                  Distlbl.Text = "Distance: " + ScopetoTarget.ToString("F");

                  SlewScope();

                  while (ScopetoTarget > 0.04) {

                      while (ScopeInUse) { }
                      ScopeInUse = true;
                      ScopeSerialPort.DiscardInBuffer();
                      ScopeSerialPort.Write("R7 0\r");

                      string linein;
                      try {
                          linein = ScopeSerialPort.ReadLine();

                      } catch {
                          linein = "TIMEOUT";
                      }
                      //richTextBox1.AppendText(linein);
                      ScopeInUse = false;

                      try {
                          string[] words = linein.Split(' ');
                          ScopeAz = Convert.ToDouble(words[7]);
                          ScopeAlt = Convert.ToDouble(words[8]);
                          ScopeAzLbl.Text = "Azimuth: " + ScopeAz.ToString("F");
                          ScopeAltLbl.Text = "Elevation: " + ScopeAlt.ToString("F");

                      } catch {
                          ScopeAzLbl.Text = "Azimuth: ERROR";
                          ScopeAltLbl.Text = "Elevation:  ERROR";
                      }

                      AzR1 = Globals.TargetAz);
                      AzR2 = ScopeAz);
                      ElR1 = Globals.TargetAlt);
                      ElR2 = ScopeAlt);
                      ScopetoTarget = rad2deg(Math.Acos(Math.Sin(ElR1) * Math.Sin(ElR2) + Math.Cos(ElR1) * Math.Cos(ElR2) * Math.Cos(AzR1 - AzR2)));

                      Distlbl.Text = "Distance: " + ScopetoTarget.ToString("F");
                      Application.DoEvents();
                      Thread.Sleep(250);
                  }
        
                  while (RxInUse) { }
                  RxInUse = true;
                  RxSerialPort.DiscardInBuffer();
                  RxSerialPort.Write("!D000");
                  ContReading.Text = "";
                  inbyte = RxSerialPort.ReadByte();
                  string hexVal = "";
                  for (i = 0; i < 3; i++) {
                      inbyte = RxSerialPort.ReadByte();
                      hexVal += Convert.ToChar(inbyte);
                      ContReading.AppendText(Convert.ToChar(inbyte).ToString());
                  }

                  RxInUse = false;

                  int reading = Convert.ToInt16(hexVal, 16);
                  richTextBox1.AppendText(ScopeAz.ToString("F") + " " + ScopeAlt.ToString("F") + " " + reading.ToString() + Environment.NewLine);

                  var objWriter = new System.IO.StreamWriter("C:\\Program Files\\Apache Software Foundation\\Apache2.2\\htdocs\\forrest\\scan51.txt", true);
                  objWriter.WriteLine(ScopeAz.ToString("F") + " " + ScopeAlt.ToString("F") + " " + reading.ToString());
                  objWriter.Close();

              }

          } else {

              for (Globals.TargetAlt = 60; Globals.TargetAlt <= 70; Globals.TargetAlt = Globals.TargetAlt + 0.5) {

                  TargetAzLbl.Text = "Azimuth: " + Globals.TargetAz.ToString("F");
                  TargetAltLbl.Text = "Elevation: " + Globals.TargetAlt.ToString("F");

                  double AzR1 = Globals.TargetAz);
                  double AzR2 = ScopeAz);
                  double ElR1 = Globals.TargetAlt);
                  double ElR2 = ScopeAlt);
                  ScopetoTarget = rad2deg(Math.Acos(Math.Sin(ElR1) * Math.Sin(ElR2) + Math.Cos(ElR1) * Math.Cos(ElR2) * Math.Cos(AzR1 - AzR2)));
                  Distlbl.Text = "Distance: " + ScopetoTarget.ToString("F");

                  SlewScope();

                  while (ScopetoTarget > 0.04) {

                      while (ScopeInUse) { }
                      ScopeInUse = true;
                      ScopeSerialPort.DiscardInBuffer();
                      ScopeSerialPort.Write("R7 0\r");

                      string linein;
                      try {
                          linein = ScopeSerialPort.ReadLine();

                      } catch {
                          linein = "TIMEOUT";
                      }
                      //richTextBox1.AppendText(linein);
                      ScopeInUse = false;

                      try {
                          string[] words = linein.Split(' ');
                          ScopeAz = Convert.ToDouble(words[7]);
                          ScopeAlt = Convert.ToDouble(words[8]);
                          ScopeAzLbl.Text = "Azimuth: " + ScopeAz.ToString("F");
                          ScopeAltLbl.Text = "Elevation: " + ScopeAlt.ToString("F");

                      } catch {
                          ScopeAzLbl.Text = "Azimuth: ERROR";
                          ScopeAltLbl.Text = "Elevation:  ERROR";
                      }

                      AzR1 = Globals.TargetAz);
                      AzR2 = ScopeAz);
                      ElR1 = Globals.TargetAlt);
                      ElR2 = ScopeAlt);
                      ScopetoTarget = rad2deg(Math.Acos(Math.Sin(ElR1) * Math.Sin(ElR2) + Math.Cos(ElR1) * Math.Cos(ElR2) * Math.Cos(AzR1 - AzR2)));

                      Distlbl.Text = "Distance: " + ScopetoTarget.ToString("F");
                      Application.DoEvents();
                      Thread.Sleep(250);
                  }

                  while (RxInUse) { }
                  RxInUse = true;
                  RxSerialPort.DiscardInBuffer();
                  RxSerialPort.Write("!D000");
                  ContReading.Text = "";
                  inbyte = RxSerialPort.ReadByte();
                  string hexVal = "";
                  for (i = 0; i < 3; i++) {
                      inbyte = RxSerialPort.ReadByte();
                      hexVal += Convert.ToChar(inbyte);
                      ContReading.AppendText(Convert.ToChar(inbyte).ToString());
                  }

                  RxInUse = false;

                  int reading = Convert.ToInt16(hexVal, 16);
                  richTextBox1.AppendText(ScopeAz.ToString("F") + " " + ScopeAlt.ToString("F") + " " + reading.ToString() + Environment.NewLine);

                  var objWriter = new System.IO.StreamWriter("C:\\Program Files\\Apache Software Foundation\\Apache2.2\\htdocs\\forrest\\scan51.txt", true);
                  objWriter.WriteLine(ScopeAz.ToString("F") + " " + ScopeAlt.ToString("F") + " " + reading.ToString());
                  objWriter.Close();

              }
          }


      }

      PositionRead.Enabled = true;

  }
       */
        private void goto_Click(object sender, EventArgs e) {

            //TargetHor.az = Convert.ToDouble(AzUD.Value).ToRad();
            //TargetHor.alt = Convert.ToDouble(AltUD.Value).ToRad();

            /* 
             double AzR1 = Globals.TargetAz);
             double AzR2 = ScopeAz);
             double ElR1 = Globals.TargetAlt);
             double ElR2 = ScopeAlt);
             ScopetoTarget = rad2deg(Math.Acos(Math.Sin(ElR1) * Math.Sin(ElR2) + Math.Cos(ElR1) * Math.Cos(ElR2) * Math.Cos(AzR1 - AzR2)));
             Distlbl.Text = "Distance: " + ScopetoTarget.ToString("F");
             */

            PositionRead.Enabled = true;

            SlewScope();
        }

        private void ParkBtn_Click(object sender, EventArgs e) {
            //Park the telescope 
            sourceRB.Checked = false;
            AltAzRB.Checked = false;
            eqRB.Checked = false;
            GalRB.Checked = false;
            TargetHor.alt = 82d.ToRad();
            TargetHor.az = 180d.ToRad();
            SlewScope();
        }

        private void track_Click(object sender, EventArgs e) {
            //switch tracking state

            if(tracking) {
                trackBtn.Text = "Track";
                tracking = false;
            } else {
                trackBtn.Text = "Stop tracking";
                tracking = true;

            }

        }

        void stopScope() {
            while(ScopeInUse) { }
            ScopeInUse = true;
            ScopeSerialPort.Write("C1 2\r");
            ScopeInUse = false;
        }

        private void clearErrors_Click(object sender, EventArgs e) {
            //try and clear any telescope errors

            while(ScopeInUse) { }
            ScopeInUse = true;
            ScopeSerialPort.DiscardInBuffer();
            ScopeSerialPort.Write("C3\r");
            ScopeInUse = false;

        }
        //------------------------------------------------ 
        //Skymap stuff

        uint hdc;
        double thetaC = Math.PI;
        double phiC = 90d.ToRad();

        int list;
        double JD = 2451545;
        double lst;

        double Scopey, Scopex, Scopez;

        public static double centerx, centery, centerz, fov, tanfov;
        public static bool ortho = false;

        public void Draw() {
            //Updates the skymap

            //Alpha smoothing and blending and whatnot------------------------------------------------
            Gl.glEnable(Gl.GL_MULTISAMPLE);
            Gl.glEnable(Gl.GL_POINT_SMOOTH);
            Gl.glEnable(Gl.GL_POINT_SMOOTH_HINT);
            Gl.glEnable(Gl.GL_BLEND);
            Gl.glBlendFunc(Gl.GL_SRC_ALPHA, Gl.GL_ONE_MINUS_SRC_ALPHA);


            //Make sphere for sky map texture------------------------------------------------
            Gl.glClearColor(1, 1, 1, 1);
            Gl.glColor3f(1.0f, 1.0f, 1.0f);

            Gl.glEnable(Gl.GL_TEXTURE_2D);
            if(rb_2cm.Checked) {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("21cm-2.jpg"));
            } else if(rb_408.Checked) {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("408-2.jpg"));
            } else if(rb_halpha.Checked) {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("halpha.jpg"));
            } else if(rb_optical.Checked) {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("optical.jpg"));
            } else if(rb_ir.Checked) {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("iras.jpg"));
            } else if(rb_wmap.Checked) {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("wmap.jpg"));
            } else {
                Gl.glBindTexture(Gl.GL_TEXTURE_2D, ContentManager.GetTextureByName("tycho.jpg"));
            }
            Gl.glPushMatrix();

            //Rotate 90 degrees
            Gl.glRotated(90, -1, 0, 0);
            //Rotate to equatorial axis
            Gl.glRotated(90 - location.lat.ToDeg(), 0, 0, 1);
            //rotate depending on LST
            Gl.glRotated(lst.ToDeg() + 90, 0, 1, 0);

            Gl.glCallList(list);
            Gl.glPopMatrix();
            Gl.glDisable(Gl.GL_TEXTURE_2D);

            //Sun------------------------------------------------

            if(SunHor.alt > 0.035) { //~2 degs

                Gl.glPointSize(15);
                Gl.glColor4d(1, 0.7, 0, 0.99);

                Gl.glBegin(Gl.GL_POINTS);

                Gl.glVertex3d(Math.Cos(SunHor.alt) * Math.Cos(SunHor.az), -Math.Cos(SunHor.alt) * Math.Sin(SunHor.az), 0);

                Gl.glEnd();
            }

            //Moon------------------------------------------------

            if(MoonHor.alt > 0.035) { //~2 degs

                Gl.glPointSize(15);
                Gl.glColor4d(0.6, 0.6, 0.6, 0.99);

                Gl.glBegin(Gl.GL_POINTS);

                Gl.glVertex3d(Math.Cos(MoonHor.alt) * Math.Cos(MoonHor.az), -Math.Cos(MoonHor.alt) * Math.Sin(MoonHor.az), 0);

                Gl.glEnd();

            }

            //Alt az grid------------------------------------------------

            Gl.glColor4d(0.8, 0.8, 0.8, 0.6);

            for(int i = 0; i < 360; i = i + 30) {
                Gl.glBegin(Gl.GL_LINE_STRIP);
                for(int j = 0; j < 91; j++) {
                    Gl.glVertex3d(
                        1.001 * Math.Cos((j * 1.0).ToRad()) * Math.Cos((i * 1.0).ToRad()),
                        1.001 * Math.Cos((j * 1.0).ToRad()) * Math.Sin((i * 1.0).ToRad()),
                        1.001 * Math.Sin((j * 1.0).ToRad()));
                }
                Gl.glEnd();
            }

            for(int j = 10; j < 91; j = j + 10) {
                Gl.glBegin(Gl.GL_LINE_LOOP);
                for(int i = 0; i < 361; i++) {
                    Gl.glVertex3d(
                        1.001 * Math.Cos((j * 1.0).ToRad()) * Math.Cos((i * 1.0).ToRad()),
                        1.001 * Math.Cos((j * 1.0).ToRad()) * Math.Sin((i * 1.0).ToRad()),
                        1.001 * Math.Sin((j * 1.0).ToRad()));
                }
                Gl.glEnd();
            }

            //zenith blocking disc------------------------------------------------

            double x1, y1, z1, x2, y2, z2;
            double angle;
            double radius = 0.21;

            x1 = 0;
            y1 = 0;
            z1 = 0.999;

            Gl.glColor4d(0, 0, 0, 0.3);
            Gl.glBegin(Gl.GL_TRIANGLE_FAN);

            Gl.glVertex3d(x1, y1, z1);

            for(angle = 0; angle < 361; angle += 0.5) {
                x2 = x1 + Math.Sin(angle.ToRad()) * radius;
                y2 = y1 + Math.Cos(angle.ToRad()) * radius;

                z2 = z1;
                Gl.glVertex3d(x2, y2, z2);
            }
            Gl.glEnd();

            //Target circle------------------------------------------------

            if(TargetHor.alt > 0) {

                Gl.glColor4d(0.2, 0.6, 1.0, 0.9);

                x1 = Math.Cos(TargetHor.alt) * Math.Cos(TargetHor.az);
                y1 = -Math.Cos(TargetHor.alt) * Math.Sin(TargetHor.az);

                z1 = -0.04;
                radius = 0.03;

                Gl.glLineWidth(2.0f);

                Gl.glBegin(Gl.GL_LINE_LOOP);

                for(angle = 0; angle < 361; angle += 0.5) {
                    x2 = x1 + Math.Sin(angle.ToRad()) * radius;
                    y2 = y1 + Math.Cos(angle.ToRad()) * radius;

                    z2 = z1;
                    Gl.glVertex3d(x2, y2, z2);
                }

                Gl.glEnd();

                //Scope-to-target line------------------------------------------------

                Gl.glLineWidth(1.0f);

                Gl.glColor4d(0.1, 0.1, 1.0, 0.9);
                Gl.glBegin(Gl.GL_LINE_STRIP);

                Gl.glVertex3d(Scopex, Scopey, Scopez);
                Gl.glVertex3d(x1, y1, z1);

                Gl.glEnd();

            }

            //Scope------------------------------------------------

            Gl.glPointSize(20);
            if(ScopetoTarget < 0.06) {
                Gl.glColor4d(0.1, 1.0, 0.1, 0.9);
            } else {
                Gl.glColor4d(0.0, 0.6, 0.0, 0.9);
            }

            Gl.glBegin(Gl.GL_POINTS);

            if(followScope.Checked) {
                Scopex = Math.Cos(ScopeHor.alt) * Math.Cos(ScopeHor.az);
                Scopey = -Math.Cos(ScopeHor.alt) * Math.Sin(ScopeHor.az);
                Scopez = Math.Sin(ScopeHor.alt);

            } else {
                Scopex = Math.Cos(ScopeHor.alt) * Math.Cos(ScopeHor.az);
                Scopey = -Math.Cos(ScopeHor.alt) * Math.Sin(ScopeHor.az);
                Scopez = -0.05;

            }

            Gl.glVertex3d(Scopex, Scopey, Scopez);

            Gl.glEnd();

        }

        public void SkyMapUpdate() {
            //Changes sky map projection parameters

            Gl.glViewport(0, 0, pnlViewPort.Width, pnlViewPort.Height);
            Gl.glMatrixMode(Gl.GL_PROJECTION);
            Gl.glLoadIdentity();

            fov = 90;
            /*
            if (checkBox4.Checked) {
                fov = trackBar3.Value;
                if (fov > ScopeHor.alt.ToDeg()) {
                    //fov = rad2deg(ScopeAlt);
                }
            } else {
                fov = 90;
            }
              */

            //label1.Text = fov.ToString() + " " + ScopeAlt.ToString();

            tanfov = Math.Tan((fov + 1) * Math.PI / 360);
            if(followScope.Checked) {
                Glu.gluPerspective(fov, 1.0, 0.0, 5);
            } else {
                Gl.glOrtho(-tanfov, tanfov, -tanfov, tanfov, -0.05, 5);
            }
            Gl.glMatrixMode(Gl.GL_MODELVIEW);
            Gl.glLoadIdentity();

            Glu.gluLookAt(0, 0, 0, centerx, centery, centerz, 0, 0, 1);

        }

        private void SkyMapTimer_Tick(object sender, System.EventArgs e) {

            Gl.glClear(Gl.GL_COLOR_BUFFER_BIT | Gl.GL_DEPTH_BUFFER_BIT);

            if(followScope.Checked) {
                centerx = Scopex;
                centery = Scopey;
                centerz = Scopez;
            } else {
                centerx = Math.Cos(phiC) * Math.Cos(thetaC);
                centery = Math.Cos(phiC) * Math.Sin(thetaC);
                centerz = Math.Sin(phiC);
            }

            SkyMapUpdate();

            //Redraws
            Draw();

            SwapBuffers(hdc);

            Gl.glFlush();
        }

        private void InfoTimer_Tick(object sender, EventArgs e) {
            //Julian date
            JD = Utils.jd();

            double T = Utils.JDtoT(JD);

            double t = Utils.JDtot(JD);

            NutObl nutobl = new NutObl(T);

            //Sidereal time at Greenwich
            double thetaa = 280.46061837 + (360.98564736629 * (JD - 2451545)) + (0.000387933 * T * T) - ((T * T * T) / 38710000);

            //Local sidereal time
            lst = Utils.lst(JD, T, location);

            double lstdeg = Utils.quad(lst.ToDeg());

            lstLbl.Text = "LST: " + Utils.hh(lstdeg / 15) + "h " + Utils.dm(lstdeg / 15) + "m " + Utils.ds(lstdeg / 15) + "s";
            jdLbl.Text = "JD: " + JD.ToString("F5");

            DateTime time = DateTime.UtcNow;

            timeLbl.Text = "Time: " + time.ToString("HH") + "h " + time.ToString("mm") + "m " + time.ToString("ss") + "s UTC";
            dateLbl.Text = "Date: " + time.ToString("dd MMMM yyyy");

            //Sun------------------------------------------------ 

            EclipticalCoords Earth = SolarSystem.Planet("Earth", t);
            EclipticalCoords Sun = SolarSystem.Sun(Earth);

            SunEq = Sun.EclipticalToEquatorial(nutobl).J2000ToNow(JD);
            CoordsString SunEqS = SunEq.ToDMS();
            SunHor = SunEq.ToHorizontal(location, T);

            sunEqLbl.Text = "RA:" + SunEqS.ra + " Dec: " + SunEqS.dec;

            sunHorLbl.Text = "Az: " + SunHor.az.ToDeg().ToString("F3") + " Alt: " + SunHor.alt.ToDeg().ToString("F3");

            //Moon------------------------------------------------

            EclipticalCoords Moon = SolarSystem.MoonLow(T);
            Moon.l += nutobl.deltapsi;

            MoonEq = Moon.EclipticalToEquatorial(nutobl);

            //Parallax correction

            double deltaau = Moon.r / 149597870.691;

            double H = 70;

            double sinpi = (Math.Sin((8.794 / 3600).ToRad())) / deltaau;

            double ruserlat = location.lat;
            //double ruserlong = -location.lon;

            double C = 1 - (1 / 298.257);
            double u = Math.Atan(C * Math.Tan(ruserlat));
            double psin = C * Math.Sin(u) + (H / 6378140) * Math.Sin(ruserlat);
            double pcos = Math.Cos(u) + (H / 6378140) * Math.Cos(ruserlat);

            H = lst - MoonEq.ra;
            //hour angle

            //double moonH = H;
            double mA = (Math.Cos(MoonEq.dec)) * (Math.Sin(H));
            double mB = (Math.Cos(MoonEq.dec) * Math.Cos(H)) - (pcos * sinpi);
            double mC = (Math.Sin(MoonEq.dec)) - (psin * sinpi);
            double Q = Math.Sqrt((mA * mA) + (mB * mB) + (mC * mC));

            MoonEq.ra = lst - Math.Atan2(mA, mB);

            MoonEq.dec = Math.Asin(mC / Q);

            CoordsString MoonEqS = MoonEq.ToDMS();

            moonEqLbl.Text = "RA:" + MoonEqS.ra + " Dec: " + MoonEqS.dec;

            MoonHor = MoonEq.ToHorizontal(location, T);
            moonHorLbl.Text = "Az: " + MoonHor.az.ToDeg().ToString("F3") + " Alt: " + MoonHor.alt.ToDeg().ToString("F3");

        }

        //-------------------------------------------------------------------------------------
        //simple formy bits

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e) {
            AboutBox1 aboutbox = new AboutBox1();
            aboutbox.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
            RxSerialPort.Close();
            ScopeSerialPort.Close();

            RxSerialPort.Dispose();
            ScopeSerialPort.Dispose();
        }

        private void Gain70_MouseUp_1(object sender, MouseEventArgs e) {
            Gain70Update();
        }

        private void sourceBox_SelectedIndexChanged(object sender, EventArgs e) {
            tracking = false;
            trackBtn.Text = "Track";
        }

        private void ContStartStop_Click(object sender, EventArgs e) {

            if(ContTimer.Enabled) {

                ContTimer.Enabled = false;
                ContStartStop.Text = "Start";

            } else {

                saveName = "";
                SaveData.FileName = "l" + ScopeGal.l.ToDeg().ToString("F0") + "b" + ScopeGal.b.ToDeg().ToString("F0") + ".csv";
                SaveData.ShowDialog();

                if(saveName != "") {
                    File.AppendAllText(saveName, DateTime.Now.ToString() + Environment.NewLine);
                }

                ContTimer.Enabled = true;
                ContStartStop.Text = "Stop";
            }
        }

        private void sChanStartStop_Click(object sender, EventArgs e) {

            if(sChanTimer.Enabled) {

                sChanTimer.Enabled = false;
                sChanStartStop.Text = "Start";

            } else {

                saveName = "";
                SaveData.FileName = "l" + ScopeGal.l.ToDeg().ToString("F0") + "b" + ScopeGal.b.ToDeg().ToString("F0") + ".csv";
                SaveData.ShowDialog();

                if(saveName != "") {
                    File.AppendAllText(saveName, DateTime.Now.ToString() + Environment.NewLine);
                }

                sChanTimer.Enabled = true;
                sChanStartStop.Text = "Stop";
            }
        }

        private void button4_Click_1(object sender, EventArgs e) {
            PositionRead.Enabled = false;
        }

        private void pnlViewPort_MouseUp(object sender, MouseEventArgs e) {

            double x = ((double)e.X - (pnlViewPort.Width / 2)) / (pnlViewPort.Width / 2);
            double y = ((double)e.Y - (pnlViewPort.Height / 2)) / (pnlViewPort.Height / 2);

            double r = Math.Sqrt(x * x + y * y);

            double theta = Math.Atan2(x, y).ToDeg() + 180;

            double phi;
            if(r < 0.985) {
                phi = Math.Acos(r / 0.985).ToDeg();
            } else {
                phi = 10;
            }

            DialogResult dialogResult = MessageBox.Show("Do you want to set target coordinates to\r\nAzimuth: " + Math.Round(theta, 2).ToString() + (char)176 + " Altitude: " + Math.Round(phi, 2).ToString() + (char)176 + " ?", "Set target from map click", MessageBoxButtons.YesNo);

            if(dialogResult == DialogResult.Yes) {

                AltAzRB.Checked = true;
                TargetHor.az = theta;
                TargetHor.alt = phi;
                AzUD.Value = Convert.ToDecimal(theta);
                AltUD.Value = Convert.ToDecimal(phi);
            }
        }

        private void SaveData_FileOk(object sender, CancelEventArgs e) {
            saveName = SaveData.FileName;
        }

        private void ContOffset_MouseUp(object sender, MouseEventArgs e) {
            ContOffsetUpdate();
        }

        private void ContGain_SelectedIndexChanged(object sender, EventArgs e) {
            ContGainUpdate();
        }

        private void ContInt_SelectedIndexChanged(object sender, EventArgs e) {
            ContIntUpdate();
        }

        private void SpecOffset_MouseUp(object sender, MouseEventArgs e) {
            SpecOffsetUpdate();
        }

        private void SpecGain_SelectedIndexChanged(object sender, EventArgs e) {
            SpecGainUpdate();
        }

        private void SpecInt_SelectedIndexChanged(object sender, EventArgs e) {
            SpecIntUpdate();
        }

        private void NumSteps_ValueChanged(object sender, EventArgs e) {
            SpecAxisUpdate();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e) {
            ContTimer.Interval = Convert.ToInt32(ContGraphInt.Value * 1000);
        }

        private void ResetGraph_Click(object sender, EventArgs e) {
            //reset continuum graph
            for(k = 0; k < 600; k++) {
                ContArray[k] = 4097;
            }
        }

        private void start_ValueChanged(object sender, EventArgs e) {

            if(start.Value >= stop.Value) {
                start.Value = stop.Value - 1;
            }

            SpecAxisUpdate();
        }

        private void stop_ValueChanged(object sender, EventArgs e) {

            if(stop.Value <= start.Value) {
                stop.Value = start.Value + 1;
            }

            SpecAxisUpdate();
        }

        private void Spectrum_MouseDoubleClick(object sender, MouseEventArgs e) {
            SaveChart.ShowDialog();
        }

        private void SpectrumSaveChart_FileOk(object sender, CancelEventArgs e) {
            string name = SaveChart.FileName;
            Spectrum.SaveImage(name, ChartImageFormat.Png);
        }

        private void StopScope_Click(object sender, EventArgs e) {
            stopScope();
        }

        private void stopScope_Click(object sender, EventArgs e) {
            stopScope();
        }

        private void chanUpDown_ValueChanged(object sender, EventArgs e) {
            SpecChannelUpdate();
        }
    }
}