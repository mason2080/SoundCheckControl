using System;
using System.Windows.Forms;
using System.Net;
using System.Linq;
using System.Net.Sockets;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Data;
using Driver.File.Ini;
using Driver.Port.CSerialPort;
using System.Threading;
using System.Drawing;
using System.Data.OleDb;
using ThreadState = System.Threading.ThreadState;
using System.Runtime.InteropServices;

namespace SoundCheckTCPClient
{
   
    enum BoxStatus
    {
        UNKNOW,
        CLOSE,
        OPEN
    }



    public partial class Form1 : Form
    {
        [DllImport("SetRange.dll", EntryPoint = "SetRange")]
        static extern void SetRange(double index, double[] ValueX, double[] ValueY, Int32 LenX, Int32 LenY);

        private static System.Diagnostics.Process p;

        private TcpClient client;
        public StreamReader STR;
        public StreamWriter STW;
        dynamic json; // Dynamic object for working with JSON response from SoundCheck
        private JArray stepsList; // JSON array to store steps list from Sequence.GetStepsList

        CSerialPort com1 = new CSerialPort();
        CSerialPort com2 = new CSerialPort();
        CSerialPort com3= new CSerialPort();
        CSerialPort com4 = new CSerialPort();
        CSerialPort com5 = new CSerialPort();
        CSerialPort com6 = new CSerialPort();

        IniFileClass iniFile = new IniFileClass();
        Thread RecvCom1Thread;
        Thread RecvCom2Thread;
        Thread RecvCom3Thread;
        Thread RecvCom4Thread;
        Thread RecvCom5Thread;
        Thread RecvCom6Thread;


        Thread LeftTestThread;
        Thread RightTestThread;
        Thread RunSeqThread;
        byte timeCount1;

        BoxStatus box1Status = BoxStatus.UNKNOW;
        BoxStatus box2Status = BoxStatus.UNKNOW;
        bool flagCom4OK;
        bool flagCom4OK_HFP;
        bool flag_TOHFP;

        bool flagPassrowd = false;
        string seqPath = "";
        string seqPathMainTest = "";
        string seqPathNullTest = "";
        string stepPathLeft = "";
        string stepPathRight = "";
        string rangePath = "";
        string curvePath = "";

        public delegate void CrossDelegateShowCurve();

        public  string projectName;

        bool flagUpadateCurveCombo = false;

        bool[] busyFlag = new bool[2];
        bool testings = false;

        int[] curveIndex = new int[6];
        int[] curveJudgeType = new int[12];
        double[] curveLowLimit = new double[6];
        double[] curveHighLimit = new double[6];

        int failCntLeft = 0;
        int passCntLeft = 0;

        int failCntRight = 0;
        int passCntRight = 0;
        string whichcontrol_name = "";
        int curveRightClickIndex = 0;

        double[] XData;
        double[] YData;
        double[] XData_Min;
        double[] YData_Min;
        double[] XData_Max;
        double[] YData_Max;


        double[] XData1;
        double[] YData1;
        double[] XData1_Min;
        double[] YData1_Min;
        double[] XData1_Max;
        double[] YData1_Max;

        double[] XData2;
        double[] YData2;
        double[] XData2_Min;
        double[] YData2_Min;
        double[] XData2_Max;
        double[] YData2_Max;

        double[] XData3;
        double[] YData3;
        double[] XData3_Min;
        double[] YData3_Min;
        double[] XData3_Max;
        double[] YData3_Max;

        double[] XData4;
        double[] YData4;
        double[] XData4_Min;
        double[] YData4_Min;
        double[] XData4_Max;
        double[] YData4_Max;

        double[] XData5;
        double[] YData5;
        double[] XData5_Min;
        double[] YData5_Min;
        double[] XData5_Max;
        double[] YData5_Max;
        //

        double[] RXData;
        double[] RYData;
        double[] RXData_Min;
        double[] RYData_Min;
        double[] RXData_Max;
        double[] RYData_Max;


        double[] RXData1;
        double[] RYData1;
        double[] RXData1_Min;
        double[] RYData1_Min;
        double[] RXData1_Max;
        double[] RYData1_Max;

        double[] RXData2;
        double[] RYData2;
        double[] RXData2_Min;
        double[] RYData2_Min;
        double[] RXData2_Max;
        double[] RYData2_Max;

        double[] RXData3;
        double[] RYData3;
        double[] RXData3_Min;
        double[] RYData3_Min;
        double[] RXData3_Max;
        double[] RYData3_Max;

        double[] RXData4;
        double[] RYData4;
        double[] RXData4_Min;
        double[] RYData4_Min;
        double[] RXData4_Max;
        double[] RYData4_Max;

        double[] RXData5;
        double[] RYData5;
        double[] RXData5_Min;
        double[] RYData5_Min;
        double[] RXData5_Max;
        double[] RYData5_Max;


        string[] stepNames ={"Wait",
                             "RunSeq",
                             "Stop",
                             "WatingToTest",
                             "SetBusy",
                             "SetReady",
                             "BOX1/COM1_SEND",
                             "BOX1/COM1_SEND_GET",
                             "BOX1/COM1_GET",
                             "BOX2/COM2_SEND",
                             "BOX2/COM2_SEND_GET",
                             "BOX2/COM2_GET",
                             "SWITCH/COM3_SEND",
                             "SWITCH/COM3_SEND_GET",
                             "BT1/COM4_SEND",
                             "BT1/COM4_SEND_GET",
                             "BT2/COM5_SEND",
                             "BT2/COM5_SEND_GET",
                             "VP/COM7_SEND",
                             "VP/COM7_GET",
                             "ShowDialog",
                             "LogMsg",
                             "LoopTest"
                             };
        string[] stepTackleError= { "重测", "继续", "停止" };
        string[] retryError = { "继续", "停止" };
        public Form1(string Name,string sqcpath)
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            ConnectToIPAddress.Text = "127.0.0.1"; // loopback address (also known as localhost)
            ConnectToPort.Text = "4444"; // Default port
            //Disable all controls other than server connection controls
            ConnectToIPAddress.Enabled = true;
            ConnectToPort.Enabled = true;
            btnConnectToServer.Enabled = true;
            //textBoxSeqFilePath.Enabled = false;
           // btnSeqSelect.Enabled = false;
            //btnSeqOpen.Enabled = false;
            btnSeqRun.Enabled = false;
            comboCurveNames.Enabled = false;
            textBoxLotNumber.Enabled = false;
            textBoxSerialNumber.Enabled = false;
            btnSetLotNumber.Enabled = false;
            btnSetSerialNumber.Enabled = false;
            btnExitSoundCheck.Enabled = false;

            labelMargin.Text = "";
            labelVerdict.Text = "";
            comboBox1.SelectedIndex = 0;
            //comboBoxWindowState.SelectedIndex = 1;
            textBoxLog.Text = DateTime.Now.ToString() + ": This application started.";

            projectName = Name;
            flagCom4OK = false;
            flagCom4OK_HFP = false;
            flag_TOHFP = false;
            string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";

            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            buttonStart.Enabled = false;
            btnStop.Enabled = false;
            timeCount1 = 0;
            textBoxSeqFilePath.Text = sqcpath;
            try
            {

                comboBox1.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "SerialPort1"));
                comboBox2.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "SerialPort2"));
                comboBox3.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "SerialPort3"));
                comboBox4.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "SerialPort4"));
                comboBox5.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "SerialPort5"));
                comboBoxVirtual.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "VirtualPort"));
                comboBoxWindowState.SelectedIndex = int.Parse(iniFile.IniReadValue(fName, "基准", "SouncCheckStatus"));
                textBoxExePath.Text= iniFile.IniReadValue(fName, "基准", "SouncCheckPath");
                //textBoxSeqFilePath.Text = iniFile.IniReadValue(fName, "基准", "SeqPath");
                ConnectToIPAddress.Text = iniFile.IniReadValue(fName, "基准", "IP");
                ConnectToPort.Text = iniFile.IniReadValue(fName, "基准", "IP_Port");

                textSCConfigFileName.Text = iniFile.IniReadValue(fName, "基准", "SoundCheckConfigFileName");

                curveIndex[0] = int.Parse(iniFile.IniReadValue(fName, "基准", "comboCurve1"));
                curveIndex[1] = int.Parse(iniFile.IniReadValue(fName, "基准", "comboCurve2"));
                curveIndex[2] = int.Parse(iniFile.IniReadValue(fName, "基准", "comboCurve3"));
                curveIndex[3] = int.Parse(iniFile.IniReadValue(fName, "基准", "comboCurve4"));
                curveIndex[4] = int.Parse(iniFile.IniReadValue(fName, "基准", "comboCurve5"));
                curveIndex[5] = int.Parse(iniFile.IniReadValue(fName, "基准", "comboCurve6"));

                //curveJudgeType[0] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge1"));
                //curveJudgeType[1] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge2"));
                //curveJudgeType[2] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge3"));
                //curveJudgeType[3] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge4"));
                //curveJudgeType[4] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge5"));
                //curveJudgeType[5] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge6"));

                //curveLowLimit[0] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveLowLimit1"));
                //curveLowLimit[1] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveLowLimit2"));
                //curveLowLimit[2] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveLowLimit3"));
                //curveLowLimit[3] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveLowLimit4"));
                //curveLowLimit[4] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveLowLimit5"));
                //curveLowLimit[5] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveLowLimit6"));

                //curveHighLimit[0] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveHighLimit1"));
                //curveHighLimit[1] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveHighLimit2"));
                //curveHighLimit[2] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveHighLimit3"));
                //curveHighLimit[3] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveHighLimit4"));
                //curveHighLimit[4] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveHighLimit5"));
                //curveHighLimit[5] = double.Parse(iniFile.IniReadValue(fName, "基准", "CurveHighLimit6"));
            }
            catch
            { }

            RunSoundCheck();
            timerConnectSC.Enabled = true;
            //InitMenu();
        }

        ////Add Procuct To Menu
        //private void InitMenu()
        //{
        //    ToolStripMenuItem subItem;
        //    subItem = 项目选择ToolStripMenuItem;

        //    string fName = Directory.GetCurrentDirectory().ToString() + "\\测试产品";
        //    string[] subfolders = Directory.GetDirectories(fName);
        //    foreach (string s in subfolders)
        //    {
        //        ToolStripMenuItem tsmi2 = new ToolStripMenuItem(s.Substring(s.LastIndexOf(@"\") + 1));
        //        tsmi2.Click += MenuClick;

        //        subItem.DropDownItems.Add(tsmi2);
        //    }

        //}

        //自己定义个点击事件需要执行的动作
        private void MenuClick(object sender, EventArgs e)
        {
            ToolStripMenuItem but = sender as ToolStripMenuItem;
            textProduct.Text = but.Text;


            seqPathMainTest = "";
            seqPathNullTest = "";

            try
            {
                string fPath = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + but.Text;

                stepPathLeft = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + but.Text + "\\StepLeft.csv";
                stepPathRight = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + but.Text + "\\StepRight.csv";

                rangePath = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + but.Text + "\\Range.csv";

                if (File.Exists(stepPathLeft))
                {
                    ReadCsvToDataGrid(stepPathLeft, 1, dataGridViewLeft);

                    AppendLogMessage("Load Steps:" + stepPathLeft);
                }
                else
                {
                    MessageBox.Show("工步配置文件不存在");
                    textTestResultL.Text = "加载失败";
                    textTestResultL.BackColor = Color.Red;
                    return;
                }

                if (File.Exists(stepPathRight))
                {
                    ReadCsvToDataGrid(stepPathRight, 1, dataGridViewRight);

                    AppendLogMessage("Load Steps:" + stepPathRight);
                }
                else
                {
                    MessageBox.Show("工步配置文件不存在");
                    textTestResultL.Text = "加载失败";
                    textTestResultL.BackColor = Color.Red;
                    return;
                }

                if (File.Exists(rangePath))
                {
                    ReadCsvToValueTable(rangePath, 1, dataGridView);

                    AppendLogMessage("Load Range:" + rangePath);
                }
                else
                {
                    MessageBox.Show("判断标准文件不存在");
                    textTestResultL.Text = "加载失败";
                    textTestResultL.BackColor = Color.Red;
                    return;
                }

                DirectoryInfo folder = new DirectoryInfo(fPath);
                foreach (FileInfo file in folder.GetFiles("*.sqc"))
                {
                    if (file.FullName.IndexOf("(Main).sqc") >= 0)
                    {
                        seqPathMainTest = file.FullName;
                        AppendLogMessage("Load Main sqc:" + file.FullName);
                    }

                    //if (file.FullName.IndexOf("(Null).sqc") >= 0)
                    //{
                    //    seqPathNullTest = file.FullName;
                    //    AppendLogMessage("Load Null sqc:" + file.FullName);
                    //}
                }

                if ((seqPathMainTest != ""))// && (seqPathNullTest != ""))
                {
                    buttonStart.Visible = true;
                    btnStop.Visible = true;
                    tabControl2.Enabled = true;
                    tabControl2.SelectedIndex = 0;

                    textTestResultL.Text = "加载成功";
                   
                }
                else
                {
                    textTestResultL.Text = "加载失败";
                    textTestResultL.BackColor = Color.Red;
                }

            }
            catch
            {
                textProduct.Text = but.Text+"配置文件异常";
            }

            
        }


       
        private void LoadProjectConfig(string projectName)
        {
    

            seqPathMainTest = "";
            seqPathNullTest = "";

            try
            {
                string fPath = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" +projectName;
                stepPathLeft = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + projectName + "\\StepLeft.csv";
                stepPathRight = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + projectName + "\\StepRight.csv";

                rangePath = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + projectName + "\\Range.csv";
                curvePath = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + projectName + "\\curve.ini";
                if (File.Exists(stepPathLeft))
                {
                    ReadCsvToDataGrid(stepPathLeft, 1, dataGridViewLeft);

                    AppendLogMessage("Load Steps:" + stepPathLeft);
                }
                else
                {
                    MessageBox.Show("工步配置文件不存在");
                    textTestResultL.Text = "加载失败";
                    textTestResultL.BackColor = Color.Red;
                    return;
                }

                if (File.Exists(stepPathRight))
                {
                    ReadCsvToDataGrid(stepPathRight, 1, dataGridViewRight);

                    AppendLogMessage("Load Steps:" + stepPathRight);
                }
                else
                {
                    MessageBox.Show("工步配置文件不存在");
                    textTestResultL.Text = "加载失败";
                    textTestResultL.BackColor = Color.Red;
                    return;
                }
                    buttonStart.Visible = true;
                    btnStop.Visible = true;
                    tabControl2.Enabled = true;
                    tabControl2.SelectedIndex = 0;
            }
            catch
            {
                textProduct.Text = projectName + "配置文件异常";
            }


        }


        private string ReadLineFromStream() // Read a CRLF terminated string from TCP connection
        {
            try
            {
                string receive = null; // Clear receive string

                while (receive == null) // Keep trying to read from stream untill we read a valid line
                {
                    receive = STR.ReadLine(); //Read line from stream

                    if (receive == null)
                    { System.Threading.Thread.Sleep(100); } // No line read yet, wait for some time and then check again
                }
                return receive; // Line read, pass it out
            }
            catch (Exception x)
            {
                if (client.Connected)
                { MessageBox.Show(x.Message.ToString()); } //Client is connected, not sure what error this is. Show mesagebox.

                return null;
            }
        }
              
        private bool GetResponseJSON() // Get JSON response string from SoundCheck and convert it into convert it into dynamic json object
        {
            string response = ReadLineFromStream();

            if (response != null)
            {
                json = JValue.Parse(response); // Parse the received JSON string as a Dynamic JObject
                return true; // Return success, which means the reaponse was received from SoundCheck
            } 
            else
            {         
                return false; // Return failure, which means the reaponse was not received from SoundCheck
            }
        }

        private bool SendCommandAndGetResponse(string SCCommand) // Send command to SoundCheck and wait for response. Convert Response to dynamic object.
        {
            if (client.Connected)
            {
                STW.WriteLine(SCCommand + "\r\n"); //Send command to server, with CRLF termination

                if (GetResponseJSON()) // Wait for response from server
                { return GetCommandCompleted(); } // return the cmdCompleted boolean from command response as the result of this method call
                else
                {
                    if (client.Connected == false) // The command was sent, but did not complete because connection was lost while trying to read the response
                    {
                        SetConnectedState(false);
                        MessageBox.Show("You are not connected to SoundCheck." + Environment.NewLine + "Please connect and try again!");
                    }
                    return false;
                }
            }
            else
            {
                SetConnectedState(false);
                MessageBox.Show("You are not connected to SoundCheck." + Environment.NewLine + "Please connect and try again!");
                return false;
            }
        }

        private void SetConnectedState(Boolean connected) // Update the UI to reflect the connection status
        {
            if (connected)
            {
                // Disable server connection controls
                ConnectToIPAddress.Enabled = false;
                ConnectToPort.Enabled = false;
                btnConnectToServer.Enabled = false;

                // Enable sequence open controls
                textBoxSeqFilePath.Enabled = true;
               // btnSeqSelect.Enabled = true;
               // btnSeqOpen.Enabled = true;
                textBoxLotNumber.Enabled = true;
                textBoxSerialNumber.Enabled = true;
                btnSetLotNumber.Enabled = true;
                btnSetSerialNumber.Enabled = true;
                btnExitSoundCheck.Enabled = true;
                buttonStart.Visible = true;
                timerConnectSC.Enabled = false;

            }
            else
            {
                // Enable server connection controls
                ConnectToIPAddress.Enabled = true;
                ConnectToPort.Enabled = true;
                btnConnectToServer.Enabled = true;
            }
        }

        private void btnSelectExe_Click(object sender, EventArgs e) // Select SoundCheck exe
        {
            if (dialogExeSelect.ShowDialog() == DialogResult.OK)
            {
                textBoxExePath.Text = dialogExeSelect.FileName;
            }           
        }


        private void RunSoundCheck()
        {
            var exeFileName = Path.GetFileNameWithoutExtension(textBoxExePath.Text);
            Process[] SoundCheckProcesses = Process.GetProcessesByName(exeFileName);
            if (SoundCheckProcesses.Length == 0)
            {
                string option;
                switch ((string)comboBoxWindowState.SelectedItem) // Get commandline option from user selection
                {
                    case "Hidden":
                        option = "-h";
                        break;
                    case "Minimized":
                        option = "-m";
                        break;
                    default:
                        option = "";
                        break;
                }

                try
                {
                    Process.Start(textBoxExePath.Text, option); // Start SoundCheck with commandline arguments
                    AppendLogMessage(textBoxExePath.Text + " started.");
                }
                catch
                {
                   // Process.Start(textBoxExePath.Text, option); // Start SoundCheck with commandline arguments
                    AppendLogMessage( " SC can't be stared.");
                }

            }
            else
            {
                AppendLogMessage(Path.GetFileName(textBoxExePath.Text) + " is already running.");
            }
        }


        private void btnRunSoundCheck_Click(object sender, EventArgs e) // Run SoundCheck, if not already running
        {
            var exeFileName = Path.GetFileNameWithoutExtension(textBoxExePath.Text);
            Process[] SoundCheckProcesses = Process.GetProcessesByName(exeFileName);
            if (SoundCheckProcesses.Length == 0)
            {
                string option;
                switch ((string)comboBoxWindowState.SelectedItem) // Get commandline option from user selection
                {
                    case "Hidden":
                        option = "-h";
                        break;
                    case "Minimized":
                        option = "-m";
                        break;
                    default:
                        option = "";
                        break;
                }
                Process.Start(textBoxExePath.Text, option); // Start SoundCheck with commandline arguments
                AppendLogMessage(textBoxExePath.Text + " started.");
            }
            else
            {
                AppendLogMessage(Path.GetFileName(textBoxExePath.Text) + " is already running.");
            }
        }

        private void ConnectToServer() //Connect to server
        {
            // AppendLogMessage("Connecting to SoundCheck...");
            textSCStatus.Text = "尝试...";
            if (textSCStatus.BackColor == Color.Yellow)
            {
                textSCStatus.BackColor = Color.White;
            }
            else
            {
                textSCStatus.BackColor = Color.Yellow;
            }

            btnConnectToServer.Enabled = false; // Disable button so user doesn't click on it again

            client = new TcpClient();
            IPEndPoint IP_End = new IPEndPoint(IPAddress.Parse(ConnectToIPAddress.Text), int.Parse(ConnectToPort.Text));
            try
            {
                client.Connect(IP_End);
                if(client.Connected)
                {
                    STR = new StreamReader(client.GetStream());
                    STW = new StreamWriter(client.GetStream());
                    STW.AutoFlush = true;

                    ReadLineFromStream(); // SoundCheck will send an acknowledgement on connection. Read it.

                    // Send command to SoundCheck to set the strings to receive for 'NaN','Infinity','-Infinity' float values.
                    // C# uses the same strings as SoundCheck, so this step wouldn't be necessary if this were the 
                    // only application connecting to SoundCheck.
                    SendCommandAndGetResponse("SoundCheck.SetFloatStrings('NaN','Infinity','-Infinity')"); 

                    AppendLogMessage("Connected to SoundCheck.");
                    textSCStatus.Text = "连接成功";
                    textSCStatus.BackColor = Color.Green;
                    SetConnectedState(true);

                    btnOpenSeq2.Enabled = true;
                    btnOpenSeq2.BackColor = Color.Green;

                    //if (SendCommandAndGetResponse("Sequence.GetStepsList")) // Send Sequence.GetStepsList command and wait for result.
                    //{
                    //    // Command completed successfully
                    //    // Populate Data Table
                    //    stepsList = json.Value<JArray>("returnData"); // Convert return data to dynamic objects array
                    //}
                 }
            }
            catch(Exception)
            {
                btnConnectToServer.Enabled = true;
                //AppendLogMessage("Failed to connect to SoundCheck.");
        
                //MessageBox.Show("Could not connect to SoundCheck because the target machine refused it." + Environment.NewLine +
                //"Please make sure that TCP/IP server is enabled in SoundCheck Preferences dialog and try again.");
            }
        }
        private void btnConnectToServer_Click(object sender, EventArgs e) //Connect to server
        {
            AppendLogMessage("Connecting to SoundCheck...");
            btnConnectToServer.Enabled = false; // Disable button so user doesn't click on it again

            client = new TcpClient();
            IPEndPoint IP_End = new IPEndPoint(IPAddress.Parse(ConnectToIPAddress.Text), int.Parse(ConnectToPort.Text));

            try
            {
                client.Connect(IP_End);
                if (client.Connected)
                {
                    STR = new StreamReader(client.GetStream());
                    STW = new StreamWriter(client.GetStream());
                    STW.AutoFlush = true;

                    ReadLineFromStream(); // SoundCheck will send an acknowledgement on connection. Read it.

                    // Send command to SoundCheck to set the strings to receive for 'NaN','Infinity','-Infinity' float values.
                    // C# uses the same strings as SoundCheck, so this step wouldn't be necessary if this were the 
                    // only application connecting to SoundCheck.
                    SendCommandAndGetResponse("SoundCheck.SetFloatStrings('NaN','Infinity','-Infinity')");

                    AppendLogMessage("Connected to SoundCheck.");
                    textSCStatus.Text = "连接成功";
                    SetConnectedState(true);
                }
            }
            catch (Exception)
            {
                btnConnectToServer.Enabled = true;
                AppendLogMessage("Failed to connect to SoundCheck.");
                //MessageBox.Show("Could not connect to SoundCheck because the target machine refused it." + Environment.NewLine +
                //"Please make sure that TCP/IP server is enabled in SoundCheck Preferences dialog and try again.");
            }
        }

        private void btnSeqSelect_Click(object sender, EventArgs e) // Select Sequence: Show sequence file open dialog
        {
            if (dialogSeqSelect.ShowDialog() == DialogResult.OK)
            {
                textBoxSeqFilePath.Text = dialogSeqSelect.FileName;
            }
        }

        private void openSeqFile(string path)
        {
            if (path != "")
            {
                AppendLogMessage("Opening sequence ...");

                if (SendCommandAndGetResponse("Sequence.Open('" + path + "')")) // Send Sequence.Open command and wait for result.
                {
                    // Command completed successfully
                    if (GetReturnDataBoolean()) // Sequence.Open returns boolean data indicating if the sequence was opened.
                    {
                        btnSeqRun.Enabled = true;
                        AppendLogMessage("Sequence Opened. Ready to run!");
                    }
                    else
                    {
                        AppendLogMessage("Sequence failed to open.");
                    }
                }
                else
                {
                    // Command did not complete successfully
                    if (GetErrorType() == "Timeout") // Check if command timed out
                    {
                        AppendLogMessage("Sequence failed to open. Command timed out!");
                    }
                }
            }
        }

        private void btnSeqOpen_Click(object sender, EventArgs e) // Open Sequence: Send command to open sequence
        {
            if (textBoxSeqFilePath.Text != "")
            {
                AppendLogMessage("Opening sequence ...");

                if(SendCommandAndGetResponse("Sequence.Open('" + textBoxSeqFilePath.Text + "')")) // Send Sequence.Open command and wait for result.
                {
                    // Command completed successfully
                    if (GetReturnDataBoolean()) // Sequence.Open returns boolean data indicating if the sequence was opened.
                    {
                        btnSeqRun.Enabled = true;
                        AppendLogMessage("Sequence Opened. Ready to run!");
                        flagUpadateCurveCombo = false;
                        if (SendCommandAndGetResponse("Sequence.GetStepsList")) // Send Sequence.GetStepsList command and wait for result.
                        {
                            // Command completed successfully
                            
                            //Column Headers
                            string[] colHeaders = new string[] { "Step Name", "Step Type", "Input Channel", "Output Channel" };

                            DataTable dataTable = InitializeDataTable(colHeaders.Length);

                            // Populate Data Table
                            stepsList = json.Value<JArray>("returnData"); // Convert return data to dynamic objects array

                            dataGridView.Rows.Clear();

                            int i = 0;
                            foreach (JObject row in stepsList.Children<JObject>())
                            {
                                dataGridView.Rows.Add();

                                dataGridView.Rows[i].Cells[0].Value = false;
                                dataGridView.Rows[i].Cells[1].Value = row.Value<string>("Name");
                                dataGridView.Rows[i].Cells[2].Value = "";
                                dataGridView.Rows[i].Cells[3].Value = "";
                                dataGridView.Rows[i].Cells[4].Value = "";
                                i++;

                            }
                        }
                    }
                    else
                    {
                        AppendLogMessage("Sequence failed to open.");
                    }
                }
                else
                {
                    // Command did not complete successfully
                    if (GetErrorType() == "Timeout") // Check if command timed out
                    {
                        AppendLogMessage("Sequence failed to open. Command timed out!");
                    }
                }
            }
        }

        private DataTable InitializeDataTable(int numOfColumns) // Create new data table and set number of columns
        {
                // Create data table to hold data read from JSON
                DataTable dataTable = new DataTable();
                DataColumn dataColumn;

                for (int i = 0; i<numOfColumns; i++)
                {
                    dataColumn = new DataColumn();
                    dataTable.Columns.Add(dataColumn);
                }
            return dataTable;
            } 

        private void UpdateHeaders(string[] colHeaders) // Update row and column headers
        {
            // Update column Headers
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.HeaderCell.Value = colHeaders[column.Index];
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            // Update row Headers
            foreach (DataGridViewRow row in dataGridView.Rows)
            { row.HeaderCell.Value = (row.Index + 1).ToString(); }
        }

        private string FormatChannelNames(JArray channelNames)
        {    
            return string.Join(", ", channelNames.ToObject<string[]>()); // Convert JArray to string array and created comma separated string
        }


        private void GetData() // Run Null Sequence
        {
            int rowIndex = 0; ;

            //labelMargin.Text = ""; labelVerdict.Text = "";

            //openSeqFile(seqPathNullTest);

            AppendLogMessage("Running sequence ...");

            if (SendCommandAndGetResponse("Sequence.Run")) // Send Sequence.Run command and wait for result.
            {
                AppendLogMessage("Sequence ran to completion.");

                try
                {

                    // Update Data Table
                    // Column Headers`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
                    //string[] colHeaders = new string[] { "Step Name", "Step Type", "Input Channel", "Output Channel", "Verdict", "Margin", "Limit", "Max/Min" };
                    string[] colHeaders = new string[] { "Step Name", "Step Type", "Margin", "Limit", "Max/Min" };

                    DataTable dataTable = InitializeDataTable(colHeaders.Length);

                    // Populate Data Table
                    JArray stepResults = json.returnData.Value<JArray>("StepResults"); // Convert return data to dynamic objects array

                    dataGridViewRange.Rows.Clear();

                    for (int i = 0; i < stepResults.Count; i++)
                    // JObject row in stepResults.Children<JObject>()
                    {
                        DataRow dataRow = dataTable.NewRow();
                        dataGridViewRange.Rows.Add();

                        dataGridViewRange.Rows[i].Cells[0].Value = false;
                        dataGridViewRange.Rows[i].Cells[1].Value = stepsList[i].Value<string>("Name");
                        dataGridViewRange.Rows[i].Cells[1].Value = "";
                        dataGridViewRange.Rows[i].Cells[1].Value = "";
                        dataGridViewRange.Rows[i].Cells[1].Value = "";
                    }
                }
                catch
                {
                    AppendLogMessage("数据获取失败");
                }
            }
        }

        private void LeftSeqRun() // Run Sequence
        {
            AppendLogMessage("LeftChan:Running sequence ...");

            if (SendCommandAndGetResponse("Sequence.Run")) // Send Sequence.Run command and wait for result.
            {

                // Command completed successfully
                AppendLogMessage("Sequence ran to completion.");
                passCntLeft = 0;
                failCntLeft = 0;
                try
                {
                    // Populate Data Table
                    JArray stepResults = json.returnData.Value<JArray>("StepResults"); // Convert return data to dynamic objects array

                    for (int i = 0; i < stepResults.Count; i++)
                    // JObject row in stepResults.Children<JObject>()
                    {
                        for (int j = 0; j < dataGridView.RowCount; j++)
                        {
                            if ((string)(dataGridView.Rows[j].Cells[1].Value) == stepsList[i].Value<string>("Name"))
                            {
                                dataGridView.Rows[j].Cells[4].Value= stepResults[i].Value<Double>("Margin").ToString();

                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.White;

                                try
                                {
                                    if ((double.Parse((string)(dataGridView.Rows[j].Cells[4].Value)) >= double.Parse((string)(dataGridView.Rows[j].Cells[2].Value))) && (double.Parse((string)(dataGridView.Rows[j].Cells[4].Value)) <= double.Parse((string)(dataGridView.Rows[j].Cells[3].Value))))
                                    {
                                        dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Green;
                                         passCntLeft++;
                                    }
                                    else
                                    {
                                        dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Red;
                                        failCntLeft++;
                                    }
                                }
                                catch
                                {

                                }
                            }
                        }
                    }
                }
                catch
                {
                    AppendLogMessage("数据获取失败");
                }

                try
                {

                    if (flagUpadateCurveCombo == false)
                    {
                        flagUpadateCurveCombo = true;
                        if (SendCommandAndGetResponse("MemoryList.GetAllNames")) // Send MemoryList.GetAllNames command and wait for result.
                        {
                            // Command completed successfully
                            JArray jarray = json.returnData.Value<JArray>("Curves"); // Convert return data to dynamic objects array 
                            comboCurveNames.Items.Clear(); // Clear existing items
                            comboCurveNames.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboCurveNames.SelectedIndex = curveIndex[0];

                            comboBox6.Items.Clear(); // Clear existing items
                            comboBox6.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox6.SelectedIndex = curveIndex[1];


                            comboBox7.Items.Clear(); // Clear existing items
                            comboBox7.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox7.SelectedIndex = curveIndex[2];


                            comboBox8.Items.Clear(); // Clear existing items
                            comboBox8.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox8.SelectedIndex = curveIndex[3];

                            comboBox9.Items.Clear(); // Clear existing items
                            comboBox9.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox9.SelectedIndex = curveIndex[4];

                            comboBox10.Items.Clear(); // Clear existing items
                            comboBox10.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox10.SelectedIndex = curveIndex[5];


                            
                        }
                    }
                    CrossDelegateShowCurve dl = new CrossDelegateShowCurve(showCurveLeft);
                    this.BeginInvoke(dl);
                    comboCurveNames.Enabled = true; // Enable the combobox
                }
                catch
                {
                    AppendLogMessage("图表获取失败");
                }
            }
        }
        private void RightSeqRun() // Run Sequence
        {

            AppendLogMessage("RightChannel:Running sequence ...");

            if (SendCommandAndGetResponse("Sequence.Run")) // Send Sequence.Run command and wait for result.
            {

                // Command completed successfully
                AppendLogMessage("Sequence ran to completion.");
                try
                {
                    JArray stepResults = json.returnData.Value<JArray>("StepResults"); // Convert return data to dynamic objects array
                    failCntRight = 0;
                    passCntRight = 0;

                    for (int i = 0; i < stepResults.Count; i++)
                    // JObject row in stepResults.Children<JObject>()
                    {
                        for (int j = 0; j < dataGridView.RowCount; j++)
                        {
                            if ((string)(dataGridView.Rows[j].Cells[1].Value) == stepsList[i].Value<string>("Name"))
                            {
                                dataGridView.Rows[j].Cells[5].Value = stepResults[i].Value<Double>("Margin").ToString();

                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.White;

                                try
                                {
                                    if ((double.Parse((string)(dataGridView.Rows[j].Cells[5].Value)) >= double.Parse((string)(dataGridView.Rows[j].Cells[2].Value))) && (double.Parse((string)(dataGridView.Rows[j].Cells[5].Value)) <= double.Parse((string)(dataGridView.Rows[j].Cells[3].Value))))
                                    {
                                        dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Green;
                                        passCntRight++;
                                    }
                                    else
                                    {
                                        dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Red;
                                        failCntRight++;
                                    }
                                }
                                catch
                                {

                                }
                            }
                        }

                    }
                }
                catch
                {
                    AppendLogMessage("数据获取失败");
                }

                try
                {
                    if (flagUpadateCurveCombo == false)
                    {
                        flagUpadateCurveCombo = true;
                        if (SendCommandAndGetResponse("MemoryList.GetAllNames")) // Send MemoryList.GetAllNames command and wait for result.
                        {
                            // Command completed successfully
                            JArray jarray = json.returnData.Value<JArray>("Curves"); // Convert return data to dynamic objects array 
                            comboCurveNames.Items.Clear(); // Clear existing items
                            comboCurveNames.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboCurveNames.SelectedIndex = curveIndex[0];



                            comboBox6.Items.Clear(); // Clear existing items
                            comboBox6.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox6.SelectedIndex = curveIndex[1];


                            comboBox7.Items.Clear(); // Clear existing items
                            comboBox7.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox7.SelectedIndex = curveIndex[2];


                            comboBox8.Items.Clear(); // Clear existing items
                            comboBox8.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox8.SelectedIndex = curveIndex[3];

                            comboBox9.Items.Clear(); // Clear existing items
                            comboBox9.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox9.SelectedIndex = curveIndex[4];

                            comboBox10.Items.Clear(); // Clear existing items
                            comboBox10.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                            comboBox10.SelectedIndex = curveIndex[5];

                            Thread.Sleep(100);


                        }
                    }
                    CrossDelegateShowCurve dl = new CrossDelegateShowCurve(showCurveRight);
                    this.BeginInvoke(dl);

                   // showCurve();
                        comboCurveNames.Enabled = true; // Enable the combobox
                }
                catch
                {
                    AppendLogMessage("图表获取失败");
                }
            }
        }

        private void btnSeqRun_Click(object sender, EventArgs e) // Run Sequence
        {
            //labelMargin.Text = ""; labelVerdict.Text = "";

            try
            {
                AppendLogMessage("Running sequence ...");

                if (SendCommandAndGetResponse("Sequence.Run")) // Send Sequence.Run command and wait for result.
                {
                    // Command completed successfully
                    AppendLogMessage("Sequence ran to completion.");

                    // Populate Data Table
                    JArray stepResults = json.returnData.Value<JArray>("StepResults"); // Convert return data to dynamic objects array

                    for (int i = 0; i < stepResults.Count; i++)
                    // JObject row in stepResults.Children<JObject>()
                    {
                        for (int j = 0; j < dataGridView.RowCount; j++)
                        {
                            if ((string)(dataGridView.Rows[j].Cells[1].Value) == stepsList[i].Value<string>("Name"))
                            {
                                dataGridView.Rows[j].Cells[4].Value = stepResults[i].Value<Double>("Margin").ToString();

                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.White;

                                try
                                {
                                    if ((double.Parse((string)(dataGridView.Rows[j].Cells[4].Value)) >= double.Parse((string)(dataGridView.Rows[j].Cells[2].Value))) && (double.Parse((string)(dataGridView.Rows[j].Cells[4].Value)) <= double.Parse((string)(dataGridView.Rows[j].Cells[3].Value))))
                                    {
                                        dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Green;
                                    }
                                    else
                                    {
                                        dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Red;
                                    }
                                }
                                catch
                                {

                                }
                            }
                            break;
                        }
                    }

                    if (SendCommandAndGetResponse("MemoryList.GetAllNames")) // Send MemoryList.GetAllNames command and wait for result.
                    {
                        // Command completed successfully
                        JArray jarray = json.returnData.Value<JArray>("Curves"); // Convert return data to dynamic objects array 
                        comboCurveNames.Items.Clear(); // Clear existing items
                        comboCurveNames.Items.AddRange(jarray.ToObject<string[]>()); //Update comboCurveNames combobox with the name of all curves received
                        comboCurveNames.Enabled = true; // Enable the combobox
                    }
                }

            }
            catch
            { }
        }

        private void judgeCurveFuncLeft(byte index)
        {
            int tempFailCnt = 0;

            double[] value_x;
            double[] value_y;
            double[] min_x;
            double[] min_y;
            double[] max_x;
            double[] max_y;

            if (index == 0)
            {
                value_x = XData;
                value_y = YData;
                min_x = XData_Min;
                min_y = YData_Min;
                max_x = XData_Max;
                max_y = YData_Max;
            }
            else if (index == 1)
            {
                value_x = XData1;
                value_y = YData1;
                min_x = XData1_Min;
                min_y = YData1_Min;
                max_x = XData1_Max;
                max_y = YData1_Max;
            }
            else if (index == 2)
            {
                value_x = XData2;
                value_y = YData2;
                min_x = XData2_Min;
                min_y = YData2_Min;
                max_x = XData2_Max;
                max_y = YData2_Max;
            }
            else if (index == 3)
            {
                value_x = XData3;
                value_y = YData3;
                min_x = XData3_Min;
                min_y = YData3_Min;
                max_x = XData3_Max;
                max_y = YData3_Max;
            }
            else if (index == 4)
            {
                value_x = XData4;
                value_y = YData4;
                min_x = XData4_Min;
                min_y = YData4_Min;
                max_x = XData4_Max;
                max_y = YData4_Max;
            }
            else if (index == 5)
            {
                value_x = XData5;
                value_y = YData5;
                min_x = XData5_Min;
                min_y = YData5_Min;
                max_x = XData5_Max;
                max_y = YData5_Max;
            }
            else
            {
                value_x = XData;
                value_y = YData;
                min_x = XData_Min;
                min_y = YData_Min;
                max_x = XData_Max;
                max_y = YData_Max;
            }


            if (curveJudgeType[index] > 0)
            {
                if (curveJudgeType[index] == 1)
                {
                    for (int j = 0; j < dataGridView.RowCount; j++) //find CurveData Index
                    {
                        if ((string)(dataGridView.Rows[j].Cells[1].Value) =="WaveForm"+index.ToString())
                        {
                            //dataGridView.Rows[j].Cells[4].Value = YData.Min().ToString() + "," + YData.Max().ToString();

                            for (int k = 0; k < min_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (min_x[k] == value_x[kk])//逐点比较下限范围值每个点
                                    {
                                        if (value_y[kk] >= min_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntLeft++;
                                        }
                                    }
                                }
                            }

                            for (int k = 0; k < max_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (max_x[k] == value_x[kk])//逐点比较上限限范围值每个点
                                    {
                                        if (value_y[kk] < max_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntLeft++;
                                        }
                                    }
                                }
                            }

                            ////////////////
                            if (tempFailCnt == 0)
                            {
                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Red;
                            }

                        }

                    }
                }
                else if (curveJudgeType[index] == 2)
                {
                    for (int j = 0; j < dataGridView.RowCount; j++) //find CurveData Index
                    {
                        if ((string)(dataGridView.Rows[j].Cells[1].Value) == "WaveForm" + index.ToString())
                        {
                          //  dataGridView.Rows[j].Cells[4].Value = YData.Min().ToString() + "," + YData.Max().ToString();

                            for (int k = 0; k < max_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (max_x[k] == value_x[kk])//逐点比较上限限范围值每个点
                                    {
                                        if (value_y[kk] < max_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntLeft++;
                                        }
                                    }
                                }
                            }

                            ////////////////
                            if (tempFailCnt == 0)
                            {
                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Red;
                            }

                        }

                    }
                }
                else if (curveJudgeType[index] == 3)
                {
                    for (int j = 0; j < dataGridView.RowCount; j++) //find CurveData Index
                    {
                        if ((string)(dataGridView.Rows[j].Cells[1].Value) == "WaveForm" + index.ToString())
                        {
                            //dataGridView.Rows[j].Cells[4].Value = YData.Min().ToString() + "," + YData.Max().ToString();

                            for (int k = 0; k < min_x.Length; k++)
                            {
                                for (int kk = 0; kk < XData.Length; kk++)
                                {
                                    if (min_x[k] == value_x[kk])//逐点比较下限范围值每个点
                                    {
                                        if (value_y[kk] >= min_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntLeft++;
                                        }
                                    }
                                }
                            }
                            ////////////////
                            if (tempFailCnt == 0)
                            {
                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView.Rows[j].Cells[4].Style.BackColor = Color.Red;
                            }

                        }

                    }
                }
            }
        }
        private void judgeCurveFuncRight(byte index)
        {
            int tempFailCnt = 0;

            double[] value_x;
            double[] value_y;
            double[] min_x;
            double[] min_y;
            double[] max_x;
            double[] max_y;

            if (index == 6)
            {
                value_x = RXData;
                value_y = RYData;
                min_x = RXData_Min;
                min_y = RYData_Min;
                max_x = RXData_Max;
                max_y = RYData_Max;
            }
            else if (index == 7)
            {
                value_x = RXData1;
                value_y = RYData1;
                min_x = RXData1_Min;
                min_y = RYData1_Min;
                max_x = RXData1_Max;
                max_y = RYData1_Max;
            }
            else if (index == 8)
            {
                value_x = RXData2;
                value_y = RYData2;
                min_x = RXData2_Min;
                min_y = RYData2_Min;
                max_x = RXData2_Max;
                max_y = RYData2_Max;
            }
            else if (index == 9)
            {
                value_x = RXData3;
                value_y = RYData3;
                min_x = RXData3_Min;
                min_y = RYData3_Min;
                max_x = RXData3_Max;
                max_y = RYData3_Max;
            }
            else if (index == 10)
            {
                value_x = RXData4;
                value_y = RYData4;
                min_x = RXData4_Min;
                min_y = RYData4_Min;
                max_x = RXData4_Max;
                max_y = RYData4_Max;
            }
            else if (index == 11)
            {
                value_x = RXData5;
                value_y = RYData5;
                min_x = RXData5_Min;
                min_y = RYData5_Min;
                max_x = RXData5_Max;
                max_y = RYData5_Max;
            }
            else
            {
                value_x = RXData;
                value_y = RYData;
                min_x = RXData_Min;
                min_y = RYData_Min;
                max_x = RXData_Max;
                max_y = RYData_Max;
            }


            if (curveJudgeType[index] > 0)
            {
                if (curveJudgeType[index] == 1)
                {

                    for (int j = 0; j < dataGridView.RowCount; j++) //find CurveData Index
                    {
                        if ((string)(dataGridView.Rows[j].Cells[1].Value) == "WaveForm" + (index-5).ToString())
                        {
                           // dataGridView.Rows[j].Cells[5].Value = value_y.Min().ToString() + "," + value_y.Max().ToString();

                            for (int k = 0; k < min_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (min_x[k] == value_x[kk])//逐点比较下限范围值每个点
                                    {
                                        if (value_y[kk] >= min_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntRight++;
                                        }
                                    }
                                }
                            }

                            for (int k = 0; k < max_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (max_x[k] == value_x[kk])//逐点比较上限限范围值每个点
                                    {
                                        if (value_y[kk] < max_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntRight++;
                                        }
                                    }
                                }
                            }

                            ////////////////
                            if (tempFailCnt == 0)
                            {
                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Red;
                            }

                        }

                    }
                }
                else if (curveJudgeType[index] == 2)
                {
                    for (int j = 0; j < dataGridView.RowCount; j++) //find CurveData Index
                    {
                        if ((string)(dataGridView.Rows[j].Cells[1].Value) == "WaveForm" + (index - 5).ToString())
                        {
                          //  dataGridView.Rows[j].Cells[5].Value = value_y.Max().ToString();

                            for (int k = 0; k < max_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (max_x[k] == value_x[kk])//逐点比较上限限范围值每个点
                                    {
                                        if (value_y[kk] < max_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntRight++;
                                        }
                                    }
                                }
                            }

                            ////////////////
                            if (tempFailCnt == 0)
                            {
                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Red;
                            }

                        }

                    }
                }
                else if (curveJudgeType[index] == 3)
                {
                    for (int j = 0; j < dataGridView.RowCount; j++) //find CurveData Index
                    {
                        if ((string)(dataGridView.Rows[j].Cells[1].Value) == "WaveForm" + (index - 5).ToString())
                        {
                           // dataGridView.Rows[j].Cells[5].Value = value_y.Min().ToString() + "," + value_y.Max().ToString();

                            for (int k = 0; k < min_x.Length; k++)
                            {
                                for (int kk = 0; kk < value_x.Length; kk++)
                                {
                                    if (min_x[k] == value_x[kk])//逐点比较下限范围值每个点
                                    {
                                        if (value_y[kk] >= min_y[k])
                                        {

                                        }
                                        else
                                        {
                                            tempFailCnt++;
                                            failCntRight++;
                                        }
                                    }
                                }
                            }
                            ////////////////
                            if (tempFailCnt == 0)
                            {
                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView.Rows[j].Cells[5].Style.BackColor = Color.Red;
                            }

                        }

                    }
                }
            }
        }

        private void showCurveLeft()
        {
            int pointNumbers=0;
            if (comboCurveNames.SelectedIndex > -1 && comboCurveNames.SelectedIndex < comboCurveNames.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboCurveNames.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {                                                                         //(string)comboCurveNames.SelectedText
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chartCurve.Series["Series1"].Points.Clear(); // Clear chart
                       // chartCurve.Titles["Curve"].Text = (string)comboCurveNames.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                             XData = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                             YData = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();


                            chartCurve.Series["Series1"].Points.DataBindXY(XData, YData);
                            if (curveJudgeType[0] > 0)
                            {
                                chartCurve.Series[1].Points.DataBindXY(XData_Min, YData_Min);
                                chartCurve.Series[2].Points.DataBindXY(XData_Max, YData_Max);
                            }

                            judgeCurveFuncLeft(0);

                            chartCurve.ChartAreas[0].AxisX.Minimum = XData.Min();
                            chartCurve.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

                            chartCurve.ChartAreas[1].AxisX.Minimum = XData.Min();
                            chartCurve.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";

                            chartCurve.ChartAreas[2].AxisX.Minimum = XData.Min();
                            chartCurve.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (XData.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chartCurve.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chartCurve.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chartCurve.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chartCurve.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chartCurve.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chartCurve.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                       // textBoxCurveInfo.Text = "Curve data not found!";
                        //chartCurve.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chartCurve.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chartCurve.Titles["Curve"].Text = "";
                    }

                }
            }

            Thread.Sleep(200);
            if (comboBox6.SelectedIndex > -1 && comboBox6.SelectedIndex < comboBox6.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox6.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        
                        //chart1.Titles["Curve"].Text = (string)comboBox6.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                             XData1 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                             YData1 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            //if ((XData.Length > 0) && (XData.Length == YData.Length))
                            //{
                            //    chart1.Series["Series1"].Points.Clear(); // Clear chart
                            //}
                            //else
                            //{
                            //    return;
                            //}
                            judgeCurveFuncLeft(1);


                            chart1.Series["Series1"].Points.DataBindXY(XData1, YData1);
                            if (curveJudgeType[1] > 0)
                            {
                                chart1.Series[1].Points.DataBindXY(XData1_Min, YData1_Min);
                                chart1.Series[2].Points.DataBindXY(XData1_Max, YData1_Max);
                            }

                            chart1.ChartAreas[0].AxisX.Minimum = XData1.Min();
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

                            chart1.ChartAreas[1].AxisX.Minimum = XData1.Min();
                            chart1.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";

                            chart1.ChartAreas[2].AxisX.Minimum = XData1.Min();
                            chart1.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (XData1.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart1.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart1.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart1.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart1.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart1.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart1.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart1.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart1.Titles["Curve"].Text = "";
                    }

                }
            }

            Thread.Sleep(100);
            if (comboBox7.SelectedIndex > -1 && comboBox7.SelectedIndex < comboBox7.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox7.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chart2.Series["Series1"].Points.Clear(); // Clear chart
                       // chart2.Titles["Curve"].Text = (string)comboBox7.SelectedItem; // Set chart title

                        try
                        {

                            // Get X and Y data and plot them on the chart
                             XData2 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                             YData2 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(2);

                            // chart2.ChartAreas[0].AxisX.Maximum = XData.Max();


                            chart2.Series["Series1"].Points.DataBindXY(XData2, YData2);
                            if (curveJudgeType[2] > 0)
                            {
                                chart2.Series[1].Points.DataBindXY(XData2_Min, YData2_Min);
                                chart2.Series[2].Points.DataBindXY(XData2_Max, YData2_Max);
                            }

                            chart2.ChartAreas[0].AxisX.Minimum = XData2.Min();
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";
                            chart2.ChartAreas[1].AxisX.Minimum = XData2.Min();
                            chart2.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";
                            chart2.ChartAreas[2].AxisX.Minimum = XData2.Min();
                            chart2.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (XData2.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart2.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart2.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart2.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart2.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart2.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart2.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart2.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart2.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart2.Titles["Curve"].Text = "";
                    }

                }
            }
            Thread.Sleep(200);

            if (comboBox8.SelectedIndex > -1 && comboBox8.SelectedIndex < comboBox8.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox8.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chart3.Series["Series1"].Points.Clear(); // Clear chart
                      //  chart3.Titles["Curve"].Text = (string)comboBox8.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            XData3 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            YData3 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(3);

                            chart3.Series["Series1"].Points.DataBindXY(XData3, YData3);
                            if (curveJudgeType[3] > 0)
                            {
                                chart3.Series[1].Points.DataBindXY(XData3_Min, YData3_Min);
                                chart3.Series[2].Points.DataBindXY(XData3_Max, YData3_Max);
                            }

                            chart3.ChartAreas[0].AxisX.Minimum = XData3.Min();
                            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";
                            chart3.ChartAreas[1].AxisX.Minimum = XData3.Min();
                            chart3.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";
                            chart3.ChartAreas[2].AxisX.Minimum = XData3.Min();
                            chart3.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (XData3.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart3.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart3.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart3.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart3.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart3.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart3.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart3.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart3.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart3.Titles["Curve"].Text = "";
                    }

                }
            }
            Thread.Sleep(200);
            if (comboBox9.SelectedIndex > -1 && comboBox9.SelectedIndex < comboBox9.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox9.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chart4.Series["Series1"].Points.Clear(); // Clear chart
                       // chart4.Titles["Curve"].Text = (string)comboBox9.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            XData4 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            YData4 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(4);


                            chart4.Series["Series1"].Points.DataBindXY(XData4, YData4);
                            if (curveJudgeType[4] > 0)
                            {
                                chart4.Series[1].Points.DataBindXY(XData4_Min, YData4_Min);
                                chart4.Series[2].Points.DataBindXY(XData4_Max, YData4_Max);
                            }

                            chart4.ChartAreas[0].AxisX.Minimum = XData4.Min();
                            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

                            chart4.ChartAreas[1].AxisX.Minimum = XData4.Min();
                            chart4.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";

                            chart4.ChartAreas[2].AxisX.Minimum = XData4.Min();
                            chart4.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (XData4.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart4.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart4.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart4.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart4.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart4.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart4.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart4.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart4.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart4.Titles["Curve"].Text = "";
                    }

                }
            }
            Thread.Sleep(200);
            if (comboBox10.SelectedIndex > -1 && comboBox10.SelectedIndex < comboBox10.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox10.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                      //  chart5.Series["Series1"].Points.Clear(); // Clear chart
                      //  chart5.Titles["Curve"].Text = (string)comboBox10.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            XData5 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            YData5 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(5);

                            chart5.Series["Series1"].Points.DataBindXY(XData5, YData5);

                            if (curveJudgeType[5] > 0)
                            {
                                chart5.Series[1].Points.DataBindXY(XData5_Min, YData5_Min);
                                chart5.Series[2].Points.DataBindXY(XData5_Max, YData5_Max);
                            }

                            chart5.ChartAreas[0].AxisX.Minimum = XData5.Min();
                            chart5.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";
                            chart5.ChartAreas[1].AxisX.Minimum = XData5.Min();
                            chart5.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";
                            chart5.ChartAreas[2].AxisX.Minimum = XData5.Min();
                            chart5.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (XData5.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart5.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart5.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart5.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart5.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart5.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart5.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart5.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart5.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart5.Titles["Curve"].Text = "";
                    }

                }
            }

            if (failCntLeft > 0)
            {
                textTestResultL.BackColor = Color.Red;
                textTestResultL.Text = "失败";
            }
            else
            {
                textTestResultL.BackColor = Color.Green;
                textTestResultL.Text = "成功";
            }
        }


        private void showCurveRight()
        {
            int pointNumbers = 0;
            if (comboCurveNames.SelectedIndex > -1 && comboCurveNames.SelectedIndex < comboCurveNames.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboCurveNames.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {                                                                         //(string)comboCurveNames.SelectedText
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chartCurve.Series["Series1"].Points.Clear(); // Clear chart
                        // chartCurve.Titles["Curve"].Text = (string)comboCurveNames.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            RXData = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            RYData = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();


                            chartCurve.Series["Series1"].Points.DataBindXY(RXData, RYData);

                            if (curveJudgeType[6] > 0)
                            {
                                chartCurve.Series[1].Points.DataBindXY(RXData_Min, RYData_Min);
                                chartCurve.Series[2].Points.DataBindXY(RXData_Max, RYData_Max);
                            }

                            judgeCurveFuncLeft(6);

                            chartCurve.ChartAreas[0].AxisX.Minimum = RXData.Min();
                            chartCurve.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

                            chartCurve.ChartAreas[1].AxisX.Minimum = RXData.Min();
                            chartCurve.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";

                            chartCurve.ChartAreas[2].AxisX.Minimum = RXData.Min();
                            chartCurve.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (RXData.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chartCurve.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chartCurve.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chartCurve.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chartCurve.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chartCurve.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chartCurve.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        // textBoxCurveInfo.Text = "Curve data not found!";
                        //chartCurve.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chartCurve.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chartCurve.Titles["Curve"].Text = "";
                    }

                }
            }

            Thread.Sleep(200);
            if (comboBox6.SelectedIndex > -1 && comboBox6.SelectedIndex < comboBox6.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox6.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {

                        //chart1.Titles["Curve"].Text = (string)comboBox6.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            RXData1 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            RYData1 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            //if ((XData.Length > 0) && (XData.Length == YData.Length))
                            //{
                            //    chart1.Series["Series1"].Points.Clear(); // Clear chart
                            //}
                            //else
                            //{
                            //    return;
                            //}
                            judgeCurveFuncLeft(7);


                            chart1.Series["Series1"].Points.DataBindXY(RXData1, RYData1);
                            if (curveJudgeType[7] > 0)
                            {
                                chart1.Series[1].Points.DataBindXY(RXData1_Min, RYData1_Min);
                                chart1.Series[2].Points.DataBindXY(RXData1_Max, RYData1_Max);
                            }

                            chart1.ChartAreas[0].AxisX.Minimum = RXData1.Min();
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

                            chart1.ChartAreas[1].AxisX.Minimum = RXData1.Min();
                            chart1.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";

                            chart1.ChartAreas[2].AxisX.Minimum = RXData1.Min();
                            chart1.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (RXData1.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart1.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart1.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart1.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart1.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart1.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart1.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart1.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart1.Titles["Curve"].Text = "";
                    }

                }
            }

            Thread.Sleep(100);
            if (comboBox7.SelectedIndex > -1 && comboBox7.SelectedIndex < comboBox7.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox7.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chart2.Series["Series1"].Points.Clear(); // Clear chart
                        // chart2.Titles["Curve"].Text = (string)comboBox7.SelectedItem; // Set chart title

                        try
                        {

                            // Get X and Y data and plot them on the chart
                            RXData2 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            RYData2 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(8);

                            // chart2.ChartAreas[0].AxisX.Maximum = XData.Max();


                            chart2.Series["Series1"].Points.DataBindXY(RXData2, RYData2);
                            if (curveJudgeType[8] > 0)
                            {
                                chart2.Series[1].Points.DataBindXY(RXData2_Min, RYData2_Min);
                                chart2.Series[2].Points.DataBindXY(RXData2_Max, RYData2_Max);
                            }

                            chart2.ChartAreas[0].AxisX.Minimum = RXData2.Min();
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";
                            chart2.ChartAreas[1].AxisX.Minimum = RXData2.Min();
                            chart2.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";
                            chart2.ChartAreas[2].AxisX.Minimum = RXData2.Min();
                            chart2.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (RXData2.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart2.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart2.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart2.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart2.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart2.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart2.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart2.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart2.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart2.Titles["Curve"].Text = "";
                    }

                }
            }
            Thread.Sleep(200);

            if (comboBox8.SelectedIndex > -1 && comboBox8.SelectedIndex < comboBox8.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox8.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chart3.Series["Series1"].Points.Clear(); // Clear chart
                        //  chart3.Titles["Curve"].Text = (string)comboBox8.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            RXData3 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            RYData3 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(9);

                            chart3.Series["Series1"].Points.DataBindXY(RXData3, RYData3);
                            if (curveJudgeType[9] > 0)
                            {
                                chart3.Series[1].Points.DataBindXY(RXData3_Min, RYData3_Min);
                                chart3.Series[2].Points.DataBindXY(RXData3_Max, RYData3_Max);
                            }

                            chart3.ChartAreas[0].AxisX.Minimum = RXData3.Min();
                            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";
                            chart3.ChartAreas[1].AxisX.Minimum = RXData3.Min();
                            chart3.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";
                            chart3.ChartAreas[2].AxisX.Minimum = RXData3.Min();
                            chart3.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (RXData3.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart3.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart3.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart3.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart3.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart3.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart3.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart3.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart3.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart3.Titles["Curve"].Text = "";
                    }

                }
            }
            Thread.Sleep(200);
            if (comboBox9.SelectedIndex > -1 && comboBox9.SelectedIndex < comboBox9.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox9.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //chart4.Series["Series1"].Points.Clear(); // Clear chart
                        // chart4.Titles["Curve"].Text = (string)comboBox9.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            RXData4 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            RYData4 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(10);


                            chart4.Series["Series1"].Points.DataBindXY(RXData4, RYData4);
                            if (curveJudgeType[10] > 0)
                            {
                                chart4.Series[1].Points.DataBindXY(RXData4_Min,RYData4_Min);
                                chart4.Series[2].Points.DataBindXY(RXData4_Max, RYData4_Max);
                            }

                            chart4.ChartAreas[0].AxisX.Minimum = RXData4.Min();
                            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

                            chart4.ChartAreas[1].AxisX.Minimum = RXData4.Min();
                            chart4.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";

                            chart4.ChartAreas[2].AxisX.Minimum = RXData4.Min();
                            chart4.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (RXData4.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart4.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart4.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart4.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart4.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart4.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart4.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart4.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart4.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart4.Titles["Curve"].Text = "";
                    }

                }
            }
            Thread.Sleep(200);
            if (comboBox10.SelectedIndex > -1 && comboBox10.SelectedIndex < comboBox10.Items.Count) // Make sure selection is valid
            {
                if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboBox10.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
                {
                    // Command completed successfully
                    if (json.returnData.Value<Boolean>("Found"))
                    {
                        //  chart5.Series["Series1"].Points.Clear(); // Clear chart
                        //  chart5.Titles["Curve"].Text = (string)comboBox10.SelectedItem; // Set chart title

                        try
                        {
                            // Get X and Y data and plot them on the chart
                            RXData5 = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
                            RYData5 = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

                            judgeCurveFuncLeft(11);

                            chart5.Series["Series1"].Points.DataBindXY(RXData5, RYData5);

                            if (curveJudgeType[11] > 0)
                            {
                                chart5.Series[1].Points.DataBindXY(RXData5_Min, RYData5_Min);
                                chart5.Series[2].Points.DataBindXY(RXData5_Max, RYData5_Max);
                            }

                            chart5.ChartAreas[0].AxisX.Minimum = RXData5.Min();
                            chart5.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";
                            chart5.ChartAreas[1].AxisX.Minimum = RXData5.Min();
                            chart5.ChartAreas[1].AxisX.LabelStyle.Format = "{0:0}";
                            chart5.ChartAreas[2].AxisX.Minimum = RXData5.Min();
                            chart5.ChartAreas[2].AxisX.LabelStyle.Format = "{0:0}";

                            if (RXData5.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
                            {
                                chart5.ChartAreas[0].AxisX.IsLogarithmic = true;
                                chart5.ChartAreas[1].AxisX.IsLogarithmic = true;
                                chart5.ChartAreas[2].AxisX.IsLogarithmic = true;
                            }
                            else
                            {
                                chart5.ChartAreas[0].AxisX.IsLogarithmic = false;
                                chart5.ChartAreas[1].AxisX.IsLogarithmic = false;
                                chart5.ChartAreas[2].AxisX.IsLogarithmic = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        //// textBoxCurveInfo.Text = "Curve data not found!";
                        //chart5.ChartAreas[0].AxisX.IsLogarithmic = false;
                        //foreach (var series in chart5.Series)
                        //{
                        //    series.Points.Clear();
                        //}
                        //chart5.Titles["Curve"].Text = "";
                    }

                }
            }

            if (failCntRight> 0)
            {
                textTestResultR.BackColor = Color.Red;
                textTestResultR.Text = "失败";
            }
            else
            {
                textTestResultR.BackColor = Color.Green;
                textTestResultR.Text = "成功";
            }
        }
        //private void comboCurveNames_SelectedIndexChanged(object sender, EventArgs e) // User selected a curve. Get its data and show it in the graph
        //{
        //    if (comboCurveNames.SelectedIndex > -1 && comboCurveNames.SelectedIndex < comboCurveNames.Items.Count) // Make sure selection is valid
        //    {
        //        if (SendCommandAndGetResponse("MemoryList.Get('Curve','" + (string)comboCurveNames.SelectedItem + "')")) // Send MemoryList.Get('Curve','<curve name> command and wait for response
        //        {
        //            // Command completed successfully
        //            if (json.returnData.Value<Boolean>("Found"))
        //            {
        //                //textBoxCurveInfo.Text = "Name: " + json.returnData.Curve.Value<String>("Name") + Environment.NewLine +
        //                //                        "X Unit: " + json.returnData.Curve.Value<String>("XUnit") + Environment.NewLine +
        //                //                        "Y Unit: " + json.returnData.Curve.Value<String>("YUnit") + Environment.NewLine +
        //                //                        "Z Unit: " + json.returnData.Curve.Value<String>("ZUnit") + Environment.NewLine +
        //                //                        "X Scale: " + json.returnData.Curve.Value<String>("XDataScale") + Environment.NewLine +
        //                //                        "Y Scale: " + json.returnData.Curve.Value<String>("YDataScale") + Environment.NewLine +
        //                //                        "Z Scale: " + json.returnData.Curve.Value<String>("ZDataScale"); // Update Curve Info textbox 

        //                chartCurve.Series["Series1"].Points.Clear(); // Clear chart
        //                chartCurve.Titles["Curve"].Text = (string)comboCurveNames.SelectedItem; // Set chart title

        //                chart1.Series["Series1"].Points.Clear(); // Clear chart
        //                chart1.Titles["Curve"].Text = (string)comboCurveNames.SelectedItem; // Set chart title

        //                chart2.Series["Series1"].Points.Clear(); // Clear chart
        //                chart2.Titles["Curve"].Text = (string)comboCurveNames.SelectedItem; // Set chart title

        //                chart3.Series["Series1"].Points.Clear(); // Clear chart
        //                chart3.Titles["Curve"].Text = (string)comboCurveNames.SelectedItem; // Set chart title

        //                try
        //                {
        //                    // Get X and Y data and plot them on the chart
        //                    double[] XData = json.returnData.Curve.Value<JArray>("XData").ToObject<double[]>();
        //                    double[] YData = json.returnData.Curve.Value<JArray>("YData").ToObject<double[]>();

        //                    for (int i = 0; i < XData.Length; i++)
        //                    {
        //                        chartCurve.Series["Series1"].Points.AddXY(XData[i], YData[i]);

        //                        chart1.Series["Series1"].Points.AddXY(XData[i], YData[i]);
        //                        chart2.Series["Series1"].Points.AddXY(XData[i], YData[i]);
        //                        chart3.Series["Series1"].Points.AddXY(XData[i], YData[i]);

        //                    }

        //                    chartCurve.ChartAreas[0].AxisX.Minimum = XData.Min();
        //                    chartCurve.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

        //                    chart1.ChartAreas[0].AxisX.Minimum = XData.Min();
        //                    chart2.ChartAreas[0].AxisX.LabelStyle.Format = "{0:0}";

        //                    if (XData.Min() > 0 && json.returnData.Curve.Value<String>("XAxisScale") == "Log")
        //                    {
        //                        chartCurve.ChartAreas[0].AxisX.IsLogarithmic = true;

        //                        chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
        //                        chart2.ChartAreas[0].AxisX.IsLogarithmic = true;
        //                        chart3.ChartAreas[0].AxisX.IsLogarithmic = true;
        //                    }
        //                    else
        //                    {
        //                        chartCurve.ChartAreas[0].AxisX.IsLogarithmic = false;
        //                        chart1.ChartAreas[0].AxisX.IsLogarithmic = false;
        //                        chart2.ChartAreas[0].AxisX.IsLogarithmic = false;
        //                        chart3.ChartAreas[0].AxisX.IsLogarithmic = false;
        //                    }
        //                }
        //                catch (Exception)
        //                {
        //                    //MessageBox.Show(x.Message.ToString());
        //                    // Not all curves will have array data. Ignore exceptions.
        //                }
        //            }
        //            else
        //            {
        //                textBoxCurveInfo.Text = "Curve data not found!";
        //                chartCurve.ChartAreas[0].AxisX.IsLogarithmic = false;
        //                foreach (var series in chartCurve.Series)
        //                {
        //                    series.Points.Clear();
        //                }
        //                chartCurve.Titles["Curve"].Text = "";
        //            }

        //        }
        //    }
        //}

        private void btnSetLotNumber_Click(object sender, EventArgs e)
        {
            if (SendCommandAndGetResponse("SoundCheck.SetLotNumber('" + textBoxLotNumber.Text + "')")) // Send SoundCheck.SetLotNumber command and wait for result.
            {
                // Command completed successfully
                AppendLogMessage("Lot number set.");
            }
            else
            {
                // Command did not complete successfully
                if (GetErrorType() == "Timeout") // Check if command timed out
                {
                    AppendLogMessage("Command failed; timed out!");
                }
            }

        }

        private void btnSetSerialNumber_Click(object sender, EventArgs e)
        {
            if (SendCommandAndGetResponse("SoundCheck.SetSerialNumber('" + textBoxSerialNumber.Text + "')")) // Send SoundCheck.SetSerialNumber command and wait for result.
            {
                // Command completed successfully
                AppendLogMessage("Serial number set.");
            }
            else
            {
                // Command did not complete successfully
                if (GetErrorType() == "Timeout") // Check if command timed out
                {
                    AppendLogMessage("Command failed; timed out!");
                }
            }
        }

        private void btnExitSoundCheck_Click(object sender, EventArgs e)
        {
            if (SendCommandAndGetResponse("SoundCheck.Exit")) // Send SoundCheck.Exit command and wait for result.
            {
                // Command completed successfully
                AppendLogMessage("SoundCheck exited.");

                SetConnectedState(false);
            }
            else
            {
                // Command did not complete successfully
                if (json.Value<string>("errorType") == "Timeout") // Check if command timed out
                {
                    AppendLogMessage("Command failed; timed out!");
                }
            }
        }

        public void AppendLogMessage(string message)
        { textBoxLog.AppendText(Environment.NewLine + DateTime.Now.ToString() + ": " + message); }

        public void AppendLogMessageRight(string message)
        { textBoxLog.AppendText(Environment.NewLine + DateTime.Now.ToString() + "Right: " + message); }

        // These methods help parse the response from SoundCheck 

        // The following methods are common for every command's response
        private bool GetCommandCompleted() { return json.Value<Boolean>("cmdCompleted"); }      // Whether or not the command completed
        private string GetReturnType() { return json.Value<string>("returnType"); }             // Get the data type of the data returned by the command
        private string GetErrorType() { return json.Value<string>("errorType"); }               // If command did not complete, what error type caused the failure
        private string GetErrorDescription() { return json.Value<string>("errorDescription"); } // Error description

        // The following methods help get command specific data from the response JSON

        // Boolean Data
        private bool GetReturnDataBoolean() { return json.returnData.Value<Boolean>("Value"); }     // Boolean Data
        private bool GetReturnDataBoolean(string dataName) { return json.returnData.Value<Boolean>(dataName); }     // Overload: Boolean data by name, where the return data contains more than just the boolean field

        // Integer Data
        private int GetReturnDataInteger() { return json.returnData.Value<Int32>("Value"); }        // Integer Data
        private int GetReturnDataInteger(string dataName) { return json.returnData.Value<Int32>(dataName); }     // Overload: Integer data by name, where the return data contains more than just the integer field

        // Double Date
        private double GetReturnDataDouble() { return json.returnData.Value<Double>("Value"); }     // Double Data
        private double GetReturnDataDouble(string dataName) { return json.returnData.Value<Double>(dataName); }     // Overload: Double data by name, where the return data contains more than just the double field

        // String Data
        private string GetReturnDataString() { return json.returnData.Value<String>("Value"); }     // String Data
        private string[] GetReturnDataStringArray() { return json.returnData.Value<JArray>("Value").ToObject<string[]>(); }     // String Array Data

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if (textCom1Status.Text == "连接成功")
                {
                    // RecvCom1Thread.Abort();
                    com1.Close();
                }
                if (textCom2Status.Text == "连接成功")
                {
                    // RecvCom2Thread.Abort();
                    com2.Close();
                }

                if (textCom3Status.Text == "连接成功")
                {
                    // RecvCom3Thread.Abort();
                    com3.Close();
                }

                if (textCom4Status.Text == "连接成功")
                {
                    // RecvCom4Thread.Abort();
                    com4.Close();
                }

                if (textCom5Status.Text == "连接成功")
                {
                    //RecvCom5Thread.Abort();
                    com5.Close();
                }


                if (textVitualPortStatus.Text == "连接成功")
                {
                    //RecvCom6Thread.Abort();
                    com6.Close();
                }

                testings = false;

                if (LeftTestThread.ThreadState == ThreadState.Running)
                {
                    LeftTestThread.Abort();

                }

                if (RightTestThread.ThreadState == ThreadState.Running)
                {
                    RightTestThread.Abort();

                }
            }
            catch { }

            if (textSCStatus.Text == "连接成功")
            { 
             SendCommandAndGetResponse("SoundCheck.Exit");
            }
        }

        private void buttonSaveConfig_Click(object sender, EventArgs e)
        {
            string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";
            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            try
            {
                iniFile.IniWriteValue(fName, "基准", "SerialPort1", comboBox1.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "SerialPort2", comboBox2.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "SerialPort3", comboBox3.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "SerialPort4", comboBox4.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "SerialPort5", comboBox5.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "VirtualPort", comboBoxVirtual.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "SouncCheckPath", textBoxExePath.Text);
                iniFile.IniWriteValue(fName, "基准", "SouncCheckStatus", comboBoxWindowState.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "SeqPath", textBoxSeqFilePath.Text);
                iniFile.IniWriteValue(fName, "基准", "SoundCheckConfigFileName", textSCConfigFileName.Text);

                iniFile.IniWriteValue(fName, "基准", "IP", ConnectToIPAddress.Text);
                iniFile.IniWriteValue(fName, "基准", "IP_Port", ConnectToPort.Text);

                MessageBox.Show("保存成功");
            }
            catch
            {
                MessageBox.Show("保存失败");
                return;
            }


        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if ((tabControl2.SelectedIndex >= 1) && (tabControl2.SelectedIndex <= 3) && (flagPassrowd == false))
            //{
            //    Login registerForm = new Login();
            //    registerForm.ShowDialog();
            //    if (registerForm.DialogResult == DialogResult.OK)
            //    {
            //        flagPassrowd = true;
            //    }
            //    else
            //    {
            //        tabControl2.SelectedIndex = 0;
            //    }
            //}

            if (tabControl2.SelectedIndex == 3)
            {
                ReadCsvToRangeView(rangePath, 1, dataGridViewRange);
            }
        }

        private void timerConnectSC_Tick(object sender, EventArgs e)
        {
            timeCount1++;

            if (timeCount1 == 1)
            {
               // InitDevice();
            }
            if (timeCount1 >= 6)
            {
                if (textSCStatus.Text !="连接成功")
                {
                    //btnConnectToServer.PerformClick();
                    ConnectToServer();
                }
            }
        }

        private void InitDevice()
        {
            //open device1
            try
            {
                if (textCom1Status.Text != "连接成功")
                {
                    if (com1.Open(("COM" + (comboBox1.SelectedIndex + 1).ToString()), 9600) == 1)
                    {
                        textCom1Status.Text = "连接成功";
                        textCom1Status.BackColor = Color.Green;
                    }
                    else
                    {
                        textCom1Status.Text = "连接失败";
                        textCom1Status.BackColor = Color.White;
                    }
                }
            }
            catch
            {
                textCom1Status.Text = "连接失败";
                textCom1Status.BackColor = Color.White;
            }

            try
            {
                if (textCom2Status.Text != "连接成功")
                {
                    if (com2.Open(("COM" + (comboBox2.SelectedIndex + 1).ToString()), 9600) == 1)
                    {
                        textCom2Status.Text = "连接成功";
                        textCom2Status.BackColor = Color.Green;
                    }
                    else
                    {
                        textCom2Status.Text = "连接失败";
                        textCom2Status.BackColor = Color.White;
                    }
                }
            }
            catch
            {
                textCom2Status.Text = "连接失败";
                textCom2Status.BackColor = Color.White;
            }

            try
            {
                if (textCom3Status.Text != "连接成功")
                {
                    if (com3.Open(("COM" + (comboBox3.SelectedIndex + 1).ToString()), 9600) == 1)
                    {
                        textCom3Status.Text = "连接成功";
                        textCom3Status.BackColor = Color.Green;
                    }
                    else
                    {
                        textCom3Status.Text = "连接失败";
                        textCom3Status.BackColor = Color.White;
                    }
                }
            }
            catch
            {
                textCom3Status.Text = "连接失败";
                textCom3Status.BackColor = Color.White;
            }

            try
            {
                if (textCom4Status.Text != "连接成功")
                {
                    if (com4.Open(("COM" + (comboBox4.SelectedIndex + 1).ToString()), 115200) == 1)
                    {
                        textCom4Status.Text = "连接成功";
                        textCom4Status.BackColor = Color.Green;
                    }
                    else
                    {
                        textCom4Status.Text = "连接失败";
                        textCom4Status.BackColor = Color.White;
                    }
                }
            }
            catch
            {
                textCom4Status.Text = "连接失败";
                textCom4Status.BackColor = Color.White;
            }


            try
            {
                if (textCom5Status.Text != "连接成功")
                {
                    if (com5.Open(("COM" + (comboBox5.SelectedIndex + 1).ToString()), 115200) == 1)
                    {
                        textCom5Status.Text = "连接成功";
                        textCom5Status.BackColor = Color.Green;
                    }
                    else
                    {
                        textCom5Status.Text = "连接失败";
                        textCom5Status.BackColor = Color.White;
                    }
                }
            }
            catch
            {
                textCom5Status.Text = "连接失败";
                textCom5Status.BackColor = Color.White;
            }

            try
            {
                if (textVitualPortStatus.Text != "连接成功")
                {
                    if (com6.Open(("COM" + (comboBoxVirtual.SelectedIndex + 1).ToString()), 9600) == 1)
                    {
                        textVitualPortStatus.Text = "连接成功";
                        textVitualPortStatus.BackColor = Color.Green;
                    }
                    else
                    {
                        textVitualPortStatus.Text = "连接失败";
                        textVitualPortStatus.BackColor = Color.White;
                    }
                }
            }
            catch
            {
                textVitualPortStatus.Text = "连接失败";
                textVitualPortStatus.BackColor = Color.White;
            }
        }

        private void DeInitDevice()
        {
            if (textCom1Status.Text == "连接成功")
            {
                textCom1Status.BackColor = Color.White;
                textCom1Status.Text = "连接失败";
                Thread.Sleep(10);
                com1.Close();
            }
            if (textCom2Status.Text == "连接成功")
            {
                textCom2Status.BackColor = Color.White;
                textCom2Status.Text = "连接失败";
                Thread.Sleep(10);
                com2.Close();
            }

            if (textCom3Status.Text == "连接成功")
            {
                textCom3Status.BackColor = Color.White;
                textCom3Status.Text = "连接失败";
                Thread.Sleep(10);
                com3.Close();
            }

            if (textCom4Status.Text == "连接成功")
            {
                textCom4Status.BackColor = Color.White;
                textCom4Status.Text = "连接失败";
                Thread.Sleep(10);
                com4.Close();
            }

            if (textCom5Status.Text == "连接成功")
            {
                textCom5Status.BackColor = Color.White;
                textCom5Status.Text = "连接失败";
                Thread.Sleep(10);
                com5.Close();
            }

            if (textVitualPortStatus.Text== "连接成功")
            {
                textVitualPortStatus.BackColor = Color.White;
                textVitualPortStatus.Text = "连接失败";
                Thread.Sleep(10);
                com6.Close();
            }
        }

        /// <summary>
        /// //////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            InitDevice();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DeInitDevice();
        }

        private void 创建产品配置ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void label5_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private bool ControlSwitch(string cmd)
        {
            try
            {
                if (textCom3Status.Text == "连接成功")
                {
                    cmd = cmd.Trim();
                    com3.ClearInBuffer();
                    com3.SendStringData(cmd + "\r\n");
                }
            }
            catch
            { }
            return false;
        }
        private void buttonStart_Click(object sender, EventArgs e)
        {
            buttonStart.Enabled = false;
            textTestResultL.Text = "开始";
            textTestResultL.BackColor = Color.Azure;

            busyFlag[0] = false;
            busyFlag[1] = false;

            AppendLogMessage("测试准备:关闭所有开关");

            if (textCom3Status.Text == "连接成功")
            {
                ControlSwitch("ALL_OFF");
                ControlSwitch("ALL1_OFF");
                AppendLogMessage("ALL_OFF");
                AppendLogMessage("ALL1_OFF");
            }
            else
            {
                AppendLogMessage("COM3 未连接");
            }

            for (int j = 0; j < dataGridView.RowCount; j++)
            {
                   dataGridView.Rows[j].Cells[4].Style.BackColor = Color.White;
                   dataGridView.Rows[j].Cells[5].Style.BackColor = Color.White;
                   dataGridView.Rows[j].Cells[4].Value = "";
                   dataGridView.Rows[j].Cells[5].Value = "";
            }
            testings = true;

            LeftTestThread  = new Thread(LeftChanTest);
            RightTestThread = new Thread(RightChanTest);

            LeftTestThread.Start();
           // RightTestThread.Start();
        }

        private void LeftChanTest()
        {
            int cnt;
            int len = 0;
            string read;
            double times = 0;
            double retryTimes = 0;
            bool resultFlag = false;
            int tempcnt;
            try
            {
                DataGridView dv = dataGridViewLeft;//LeftChannel

                for (int i = 0; i < dv.RowCount - 1; i++)
                {
                    Thread.Sleep(100);
                    if (testings == true)
                    {
                        if ((bool)(dv.Rows[i].Cells[0].Value) == true)
                        {
                            // AppendLogMessage((string)(dv.Rows[i].Cells[1].Value));

                            if ((string)(dv.Rows[i].Cells[1].Value) == "ShowDialog")
                            {
                                MessageBox.Show((string)(dv.Rows[i].Cells[2].Value));
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "LogMsg")
                            {
                                AppendLogMessage((string)(dv.Rows[i].Cells[2].Value));
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "WatingToTest")
                            {
                                while (busyFlag[1] == true)
                                {
                                    AppendLogMessage("LeftChan:WatingToTest");
                                    Thread.Sleep(500);
                                }

                                AppendLogMessage("LeftChan:StartToTest");
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SetBusy")
                            {
                                busyFlag[0] = true;
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SetReady")
                            {
                                busyFlag[0] = false;
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "LoopTest")
                            {
                                i = 0;
                                AppendLogMessage("Left:LoopTest");
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "Wait")
                            {
                                try
                                {
                                    Thread.Sleep(int.Parse((string)(dv.Rows[i].Cells[2].Value)));
                                }
                                catch
                                {
                                    Thread.Sleep(1000);
                                    AppendLogMessage("wait 等待时间输入错误，用默认1000执行");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "RunSeq")
                            {
                                //seqPath= (string)(dataGridView2.Rows[i].Cells[2].Value);
                                if (textSCStatus.Text == "连接成功")
                                {
                                    RunSeqThread = new Thread(LeftSeqRun);
                                    RunSeqThread.Start();
                                }

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX1/COM1_SEND")
                            {
                                if (textCom1Status.Text == "连接成功")
                                {
                                    com1.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessage("BOX1/COM1_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessage("BOX1/COM1_SEND发送失败:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX1/COM1_SEND_GET")
                            {

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX1/COM1_GET")
                            {
                                resultFlag = false;
                                read = "";
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }

                                if (times == -1)//wait without timeout
                                {
                                    read = "";
                                    tempcnt = 0;
                                    while (testings == true)
                                    {
                                        Thread.Sleep(200);
                                        if (tempcnt % 5 == 0)
                                        {
                                            AppendLogMessage("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        }
                                        tempcnt++;

                                        try
                                        {
                                            read += com1.ReadStringData().Trim().Replace("\0", "");
                                            len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                            if (len >= 0)
                                            {
                                                AppendLogMessage("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                read = "";

                                                break;
                                            }
                                        }
                                        catch { }

                                    }
                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            if (testings == true)
                                            {
                                                Thread.Sleep(200);
                                                if (j % 5 == 0)
                                                {
                                                    AppendLogMessage("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                                }

                                                try
                                                {
                                                    read += com1.ReadStringData().Trim().Replace("\0", "");
                                                    len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                                    if (len >= 0)
                                                    {
                                                        AppendLogMessage("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                        read = "";
                                                        resultFlag = true;
                                                        break;
                                                    }
                                                }
                                                catch { }

                                            }
                                            else
                                            {
                                                return;
                                            }
                                        }

                                        if (resultFlag == true)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessage("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessage("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX2/COM2_SEND")
                            {
                                if (textCom2Status.Text == "连接成功")
                                {
                                    com2.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessage("BOX2/COM2_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessage("BOX2/COM2_SEND发送失败:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX2/COM2_SEND_GET")
                            {

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX2/COM2_GET")
                            {
                                resultFlag = false;
                                read = "";
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {
                                    read = "";
                                    tempcnt = 0;
                                    while (testings == true)
                                    {
                                        Thread.Sleep(200);

                                        if (tempcnt % 5 == 0)
                                        {
                                            AppendLogMessage("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        }
                                        tempcnt++;

                                        try
                                        {
                                            read += com2.ReadStringData().Trim().Replace("\0", "");
                                            len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                            if (len >= 0)
                                            {
                                                AppendLogMessage("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                read = "";

                                                break;
                                            }
                                        }
                                        catch { }

                                    }
                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            Thread.Sleep(200);
                                            if (j % 5 == 0)
                                            {
                                                AppendLogMessage("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                            }
                                            try
                                            {
                                                read += com2.ReadStringData().Trim().Replace("\0", "");
                                                len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                                if (len >= 0)
                                                {
                                                    AppendLogMessage("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                    read = "";
                                                    resultFlag = true;
                                                    break;
                                                }
                                            }
                                            catch { }

                                        }

                                        if (resultFlag == true)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessage("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessage("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SWITCH/COM3_SEND")
                            {
                                if (textCom3Status.Text == "连接成功")
                                {
                                    com3.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessage("SWITCH/COM3_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessage("SWITCH/COM3_SEND:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SWITCH/COM3_SEND_GET")
                            {

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT1/COM4_SEND")
                            {
                                if (textCom4Status.Text == "连接成功")
                                {
                                    com4.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessage("BT1/COM4_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessage("BT1/COM4_SEND:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT1/COM4_SEND_GET")
                            {
                                resultFlag = false;
                                read = "";
                                tempcnt = 0;
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {

                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {
                                        AppendLogMessage("BT1/COM4_SEND_GET(SEND)：" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        com4.SendStringData((string)(dv.Rows[i].Cells[2].Value));

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            if (testings == true)
                                            {
                                                Thread.Sleep(200);
                                                if (j % 5 == 0)
                                                {
                                                    AppendLogMessage("BT1/COM4_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "...");
                                                }

                                                try
                                                {
                                                    read += com4.ReadStringData().Trim().Replace("\0", "");
                                                    len = read.IndexOf((string)(dv.Rows[i].Cells[3].Value));
                                                    if (len >= 0)
                                                    {
                                                        AppendLogMessage("BT1/COM4_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "成功");
                                                        read = "";
                                                        resultFlag = true;
                                                        break;
                                                    }
                                                }
                                                catch { }

                                            }
                                            else
                                            {
                                                return;
                                            }
                                        }

                                        if (resultFlag == true)
                                        {

                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessage("BT1/COM4_SEND_GET：" + (string)(dv.Rows[i].Cells[3].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessage("BT1/COM4_SEND_GET(GET):" + (string)(dv.Rows[i].Cells[3].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT2/COM5_SEND")
                            {
                                if (textCom5Status.Text == "连接成功")
                                {
                                    com5.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessage("BT2/COM5_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessage("BT2/COM5:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT2/COM5_SEND_GET")
                            {
                                resultFlag = false;
                                read = "";
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {

                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {
                                        AppendLogMessage("BT2/COM5_SEND_GET(SEND)：" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        com4.SendStringData((string)(dv.Rows[i].Cells[2].Value));

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            if (testings == true)
                                            {
                                                Thread.Sleep(200);
                                                if (j % 5 == 0)
                                                {
                                                    AppendLogMessage("BT2/COM5_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "...");
                                                }

                                                try
                                                {
                                                    read += com5.ReadStringData().Trim().Replace("\0", "");
                                                    len = read.IndexOf((string)(dv.Rows[i].Cells[3].Value));
                                                    if (len >= 0)
                                                    {
                                                        AppendLogMessage("BT2/COM5_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "成功");
                                                        read = "";
                                                        resultFlag = true;
                                                        break;
                                                    }
                                                }
                                                catch { }

                                            }
                                            else
                                            {
                                                return;
                                            }
                                        }

                                        if (resultFlag == true)
                                        {

                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessage("BT2/COM5_SEND_GET：" + (string)(dv.Rows[i].Cells[3].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessage("BT2/COM5_SEND_GET(GET):" + (string)(dv.Rows[i].Cells[3].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "VP/COM7_SEND")
                            {
                                if (textVitualPortStatus.Text == "连接成功")
                                {
                                    com6.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessage("VP/COM7_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessage("VP/COM7_SEND:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "VP/COM7_GET")
                            {
                                resultFlag = false;
                                read = "";
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {
                                    read = "";
                                    tempcnt = 0;
                                    while (testings == true)
                                    {
                                        Thread.Sleep(200);
                                        if (tempcnt % 5 == 0)
                                        {
                                            AppendLogMessage("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        }
                                        tempcnt++;

                                        try
                                        {
                                            read += com6.ReadStringData().Trim().Replace("\0", "");
                                            len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                            if (len >= 0)
                                            {
                                                AppendLogMessage("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                read = "";

                                                break;
                                            }
                                        }
                                        catch { }


                                    }
                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            if (testings == true)
                                            {
                                                Thread.Sleep(200);
                                                if (j % 5 == 0)
                                                {
                                                    AppendLogMessage("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                                }

                                                try
                                                {
                                                    read += com6.ReadStringData().Trim().Replace("\0", "");
                                                    len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                                    if (len >= 0)
                                                    {
                                                        AppendLogMessage("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                        read = "";
                                                        resultFlag = true;
                                                        break;
                                                    }
                                                }
                                                catch { }

                                            }
                                            else
                                            {
                                                return;
                                            }
                                        }

                                        if (resultFlag == true)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessage("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessage("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                buttonStart.Enabled = true;
            }
            catch { }

        }
        private void RightChanTest()
        {
            int cnt;
            int len = 0;
            string read;
            double times = 0;
            double retryTimes = 0;
            bool resultFlag = false;
            int tempcnt;
            try
            {
                DataGridView dv = dataGridViewRight;//LeftChannel

                for (int i = 0; i < dv.RowCount - 1; i++)
                {
                    Thread.Sleep(100);
                    if (testings == true)
                    {
                        if ((bool)(dv.Rows[i].Cells[0].Value) == true)
                        {
                           // AppendLogMessage((string)(dv.Rows[i].Cells[1].Value));

                            if ((string)(dv.Rows[i].Cells[1].Value) == "ShowDialog")
                            {
                                MessageBox.Show((string)(dv.Rows[i].Cells[2].Value));
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "LogMsg")
                            {
                                AppendLogMessageRight((string)(dv.Rows[i].Cells[2].Value));
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "WatingToTest")
                            {
                                while (busyFlag[0] == true)
                                {
                                    AppendLogMessageRight("WatingToTest");
                                    Thread.Sleep(500);
                                }

                                AppendLogMessageRight("StartToTest");
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SetBusy")
                            {
                                busyFlag[1] = true;
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SetReady")
                            {
                                busyFlag[1] = false;
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "LoopTest")
                            {
                                i = 0;
                                AppendLogMessageRight("LoopTest");
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "Wait")
                            {
                                try
                                {
                                    Thread.Sleep(int.Parse((string)(dv.Rows[i].Cells[2].Value)));
                                }
                                catch
                                {
                                    Thread.Sleep(1000);
                                    AppendLogMessageRight("wait 等待时间输入错误，用默认1000执行");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "RunSeq")
                            {
                                //seqPath= (string)(dataGridView2.Rows[i].Cells[2].Value);
                                if (textSCStatus.Text == "连接成功")
                                {
                                    RunSeqThread = new Thread(RightSeqRun);
                                    RunSeqThread.Start();
                                }

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX1/COM1_SEND")
                            {
                                if (textCom1Status.Text == "连接成功")
                                {
                                    com1.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessageRight("BOX1/COM1_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessageRight("BOX1/COM1_SEND发送失败:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX1/COM1_SEND_GET")
                            {

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX1/COM1_GET")
                            {
                                resultFlag = false;
                                read = "";
                                tempcnt = 0;
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {
                                    read = "";
                                    while (testings == true)
                                    {
                                        Thread.Sleep(200);
                                        try
                                        {
                                            read += com1.ReadStringData().Trim().Replace("\0", "");
                                            len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                            if (len >= 0)
                                            {
                                                AppendLogMessageRight("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                read = "";

                                                break;
                                            }
                                        }
                                        catch { }
                                        if (tempcnt % 5 == 0)
                                        {
                                            AppendLogMessageRight("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        }

                                        tempcnt++;
                                        
                                    }
                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            Thread.Sleep(200);
                                            try
                                            {
                                                read += com1.ReadStringData().Trim().Replace("\0", "");
                                                len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                                if (len >= 0)
                                                {
                                                    AppendLogMessageRight("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                    read = "";
                                                    resultFlag = true;
                                                    break;
                                                }
                                            }
                                            catch { }
                                            if (j % 5 == 0)
                                            {
                                                AppendLogMessageRight("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                            }
                                        }

                                        if (resultFlag == true)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessageRight("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessageRight("BOX1/COM1_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX2/COM2_SEND")
                            {
                                if (textCom2Status.Text == "连接成功")
                                {
                                    com2.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessageRight("BOX2/COM2_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessageRight("BOX2/COM2_SEND发送失败:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX2/COM2_SEND_GET")
                            {

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BOX2/COM2_GET")
                            {
                                resultFlag = false;
                                read = "";
                                tempcnt = 0;
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {
                                    read = "";
                                    while (testings == true)
                                    {
                                        Thread.Sleep(200);
                                        try
                                        {

                                            if (tempcnt % 5 == 0)
                                            {
                                                AppendLogMessageRight("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                            }
                                            tempcnt++;

                                            read += com2.ReadStringData().Trim().Replace("\0", "");
                                            len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                            if (len >= 0)
                                            {
                                                AppendLogMessageRight("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                read = "";

                                                break;
                                            }
                                        }
                                        catch { }


                                    }
                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            Thread.Sleep(200);
                                            if (j % 5 == 0)
                                            {
                                                AppendLogMessageRight("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                            }

                                            try
                                            {
                                                read += com2.ReadStringData().Trim().Replace("\0", "");
                                                len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                                if (len >= 0)
                                                {
                                                    AppendLogMessageRight("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                    read = "";
                                                    resultFlag = true;
                                                    break;
                                                }
                                            }
                                            catch { }
                                            
                                        }

                                        if (resultFlag == true)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessageRight("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessageRight("BOX2/COM2_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SWITCH/COM3_SEND")
                            {
                                if (textCom3Status.Text == "连接成功")
                                {
                                    com3.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessageRight("SWITCH/COM3_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessageRight("SWITCH/COM3_SEND:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "SWITCH/COM3_SEND_GET")
                            {

                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT1/COM4_SEND")
                            {
                                if (textCom4Status.Text == "连接成功")
                                {
                                    com4.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessageRight("BT1/COM4_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessageRight("BT1/COM4_SEND:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT1/COM4_SEND_GET")
                            {
                                resultFlag = false;
                                read = "";
                                tempcnt = 0;
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {

                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {
                                        AppendLogMessageRight("BT1/COM4_SEND_GET(SEND)：" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        com4.SendStringData((string)(dv.Rows[i].Cells[2].Value));

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            Thread.Sleep(200);

                                            if (j % 5 == 0)
                                            {
                                                AppendLogMessageRight("BT1/COM4_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "...");
                                            }

                                            try
                                            {
                                                read += com4.ReadStringData().Trim().Replace("\0", "");
                                                len = read.IndexOf((string)(dv.Rows[i].Cells[3].Value));
                                                if (len >= 0)
                                                {
                                                    AppendLogMessageRight("BT1/COM4_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "成功");
                                                    read = "";
                                                    resultFlag = true;
                                                    break;
                                                }
                                            }
                                            catch { }
                                           
                                        }

                                        if (resultFlag == true)
                                        {

                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessageRight("BT1/COM4_SEND_GET：" + (string)(dv.Rows[i].Cells[3].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessageRight("BT1/COM4_SEND_GET(GET):" + (string)(dv.Rows[i].Cells[3].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT2/COM5_SEND")
                            {
                                if (textCom5Status.Text == "连接成功")
                                {
                                    com5.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessageRight("BOX2/COM5_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessageRight("BT2/COM5:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "BT2/COM5_SEND_GET")
                            {
                                resultFlag = false;
                                read = "";
                                tempcnt = 0;
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {

                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {
                                        AppendLogMessageRight("BT2/COM5_SEND_GET(SEND)：" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                        com4.SendStringData((string)(dv.Rows[i].Cells[2].Value));

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            Thread.Sleep(200);

                                            if (j % 5 == 0)
                                            {
                                                AppendLogMessageRight("BT2/COM5_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "...");
                                            }

                                            try
                                            {
                                                read += com5.ReadStringData().Trim().Replace("\0", "");
                                                len = read.IndexOf((string)(dv.Rows[i].Cells[3].Value));
                                                if (len >= 0)
                                                {
                                                    AppendLogMessageRight("BT2/COM5_SEND_GET(GET)：" + (string)(dv.Rows[i].Cells[3].Value) + "成功");
                                                    read = "";
                                                    resultFlag = true;
                                                    break;
                                                }
                                            }
                                            catch { }
                                          
                                        }

                                        if (resultFlag == true)
                                        {

                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessageRight("BT2/COM5_SEND_GET：" + (string)(dv.Rows[i].Cells[3].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessageRight("BT2/COM5_SEND_GET(GET):" + (string)(dv.Rows[i].Cells[3].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "VP/COM7_SEND")
                            {
                                if (textVitualPortStatus.Text == "连接成功")
                                {
                                    com6.SendStringData((string)(dv.Rows[i].Cells[2].Value));
                                    AppendLogMessageRight("VP/COM7_SEND:" + (string)(dv.Rows[i].Cells[2].Value));
                                }
                                else
                                {
                                    AppendLogMessageRight("VP/COM7_SEND:端口未打开");
                                }
                            }
                            else if ((string)(dv.Rows[i].Cells[1].Value) == "VP/COM7_GET")
                            {
                                resultFlag = false;
                                read = "";
                                try//cal looptimes
                                {
                                    times = double.Parse((string)(dv.Rows[i].Cells[4].Value));
                                }
                                catch
                                {
                                    times = 3000;
                                }
                                cnt = (int)times / 200;

                                try//cal retry times
                                {
                                    retryTimes = int.Parse((string)(dv.Rows[i].Cells[6].Value));
                                }
                                catch
                                {
                                    retryTimes = 1;
                                }



                                if (times == -1)//wait without timeout
                                {
                                    read = "";
                                    while (testings == true)
                                    {
                                        Thread.Sleep(200);
                                        try
                                        {
                                            read += com6.ReadStringData().Trim().Replace("\0", "");
                                            len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                            if (len >= 0)
                                            {
                                                AppendLogMessageRight("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                read = "";

                                                break;
                                            }
                                        }
                                        catch { }

                                        AppendLogMessageRight("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                    }
                                }
                                else
                                {
                                    resultFlag = false;
                                    read = "";
                                    for (int k = 0; k < retryTimes; k++)
                                    {

                                        for (int j = 0; j < cnt; j++)
                                        {
                                            Thread.Sleep(200);
                                            if (j % 5 == 0)
                                            {
                                                AppendLogMessageRight("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "...");
                                            }
                                            try
                                            {
                                                read += com6.ReadStringData().Trim().Replace("\0", "");
                                                len = read.IndexOf((string)(dv.Rows[i].Cells[2].Value));
                                                if (len >= 0)
                                                {
                                                    AppendLogMessageRight("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "成功");
                                                    read = "";
                                                    resultFlag = true;
                                                    break;
                                                }
                                            }
                                            catch { }
                                      
                                        }

                                        if (resultFlag == true)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            AppendLogMessageRight("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测...");
                                        }
                                    }
                                    if (resultFlag == true)
                                    {

                                    }
                                    else
                                    {
                                        //if ((string)(dv.Rows[i].Cells[7].Value) == "停止")
                                        //{
                                        //    buttonStart.Enabled = true;
                                        //    textTestResult.Text = "停止";
                                        //    textTestResult.BackColor = Color.Red;
                                        //    AppendLogMessageRight("VP/COM7_GET:" + (string)(dv.Rows[i].Cells[2].Value) + "重测失败，退出...");
                                        //    return;
                                        //}

                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                buttonStart.Enabled = true;
            }
            catch { }

        }
        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void InitStepTable(DataGridView dv)
        {
            dv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dv.RowHeadersVisible = false;

            DataGridViewCheckBoxColumn cb0 = new DataGridViewCheckBoxColumn();
            cb0.HeaderText = "启用";
            cb0.Name = "cb_check";
            cb0.TrueValue = true;
            cb0.FalseValue = false;
            //column9.DataPropertyName = "IsScienceNature";
            cb0.DataPropertyName = "IsChecked";


            DataGridViewComboBoxColumn cb1 = new DataGridViewComboBoxColumn();
            cb1.HeaderText = "步骤名称";

            for (int i = 0; i < stepNames.Count(); i++)
            {
                cb1.Items.Add(stepNames[i]);
            }
            cb1.Width = 200;


            DataGridViewTextBoxColumn cb2 = new DataGridViewTextBoxColumn();
            cb2.HeaderText = "参数";
            cb2.Width = 180;

            DataGridViewTextBoxColumn cb3 = new DataGridViewTextBoxColumn();
            cb3.HeaderText = "返回值";
            cb3.Width = 180;


            DataGridViewTextBoxColumn cb4 = new DataGridViewTextBoxColumn();
            cb4.HeaderText = "超时ms";

            DataGridViewComboBoxColumn cb5 = new DataGridViewComboBoxColumn();
            cb5.HeaderText = "失败后动作";
            for (int i = 0; i < stepTackleError.Count(); i++)
            {
                cb5.Items.Add(stepTackleError[i]);
            }



            DataGridViewTextBoxColumn cb6 = new DataGridViewTextBoxColumn();
            cb6.HeaderText = "重测次数";


            DataGridViewComboBoxColumn cb7 = new DataGridViewComboBoxColumn();
            cb7.HeaderText = "重测失败处理";
            for (int i = 0; i < retryError.Count(); i++)
            {
                cb7.Items.Add(retryError[i]);
            }


            dv.Columns.Add(cb0);
            dv.Columns.Add(cb1);
            dv.Columns.Add(cb2);
            dv.Columns.Add(cb3);
            dv.Columns.Add(cb4);
            dv.Columns.Add(cb5);
            dv.Columns.Add(cb6);
            dv.Columns.Add(cb7);

        }

        private void InitRangeTable(DataGridView dv)
        {
            ///////////////////////////////////////////////
            dv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dv.RowHeadersVisible = false;

            DataGridViewCheckBoxColumn dcb0 = new DataGridViewCheckBoxColumn();
            dcb0.HeaderText = "启用";
            dcb0.Name = "cb_check";
            dcb0.TrueValue = true;
            dcb0.FalseValue = false;
            //column9.DataPropertyName = "IsScienceNature";
            dcb0.DataPropertyName = "IsChecked";


            DataGridViewTextBoxColumn dcb1 = new DataGridViewTextBoxColumn();
            dcb1.HeaderText = "名称";

            DataGridViewTextBoxColumn dcb2 = new DataGridViewTextBoxColumn();
            dcb2.HeaderText = "min";
            //dcb2.Width = 180;

            DataGridViewTextBoxColumn dcb3 = new DataGridViewTextBoxColumn();
            dcb3.HeaderText = "max";
            //dcb3.Width = 180;

            DataGridViewTextBoxColumn dcb4 = new DataGridViewTextBoxColumn();
            dcb4.HeaderText = "Value/L";
            //dcb4.Width = 180;

            DataGridViewTextBoxColumn dcb5 = new DataGridViewTextBoxColumn();
            dcb5.HeaderText = "Value/L";
            //dcb5.Width = 180;

            dv.Columns.Add(dcb0);
            dv.Columns.Add(dcb1);
            dv.Columns.Add(dcb2);
            dv.Columns.Add(dcb3);
            dv.Columns.Add(dcb4);
            dv.Columns.Add(dcb5);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //tabControl2.SelectedIndex = 4;
            //tabControl2.Enabled = false;

            InitStepTable(dataGridViewLeft);
            InitStepTable(dataGridViewRight);
            InitRangeTable(dataGridViewRange);
            InitRangeTable(dataGridView);
            InitDevice();

            LoadProjectConfig(projectName);
            loadCurveRange();

            if (File.Exists(rangePath))
            {
                ReadCsvToValueTable(rangePath, 1, dataGridView);

                AppendLogMessage("Load Range:" + rangePath);
            }
            else
            {
                MessageBox.Show("判断标准文件不存在");
                textTestResultL.Text = "加载失败";
                textTestResultL.BackColor = Color.Red;
                return;
            }

            //string fName = Directory.GetCurrentDirectory().ToString() + "\\AA.csv";

            //ReadCsvToDataGrid(fName, 1, dataGridView2);
        }

        private void loadCurveRange()
        {
            string fName = curvePath;// Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";

            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            try
            {
                curveJudgeType[0] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge1"));
                curveJudgeType[1] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge2"));
                curveJudgeType[2] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge3"));
                curveJudgeType[3] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge4"));
                curveJudgeType[4] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge5"));
                curveJudgeType[5] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge6"));
                curveJudgeType[6] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge7"));
                curveJudgeType[7] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge8"));
                curveJudgeType[8] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge9"));
                curveJudgeType[9] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge10"));
                curveJudgeType[10] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge11"));
                curveJudgeType[11] = int.Parse(iniFile.IniReadValue(fName, "基准", "CurveJudge12"));

                XData_Min=new double[ CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX1")).Length];
                XData_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX1"));
                YData_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY1")).Length];
                YData_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY1"));
                XData_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX1")).Length];
                XData_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX1"));
                YData_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY1")).Length];
                YData_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY1"));

                //if (curveJudgeType[0] == 0)
                //{
                //    chartCurve.Series[1].Points.DataBindXY(XData_Min, YData_Min);
                //    chartCurve.Series[2].Points.DataBindXY(XData_Max, YData_Max);
                //}

                XData1_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX2")).Length];
                XData1_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX2"));
                YData1_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY2")).Length];
                YData1_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY2"));
                XData1_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX2")).Length];
                XData1_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX2"));
                YData1_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY2")).Length];
                YData1_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY2"));

                XData2_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX3")).Length];
                XData2_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX3"));
                YData2_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY3")).Length];
                YData2_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY3"));
                XData2_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX3")).Length];
                XData2_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX3"));
                YData2_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY3")).Length];
                YData2_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY3"));

                XData3_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX4")).Length];
                XData3_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX4"));
                YData3_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY4")).Length];
                YData3_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY4"));
                XData3_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX4")).Length];
                XData3_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX4"));
                YData3_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY4")).Length];
                YData3_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY4"));


                XData4_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX5")).Length];
                XData4_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX5"));
                YData4_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY5")).Length];
                YData4_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY5"));
                XData4_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX5")).Length];
                XData4_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX5"));
                YData4_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY5")).Length];
                YData4_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY5"));

                XData5_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX6")).Length];
                XData5_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinX6"));
                YData5_Min = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY6")).Length];
                YData5_Min = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMinY6"));
                XData5_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX6")).Length];
                XData5_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxX6"));
                YData5_Max = new double[CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY6")).Length];
                YData5_Max = CsvStringToArray(iniFile.IniReadValue(fName, "基准", "CurveMaxY6"));

            }
            catch
            {

            }
        }

        private double[] CsvStringToArray(string str)
        {
            
            string[] split = str.Substring(0,str.Length-8).Split(',');
            double[] data = new double[split.Length];
            //System.Data.DataRow dr = dt.NewRow();
            for (int i = 0; i < split.Length; i++)
            {
                //if (split[0] != "")
                //{
                    data[i] = double.Parse(split[i]);
                //}
            }
            return data;
         }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath">csv文件路径</param>
        /// <param name="n">表示第n行是字段title,第n+1行是记录开始</param>
        private void ReadCsvToDataGrid(string filePath, int n, DataGridView dt)
        {
            dt.Rows.Clear();
            StreamReader reader = new StreamReader(filePath, System.Text.Encoding.Default, false);
            int i = 0, m = 0;
            reader.Peek();
            
            int rowIndex = 0;

            while (reader.Peek() > 0)
            {
                m = m + 1;
                string str = reader.ReadLine();
                if (m >= n + 1)
                {

                    string[] split = str.Split(',');
                    //System.Data.DataRow dr = dt.NewRow();

                    if (split[0] != "")
                    { 
                        dt.Rows.Add();

                        for (i = 0; i < split.Length; i++)
                        {
                            // dv.Cells[i].Value= dv.Cells[i].ValueType(split[0]);
                            //dr[i] = split[i];
                            split[i] = split[i].Replace("\"","");
                            if (i == 0)
                            {
                                if(split[0].ToUpper() == "TRUE")
                                {
                                    dt.Rows[rowIndex].Cells[0].Value = true;
                                }
                                else
                                {
                                    dt.Rows[rowIndex].Cells[0].Value = false;
                                }
                            }
                            else if(i==1)
                            {
                                try
                                {
                                    dt.Rows[rowIndex].Cells[1].Value = stepNames[Array.IndexOf(stepNames, split[1])];
                                }
                                catch { }
                            }
                            else if (i == 5)
                            {
                                try
                                {
                                    dt.Rows[rowIndex].Cells[5].Value = stepTackleError[Array.IndexOf(stepTackleError, split[5])];
                                }
                                catch { }
                            }
                            else
                            {
                                try
                                {
                                    dt.Rows[rowIndex].Cells[i].Value = split[i];
                                }
                                catch { }
                            }
                        }
                        //dt.Rows.Add(dr);
                        rowIndex++;
                    }
                    // dataGridView2.Rows.Add(dv);
                }
            }
            return ;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath">csv文件路径</param>
        /// <param name="n">表示第n行是字段title,第n+1行是记录开始</param>
        private void ReadCsvToRangeView(string filePath, int n, DataGridView dt)
        {
            dt.Rows.Clear();
            StreamReader reader = new StreamReader(filePath, System.Text.Encoding.Default, false);
            int i = 0, m = 0;
            reader.Peek();

            int rowIndex = 0;

            while (reader.Peek() > 0)
            {
                m = m + 1;
                string str = reader.ReadLine();
                if (m >= n + 1)
                {

                    string[] split = str.Split(',');
                    //System.Data.DataRow dr = dt.NewRow();

                    if (split[0] != "")
                    {
                        dt.Rows.Add();

                        for (i = 0; i < split.Length; i++)
                        {
                            // dv.Cells[i].Value= dv.Cells[i].ValueType(split[0]);
                            //dr[i] = split[i];
                            split[i] = split[i].Replace("\"", "");
                            if (i == 0)
                            {
                                if (split[0].ToUpper() == "TRUE")
                                {
                                    dt.Rows[rowIndex].Cells[0].Value = true;
                                }
                                else
                                {
                                    dt.Rows[rowIndex].Cells[0].Value = false;
                                }
                            }
                            else
                            {
                                try
                                {
                                    dt.Rows[rowIndex].Cells[i].Value = split[i];
                                }
                                catch { }
                            }
                        }
                        //dt.Rows.Add(dr);
                        rowIndex++;
                    }
                    // dataGridView2.Rows.Add(dv);
                }
            }
            reader.Close();
            return;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath">csv文件路径</param>
        /// <param name="n">表示第n行是字段title,第n+1行是记录开始</param>
        private void ReadCsvToValueTable(string filePath, int n, DataGridView dt)
        {
            dt.Rows.Clear();
            StreamReader reader = new StreamReader(filePath, System.Text.Encoding.Default, false);
            int i = 0, m = 0;
            reader.Peek();

            int rowIndex = 0;

            while (reader.Peek() > 0)
            {
                m = m + 1;
                string str = reader.ReadLine();
                if (m >= n + 1)
                {

                    string[] split = str.Split(',');
                    //System.Data.DataRow dr = dt.NewRow();

                    if (split[0] != "")
                    {
                        if (split[0].Replace("\"", "").ToUpper() == "TRUE")//only show valid items
                        {
                            dt.Rows.Add();

                            for (i = 0; i < split.Length; i++)
                            {
                                // dv.Cells[i].Value= dv.Cells[i].ValueType(split[0]);
                                //dr[i] = split[i];
                                split[i] = split[i].Replace("\"", "");
                                if (i == 0)
                                {
                                    if (split[0].ToUpper() == "TRUE")
                                    {
                                        dt.Rows[rowIndex].Cells[0].Value = true;
                                    }
                                    else
                                    {
                                        dt.Rows[rowIndex].Cells[0].Value = false;
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        dt.Rows[rowIndex].Cells[i].Value = split[i];
                                    }
                                    catch { }
                                }
                             }
                            rowIndex++;
                        }
                        
                        //dt.Rows.Add(dr);

                    }
                    // dataGridView2.Rows.Add(dv);
                }
               
            }

            for (i = 0; i < 6; i++)
            {
                if (curveJudgeType[i] > 0)
                {
                    dt.Rows.Add();
                    dt.Rows[rowIndex].Cells[0].Value = true;
                    dt.Rows[rowIndex].Cells[1].Value = "WaveForm"+(i+1).ToString();
                    //dt.Rows[rowIndex].Cells[2].Value = curveLowLimit[i].ToString();
                    //dt.Rows[rowIndex].Cells[3].Value = curveHighLimit[i].ToString();


                    rowIndex++;
                }
            }
            reader.Close();
            return;
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {

            dataGridViewToCSV(stepPathLeft, dataGridViewLeft);
            dataGridViewToCSV(stepPathRight, dataGridViewRight);


        }

        private bool dataGridViewToCSV(string path,DataGridView dataGridView)

        {


            if (dataGridView.Rows.Count == 0)

            {

                MessageBox.Show("没有数据可导出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return false;

            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "CSV files (*.csv)|*.csv";

            saveFileDialog.FilterIndex = 0;

            saveFileDialog.RestoreDirectory = true;

            saveFileDialog.CreatePrompt = true;

            saveFileDialog.FileName = path;

            saveFileDialog.Title = "保存";

            if (true)//(saveFileDialog.ShowDialog() == DialogResult.OK)

            {


                try

                {
                    //StreamWriter stream = new StreamWriter(filePath, System.Text.Encoding.Default, false);
                    Stream stream = saveFileDialog.OpenFile();
                    StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.GetEncoding(-0));

                    string strLine = "";

                    //表头

                    for (int i = 0; i < dataGridView.ColumnCount; i++)

                    {

                        if (i > 0)

                            strLine += ",";

                        strLine += dataGridView.Columns[i].HeaderText;

                    }

                    strLine.Remove(strLine.Length - 1);

                    sw.WriteLine(strLine);

                    strLine = "";

                    //表的内容

                    for (int j = 0; j < dataGridView.Rows.Count; j++)

                    {

                        strLine = "";

                        int colCount = dataGridView.Columns.Count;

                        for (int k = 0; k < colCount; k++)

                        {

                            if (k > 0 && k < colCount)

                                strLine += ",";

                            if (dataGridView.Rows[j].Cells[k].Value == null)

                                strLine += "";

                            else

                            {

                                string cell = dataGridView.Rows[j].Cells[k].Value.ToString().Trim();

                                //防止里面含有特殊符号

                                cell = cell.Replace("\"", "\"\"");

                                cell = "\"" + cell + "\"";

                                strLine += cell;

                            }

                        }

                        sw.WriteLine(strLine);

                    }

                    sw.Close();

                    stream.Close();

                   // MessageBox.Show("数据被导出到：" + saveFileDialog.FileName.ToString(), "导出完毕", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    AppendLogMessage("数据被成功导出到：" + saveFileDialog.FileName.ToString());
                   

                }

                catch (Exception ex)

                {

                    MessageBox.Show(ex.Message, "导出错误", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return false;

                }

            }

            return true;

        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (LeftTestThread.ThreadState ==ThreadState.Running)
            {
                LeftTestThread.Abort();
               
            }

            if (RightTestThread.ThreadState == ThreadState.Running)
            {
                RightTestThread.Abort();

            }

            buttonStart.Enabled = true;
        }

        private void btnDeleteStep_Click(object sender, EventArgs e)
        {
            if (dataGridViewLeft.SelectedRows[0].Index >= 0)
            {
                dataGridViewLeft.Rows.RemoveAt(dataGridViewLeft.SelectedRows[0].Index);
            }
        }

        private void btnMoveUpStep_Click(object sender, EventArgs e)
        {
            int index = dataGridViewLeft.SelectedRows[0].Index;
            if (index > 0)
            {
                DataGridViewRow row = dataGridViewLeft.SelectedRows[0]; 

                dataGridViewLeft.Rows.RemoveAt(index);
                dataGridViewLeft.Rows.Insert(index - 1, row);

                dataGridViewLeft.Rows[index].Selected = false;
                dataGridViewLeft.Rows[index - 1].Selected = true;

                dataGridViewLeft.CurrentCell = dataGridViewLeft.SelectedRows[0].Cells[0];
            }
        }

        private void btnMoveDownStep_Click(object sender, EventArgs e)
        {
            int index = dataGridViewLeft.SelectedRows[0].Index;
            if (index < dataGridViewLeft.RowCount-2)
            {
                DataGridViewRow row = dataGridViewLeft.SelectedRows[0];

                dataGridViewLeft.Rows.RemoveAt(index);
                dataGridViewLeft.Rows.Insert(index +1, row);

                dataGridViewLeft.Rows[index].Selected = false;
                dataGridViewLeft.Rows[index +1].Selected = true;

                dataGridViewLeft.CurrentCell = dataGridViewLeft.SelectedRows[0].Cells[0];
            }
        }

        private void btnInsertStep_Click(object sender, EventArgs e)
        {
            int index = dataGridViewLeft.SelectedRows[0].Index;
            if (index > 0)
            {
                DataGridViewRow row = new DataGridViewRow();
                dataGridViewLeft.Rows.Insert(index, row);
            }
            else
            {
                DataGridViewRow row = new DataGridViewRow();
                dataGridViewLeft.Rows.Insert(0, row);
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            testings = false;

            if (LeftTestThread.ThreadState == ThreadState.Running)
            {
                LeftTestThread.Abort();

            }

            if (RightTestThread.ThreadState == ThreadState.Running)
            {
                RightTestThread.Abort();

            }

            buttonStart.Enabled = true;
            btnStop.Enabled = true;
        }

        private void 项目选择ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
            //ReadCsvToRangeView(rangePath, 1, dataGridViewRange);

            if (textSCStatus.Text == "连接成功")
            {
                GetStepList();
            }
            else
            {
                MessageBox.Show("SC 未连接");
            }
        }

        private void GetStepList() // Open Sequence: Send command to open sequence
        {
            //if (textBoxSeqFilePath.Text != "")
            //{
            //    AppendLogMessage("Opening sequence ...");

            //if (SendCommandAndGetResponse("Sequence.Open('" + textBoxSeqFilePath.Text + "')")) // Send Sequence.Open command and wait for result.
            //{
            //    // Command completed successfully
            //    if (GetReturnDataBoolean()) // Sequence.Open returns boolean data indicating if the sequence was opened.
            //{
                //btnSeqRun.Enabled = true;
                //AppendLogMessage("Sequence Opened. Ready to run!");

                if (SendCommandAndGetResponse("Sequence.GetStepsList")) // Send Sequence.GetStepsList command and wait for result.
                {
                flagUpadateCurveCombo = false;
                // Command completed successfully

                //Column Headers
                //  string[] colHeaders = new string[] { "Step Name", "Step Type", "Input Channel", "Output Channel" };

                // DataTable dataTable = InitializeDataTable(colHeaders.Length);

                // Populate Data Table
                stepsList = json.Value<JArray>("returnData"); // Convert return data to dynamic objects array

                    dataGridViewRange.Rows.Clear();

                    int i = 0;
                    foreach (JObject row in stepsList.Children<JObject>())
                    {
                        //DataRow dataRow = dataTable.NewRow();
                        //dataRow[0] = row.Value<string>("Name");
                        //dataRow[1] = row.Value<string>("Type");
                        //dataRow[2] = FormatChannelNames(row.Value<JArray>("InputChannelNames"));
                        //dataRow[3] = FormatChannelNames(row.Value<JArray>("OutputChannelNames"));
                        //dataTable.Rows.Add(dataRow);

                        dataGridViewRange.Rows.Add();

                        dataGridViewRange.Rows[i].Cells[0].Value = false;
                        dataGridViewRange.Rows[i].Cells[1].Value = row.Value<string>("Name");
                        dataGridViewRange.Rows[i].Cells[2].Value = "";
                        dataGridViewRange.Rows[i].Cells[3].Value = "";
                        dataGridViewRange.Rows[i].Cells[4].Value = "";
                        i++;

                    }
                        // Update DataGridView
                        //  dataGridView.DataSource = dataTable;
                        // UpdateHeaders(colHeaders);
                    //}
                }
                //    else
                //{
                //    AppendLogMessage("Sequence failed to open.");
                //}
            //}

        }

        private void btnSaveRange_Click(object sender, EventArgs e)
        {
            dataGridViewToCSV(rangePath, dataGridViewRange);
            Thread.Sleep(100);

            ReadCsvToValueTable(rangePath, 1, dataGridView);
        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem3_Click_1(object sender, EventArgs e)
        {
            string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";
            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            curveIndex[0] = comboCurveNames.SelectedIndex;
            curveIndex[1] = comboBox6.SelectedIndex;
            curveIndex[2] = comboBox7.SelectedIndex;
            curveIndex[3] = comboBox8.SelectedIndex;
            curveIndex[4] = comboBox9.SelectedIndex;
            curveIndex[5] = comboBox10.SelectedIndex;

            try
            {
                iniFile.IniWriteValue(fName, "基准", "comboCurve1", comboCurveNames.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "comboCurve2", comboBox6.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "comboCurve3", comboBox7.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "comboCurve4", comboBox8.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "comboCurve5", comboBox9.SelectedIndex.ToString());
                iniFile.IniWriteValue(fName, "基准", "comboCurve6", comboBox10.SelectedIndex.ToString());

                MessageBox.Show("保存成功");
            }
            catch
            {
                MessageBox.Show("保存失败");
                return;
            }
        }

        private void btnOpenSeq2_Click(object sender, EventArgs e)
        {
            if (textBoxSeqFilePath.Text != "")
            {
                
                AppendLogMessage("Opening sequence ...");

                if (SendCommandAndGetResponse("Sequence.Open('" + textBoxSeqFilePath.Text + "')")) // Send Sequence.Open command and wait for result.
                {
                    // Command completed successfully
                    if (GetReturnDataBoolean()) // Sequence.Open returns boolean data indicating if the sequence was opened.
                    {
                        buttonStart.Enabled = true;
                        btnStop.Enabled = true;

                        btnOpenSeq2.Enabled = false;
                        btnSeqRun.Enabled = true;
                        AppendLogMessage("Sequence Opened. Ready to run!");


                        flagUpadateCurveCombo = false;

                        if (SendCommandAndGetResponse("Sequence.GetStepsList")) // Send Sequence.GetStepsList command and wait for result.
                        {
                            // Command completed successfully
                            //// Populate Data Table
                            stepsList = json.Value<JArray>("returnData"); // Convert return data to dynamic objects array

                        }
                    }
                    else
                    {
                        AppendLogMessage("Sequence failed to open.");
                    }
                }
                else
                {
                    // Command did not complete successfully
                    if (GetErrorType() == "Timeout") // Check if command timed out
                    {
                        AppendLogMessage("Sequence failed to open. Command timed out!");
                    }
                }
            }
        }

        private bool SaveCurveDataToIni (double[]xdata,double[] ydata)
        {
            string fName = curvePath;
            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            string strLine = "";
            for (int i = 0; i < xdata.Length; i++)
            {
                if (i > 0)
                    strLine += ",";
                strLine += xdata[i].ToString();
            }
            strLine.Remove(strLine.Length - 1);
            iniFile.IniWriteValue(fName, "基准", "CurveValueX", strLine);

            strLine = "";
            for (int i = 0; i < xdata.Length; i++)
            {
                if (i > 0)
                    strLine += ",";
                strLine += xdata[i].ToString();
            }
            strLine.Remove(strLine.Length - 1);
            iniFile.IniWriteValue(fName, "基准", "CurveValueY", strLine);

            return true;
        }
        private void 设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurveRange form = new CurveRange();
            try
            {
                if (curveRightClickIndex == 0)
                {
                    int a = XData.Length;
                }
                else if (curveRightClickIndex == 1)
                {
                    int a = XData1.Length;
                }
                else if (curveRightClickIndex == 2)
                {
                    int a = XData2.Length;
                }
                else if (curveRightClickIndex == 3)
                {
                    int a = XData3.Length;
                }
                else if (curveRightClickIndex == 4)
                {
                    int a = XData4.Length;
                }
                else if (curveRightClickIndex == 5)
                {
                    int a = XData5.Length;
                }


            }
            catch
            {
                MessageBox.Show("请先完成一次测试");
                return;
            }
            string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";
            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            iniFile.IniWriteValue(fName, "基准", "curvePath", curvePath);
            iniFile.IniWriteValue(fName, "基准", "curveSelectedIndex", curveRightClickIndex.ToString());

            if (curveRightClickIndex == 0)
            {
                SaveCurveDataToIni(XData, YData);
            }
            else if (curveRightClickIndex == 1)
            {
                SaveCurveDataToIni(XData1, YData1);
            }
            else if (curveRightClickIndex == 2)
            {
                SaveCurveDataToIni(XData2, YData2);
            }
            else if (curveRightClickIndex == 3)
            {
                SaveCurveDataToIni(XData3, YData3);
            }
            else if (curveRightClickIndex == 4)
            {
                SaveCurveDataToIni(XData4, YData4);
            }
            else if (curveRightClickIndex == 5)
            {
                SaveCurveDataToIni(XData5, YData5);
            }


            if (p == null)
            {
                p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "SetRange.exe";
                p.Start();
            }
            else
            {
                if (p.HasExited) //是否正在运行
                {
                    p.Start();
                }
            }
            p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;



            // form.ShowDialog();



            //if (form.DialogResult == DialogResult.OK)
            //{
            //    string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";
            //    FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            //    fs.Close();

            //    if (whichcontrol_name == "chartCurve")
            //    {
            //        try
            //        {
            //            iniFile.IniWriteValue(fName, "基准", "CurveJudge1", form.judgeType.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveLowLimit1", form.lower.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveHighLimit1", form.upper.ToString());

            //            //  MessageBox.Show("保存成功");
            //        }
            //        catch
            //        {
            //            MessageBox.Show("保存失败");
            //            return;
            //        }

            //        curveRightClickIndex = 0;
            //    }
            //    else if (whichcontrol_name == "chart1")
            //    {
            //        try
            //        {
            //            iniFile.IniWriteValue(fName, "基准", "CurveJudge2", form.judgeType.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveLowLimit2", form.lower.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveHighLimit2", form.upper.ToString());

            //            //  MessageBox.Show("保存成功");
            //        }
            //        catch
            //        {
            //            MessageBox.Show("保存失败");
            //            return;
            //        }
            //        curveRightClickIndex = 1;
            //    }
            //    else if (whichcontrol_name == "chart2")
            //    {
            //        try
            //        {
            //            iniFile.IniWriteValue(fName, "基准", "CurveJudge3", form.judgeType.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveLowLimit3", form.lower.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveHighLimit3", form.upper.ToString());

            //            //  MessageBox.Show("保存成功");
            //        }
            //        catch
            //        {
            //            MessageBox.Show("保存失败");
            //            return;
            //        }
            //        curveRightClickIndex = 2;
            //    }
            //    else if (whichcontrol_name == "chart3")
            //    {
            //        try
            //        {
            //            iniFile.IniWriteValue(fName, "基准", "CurveJudge4", form.judgeType.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveLowLimit4", form.lower.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveHighLimit4", form.upper.ToString());

            //            //  MessageBox.Show("保存成功");
            //        }
            //        catch
            //        {
            //            MessageBox.Show("保存失败");
            //            return;
            //        }
            //        curveRightClickIndex = 3;
            //    }
            //    else if (whichcontrol_name == "chart4")
            //    {
            //        try
            //        {
            //            iniFile.IniWriteValue(fName, "基准", "CurveJudge5", form.judgeType.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveLowLimit5", form.lower.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveHighLimit5", form.upper.ToString());

            //            //  MessageBox.Show("保存成功");
            //        }
            //        catch
            //        {
            //            MessageBox.Show("保存失败");
            //            return;
            //        }
            //        curveRightClickIndex = 4;
            //    }
            //    else if (whichcontrol_name == "chart5")
            //    {
            //        try
            //        {
            //            iniFile.IniWriteValue(fName, "基准", "CurveJudge6", form.judgeType.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveLowLimit6", form.lower.ToString());
            //            iniFile.IniWriteValue(fName, "基准", "CurveHighLimit6", form.upper.ToString());

            //            //  MessageBox.Show("保存成功");
            //        }
            //        catch
            //        {
            //            MessageBox.Show("保存失败");
            //            return;
            //        }
            //        curveRightClickIndex = 5;
            //    }

            //}


        }

        private void contextMenuCurve_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            whichcontrol_name = (sender as ContextMenuStrip).SourceControl.Name;
            if (whichcontrol_name == "chartCurve")
            {
                curveRightClickIndex = 0;
            }
            else if (whichcontrol_name == "chart1")
            {
                curveRightClickIndex = 1;
            }
            else if (whichcontrol_name == "chart2")
            {
                curveRightClickIndex = 2;
            }
            else if (whichcontrol_name == "chart3")
            {
                curveRightClickIndex = 3;
            }
            else if (whichcontrol_name == "chart4")
            {

                curveRightClickIndex = 4;
            }
            else if (whichcontrol_name == "chart5")
            {
                curveRightClickIndex = 5;
            }
            // MessageBox.Show(whichcontrol_name);
        }

        private void 设置Right范围ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurveRange form = new CurveRange();
            try
            {
                if (curveRightClickIndex == 0)
                {
                    int a = RXData.Length;
                }
                else if (curveRightClickIndex == 1)
                {
                    int a = RXData1.Length;
                }
                else if (curveRightClickIndex == 2)
                {
                    int a = RXData2.Length;
                }
                else if (curveRightClickIndex == 3)
                {
                    int a = RXData3.Length;
                }
                else if (curveRightClickIndex == 4)
                {
                    int a = RXData4.Length;
                }
                else if (curveRightClickIndex == 5)
                {
                    int a = RXData5.Length;
                }


            }
            catch
            {
                MessageBox.Show("请先完成一次测试");
                return;
            }
            string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";
            FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Close();

            iniFile.IniWriteValue(fName, "基准", "curvePath", curvePath);
            iniFile.IniWriteValue(fName, "基准", "curveSelectedIndex", (curveRightClickIndex+6).ToString());

            if (curveRightClickIndex == 0)
            {
                SaveCurveDataToIni(RXData, RYData);
            }
            else if (curveRightClickIndex == 1)
            {
                SaveCurveDataToIni(RXData1, RYData1);
            }
            else if (curveRightClickIndex == 2)
            {
                SaveCurveDataToIni(RXData2, RYData2);
            }
            else if (curveRightClickIndex == 3)
            {
                SaveCurveDataToIni(RXData3, RYData3);
            }
            else if (curveRightClickIndex == 4)
            {
                SaveCurveDataToIni(RXData4, RYData4);
            }
            else if (curveRightClickIndex == 5)
            {
                SaveCurveDataToIni(RXData5, RYData5);
            }


            if (p == null)
            {
                p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "SetRange.exe";
                p.Start();
            }
            else
            {
                if (p.HasExited) //是否正在运行
                {
                    p.Start();
                }
            }
            p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;        
        }
    }
}