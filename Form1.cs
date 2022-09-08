using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Management;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Win32;

namespace GP2K
{
    public partial class Form1 : Form
    {
        private string com_number; 
        private int com_data_len;
        private bool com_port_open;
        private bool first_open;
        private byte[] com_buff;
        public Form1()
        {
            DialogResult dialogResult = MessageBox.Show("Please use for legal purposes and this procedure will not be liable for any damages.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.No)
                Close();

            InitializeComponent();
            com_data_len = 0;
            foreach (COMPortInfo comPort in COMPortInfo.GetCOMPortsInfo())
            {
                comboBox1.Items.Add(string.Format("{0}-{1}", comPort.Name, comPort.Description));
            }
            comboBox1.SelectedIndex = 0;
            
        }
        internal class ProcessConnection
        {
            public static ConnectionOptions ProcessConnectionOptions()
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.Authentication = AuthenticationLevel.Default;
                options.EnablePrivileges = true;
                return options;
            }
            public static ManagementScope ConnectionScope(string machineName, ConnectionOptions options, string path)
            {
                ManagementScope connectScope = new ManagementScope();
                connectScope.Path = new ManagementPath(@"\\" + machineName + path);
                connectScope.Options = options;
                connectScope.Connect();
                return connectScope;
            }
        }
        public class COMPortInfo
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public COMPortInfo() { }

            public static List<COMPortInfo> GetCOMPortsInfo()
            {
                List<COMPortInfo> comPortInfoList = new List<COMPortInfo>();
                ConnectionOptions options = ProcessConnection.ProcessConnectionOptions();
                ManagementScope connectionScope = ProcessConnection.ConnectionScope(Environment.MachineName, options, @"\root\CIMV2");
                ObjectQuery objectQuery = new ObjectQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0");
                ManagementObjectSearcher comPortSearcher = new ManagementObjectSearcher(connectionScope, objectQuery);
                using (comPortSearcher)
                {
                    string caption = null;
                    foreach (ManagementObject obj in comPortSearcher.Get())
                    {
                        if (obj != null)
                        {
                            object captionObj = obj["Caption"];
                            if (captionObj != null)
                            {
                                caption = captionObj.ToString();
                                if (caption.Contains("(COM"))
                                {
                                    COMPortInfo comPortInfo = new COMPortInfo();
                                    comPortInfo.Name = caption.Substring(caption.LastIndexOf("(COM")).Replace("(", string.Empty).Replace(")", string.Empty);
                                    comPortInfo.Description = caption;
                                    comPortInfoList.Add(comPortInfo);
                                }
                            }
                        }
                    }
                }
                return comPortInfoList;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            serialPort1.PortName = com_number;
            serialPort1.BaudRate = 9600;
            serialPort1.Parity = Parity.None;
            serialPort1.DataBits = 8;
            serialPort1.WriteBufferSize = 16;
            serialPort1.ReadBufferSize = 512;
            serialPort1.StopBits = StopBits.One;
            serialPort1.WriteTimeout = 500;
            serialPort1.ReadTimeout = 1000;
            serialPort1.DtrEnable = true;
            serialPort1.RtsEnable = true;
            textBox1.Text = String.Empty;
            textBox2.Text = String.Empty;
            button1.Enabled = false;
            try
            {
                serialPort1.Open();
                com_port_open = true;
                Thread RSDataChk = new Thread(Datachk);
                RSDataChk.IsBackground = true;
                RSDataChk.Start();
            }
            catch
            {
                MessageBox.Show("Cant open " + com_number + "\nThe port is already in use", com_number + " Open error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = true;
            }
        }
        private void Datachk()
        {
            int num = 0;
            int num1 = 0;
            com_data_len = 0;
            bool get_pw = false;
            byte[] com_buffchk1 = { 0x10, 0x3C, 0x43, 0x0D, 0x1A };
            num1 = 1;
            while (true)
            {
                AppendTextBox("Connecting..." + num1);
                serialPort1.Write(new byte[] { 0x10, 0x3E, 0x43, 0x0D }, 0, 4);
                while (com_data_len <= 3)
                {
                    Thread.Sleep(1);
                    num++;
                    if (num > 1000)
                        break;
                }
                num = 0;
                //確定資料有取得
                if (com_data_len > 3)
                {
                    //比對取得資料是否一致
                    if (BitConverter.ToString(com_buffchk1) == BitConverter.ToString(com_buff))
                    {
                        com_data_len = 0;
                        //送出取得LS區命令
                        serialPort1.Write(new byte[] { 0x10, 0x3E, 0x4C, 0x53, 0x30, 0x30, 0x30, 0x30, 0x38, 0x0D }, 0, 10);
                        while (com_data_len < 32)
                        {
                            Thread.Sleep(1);
                            num++;
                            if (num > 5000)
                            {
                                AppendTextBox("Decryption failure");
                                AppendTextBox("-------------------------");
                                AppendTextBox("Please use the GP2000 software to connect and read it first, and then use this program to reconnect and decrypt it after the password window pops up. If the connection still fails, this model is not supported.");
                                AppendTextBox("-------------------------");
                                break;
                            }
                        }
                        if (num < 5000)
                        {
                            com_buffchk1 = com_buff;
                            // Thread.Sleep(1000);
                            int data_len = com_data_len;
                            com_data_len = 0;
                            AppendTextBox("Decryption complete");
                            byte[] com_buffchk2 = new byte[24];
                            int j = 0, k = 0;
                            for (int i = 0; i < data_len; i++)
                            {
                                if (com_buffchk1[i] == 0x0A)
                                    break;
                                k++;
                            }
                            k = k + 2;
                            for (int i = 0; i < 24; i++)
                            {
                                if (com_buffchk1[i + k] != 0x0A)
                                {
                                    com_buffchk2[i] = (byte)(com_buffchk1[i + k] - 0xA);
                                    j++;
                                }
                                else
                                {
                                    string str = System.Text.Encoding.ASCII.GetString(com_buffchk2);
                                    AppendTextBox("Passwords: " + str + "\n");
                                    AppendTextBox2(str);
                                    get_pw = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                com_data_len = 0;
                num1++;
                if (num1 > 10 || get_pw)
                {
                    break;
                }
            }
            if (num1 > 10)
            {
                AppendTextBox("Connection failed");
                AppendTextBox("-------------------------");
                AppendTextBox("Please use the GP2000 software to connect and read it first, and then use this program to reconnect and decrypt it after the password window pops up. If the connection still fails, this model is not supported.");
            }
            if (com_port_open)
            {
                serialPort1.Close();
                com_port_open = false;
            }
            AppendButton1(true);
        }
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string b = (string)comboBox1.SelectedItem;
            string[] sArray = b.Split('-');
            com_number = sArray[0];
            textBox2.Text = "Use " + com_number + Environment.NewLine;
            if (!first_open)
            {
                textBox2.Text += "-------------------------" + Environment.NewLine +
                                 "Please use for legal purposes and this program will not be held liable for any damages." + Environment.NewLine + "Program Development By HaoY. Wei" + Environment.NewLine +
                                 "https://github.com/kikiqqp" + Environment.NewLine + "2022/04/14";
                first_open = true;
            }
            int maxSize = 0;
            System.Drawing.Graphics g = CreateGraphics();
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                SizeF size = g.MeasureString(comboBox1.Text, comboBox1.Font);
                if (maxSize < (int)size.Width)
                    maxSize = (int)size.Width;
            }
            comboBox1.DropDownWidth = comboBox1.Width;
            if (comboBox1.DropDownWidth < maxSize)
                comboBox1.DropDownWidth = maxSize;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.ScrollBars = ScrollBars.Vertical;
            textBox2.SelectionStart = textBox2.Text.Length;
            textBox2.ScrollToCaret();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            foreach (COMPortInfo comPort in COMPortInfo.GetCOMPortsInfo())
            {
                comboBox1.Items.Add(string.Format("{0}-{1}", comPort.Name, comPort.Description));
            }
            comboBox1.SelectedIndex = 0;
        }
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                Thread.Sleep(350);
                com_data_len = serialPort1.BytesToRead;
                if (com_data_len != 0)
                {
                    com_buff = new byte[com_data_len];
                    serialPort1.Read(com_buff, 0, com_data_len);
                }
            }
            catch
            {
                return;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (com_port_open)
            {
                serialPort1.Close();
                com_port_open = false;
            }
        }
        public void AppendTextBox(string value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendTextBox), new object[] { value });
                return;
            }
            textBox2.Text += value + Environment.NewLine;
        }
        public void AppendTextBox2(string value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendTextBox2), new object[] { value });
                return;
            }
            textBox1.Text = value;
        }
        public void AppendButton1(bool value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<bool>(AppendButton1), new object[] { value });
                return;
            }
            button1.Enabled = value;
        }
    }

}
