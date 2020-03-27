using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TrinityMIB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Thread t;
        LabelManager2.Application lbl;       
        string docPath =Directory.GetCurrentDirectory()+ "\\SN.Lab";
        StringBuilder sb = new StringBuilder();
        
      
        SqlConnection conn  ;
        DBConnection DB ;
    

        DateTime startTime;
        string pImei = string.Empty;
        string pSN = string.Empty;
        string pPn = string.Empty;
        string pEMID = string.Empty;
        string BT_MAC = string.Empty;
        string WIFI_MAC = string.Empty;
        string CSN = string.Empty;
        string pedesMAC = string.Empty;
        string CN = string.Empty;
        string errmessage = string.Empty;

        byte[] bytepImei = new byte[1024];
        byte[] bytepSN = new byte[1024];
        byte[] bytepPn = new byte[1024];
        byte[] bytepEMID = new byte[1024];
        byte[] byteBT_MAC = new byte[1024];
        byte[] byteWIFI_MAC = new byte[1024];
        byte[] byteCSN = new byte[1024];
        byte[] bytepedesMAC = new byte[1024];
        byte[] byteCN = new byte[1024];
        byte[] byteerrmessage = new byte[1024];
       


        string SFC_Port= ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Port Settings", "SFCPort", "");
        int SFC_BaudRate=Convert.ToInt32(ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Port Settings", "SFCBaudRate", ""));

        string strSql = string.Empty;
        bool isExistData = false;
        string receive = string.Empty;
        string SFCFuction = string.Empty;
        private static StringBuilder sFileSplice = new StringBuilder();    //创建拼接语句对象
        string stationName = string.Empty;
        SerialPort mySerialPort = null;
        string computerName = Dns.GetHostName();

        //SqlConnection getCon;

        private static bool g_bThreadAlive = false;
        private static string RESULT_SUCCESS = "1PASS", RESULT_FAILED = "1FAILED";

        [DllImport("namepipe.dll", CharSet = CharSet.Ansi)]
        public static extern int GetInfo(int dev_type, byte[] pOutInfo);

   
        [DllImport("namepipe.dll", CharSet = CharSet.Auto)]
        public static extern int SendResult(int dev_type, byte[] msg);

     
        [DllImport("namepipe.dll", CharSet = CharSet.Auto)]
        public static extern bool GetX990(byte[] pImei, byte[] pSN, byte[] pPn, byte[] pEMID, byte[] BT_MAC, byte[] WIFI_MAC, byte[] CSN, byte[] pedesMAC, byte[] errmessage);


        private void Form1_Load(object sender, EventArgs e)
        {
            panelSFC.Visible = false;
            //mySerialPort.PortName= ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Port Settings", "SFCPort", "");
            //mySerialPort.BaudRate=Convert.ToInt32(ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Port Settings", "SFCBaudRate", ""));
            //mySerialPort.Open();
            if (mySerialPort==null)
            {
                mySerialPort = new SerialPort(SFC_Port, SFC_BaudRate, Parity.None, 8, StopBits.One); //设置串口
            }
            if (!mySerialPort.IsOpen)
            {
                mySerialPort.Open();
                lbl_SFC.ForeColor = Color.DarkGreen;
            }
            if (normalPrintToolStripMenuItem.Text.ToUpper().Contains("NORMAL"))
            {
                lblPrintMode.ForeColor = Color.DarkGreen;
            }
            
            this.Text = string.Format("Trnity MIB V{0}  Build Date {1}",
            Application.ProductVersion.ToString(), System.IO.File.GetLastWriteTime(this.GetType().Assembly.Location).ToString("yyyy/MM/dd HH:mm:ss"));    //增加 版本和显示时间

            DB = new DBConnection();
            conn = DB.CreateConection();//创建conn时，便已打开conn
            lbl_DBConn.ForeColor = Color.DarkGreen;//数据库字体变绿
            normalPrintToolStripMenuItem.Text= ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Print Settings", "PrintMode", "");

            ///   ---------------------------------------------------------------------------------------------------------------------




            timer1.Interval = 200;
        

            #region UI界面

            stationName = ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Station Settings", "StationName", "");
            SFCFuction = ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "Port Settings", "SFCFuction", "");

            lbl_StationName.Text = stationName;
            lbl_PC_Name.Text = computerName;
            lbl_ModelName.Text =txt_ModeType +"/ "+ txt_PNNO.Text + "/ " + txt_VerNO.Text;

            comBox_Com.Items.AddRange(SerialPort.GetPortNames());
            comBox_Com.SelectedItem = SFC_Port;//

            strip_ServerIP.Text = "Server IP: " + DB.Server;
            strip_LocalIP.Text = "Local IP: " + Dns.GetHostAddresses(computerName);

            #endregion

        }



  
   

        private void txt_CN_KeyPress(object sender, KeyPressEventArgs e)
        {
            lbl_Msg.Text = "";
            SFCFuction = yesToolStripMenuItem.Text.Trim();
            
            CN=txt_CN.Text.Trim().ToUpper();
            
            if (e.KeyChar == 13 && CN.Length == 12)//回车键
            {
                //查询数据库，显示MIB表里的内容
                string strQuery = "SELECT [SNCode],[CustomerPN],[BomVer],[WIFIMAC],[BTMAC],[LANMAC] FROM [Verifone].[dbo].[t_MIB_UNIT_Info] WHERE CNCode='" + CN + "'";
                string strMsg = string.Empty;
                DataTable dt = DB.GetData(strQuery, out strMsg);
                if (dt.Rows.Count != 0)//数据库表中已有数据
                {
                    isExistData = true;
                    txt_DB_pSN.Text = dt.Rows[0]["SNCode"].ToString();
                    lbl_SN.Text = txt_DB_pSN.Text;
                    txt_DB_PN.Text = dt.Rows[0]["CustomerPN"].ToString();
                    txt_DB_WIFI_MAC.Text = dt.Rows[0]["WIFIMAC"].ToString();
                    txt_DB_BT_MAC.Text = dt.Rows[0]["BTMAC"].ToString(); ;
                    txt_DB_LANMAC.Text = dt.Rows[0]["LANMAC"].ToString();

                    if (txt_DB_pSN.Text.Length != 0)
                    {
                        gb_Result.BackColor = Color.Goldenrod;//已经分配过SN，result背景变为土黄色                               
                    }
                }
                if (normalPrintToolStripMenuItem.Text.ToUpper().Trim().Contains("REPRINT"))//补打,   不过sfc
                {
                    //若是补打，数据库里必然存在板子信息
                    string pSN = dt.Rows[0]["SNCode"].ToString();
                    Print(pSN);
                    ClearData();
                }
                else
                {
                    timer1.Enabled = true;
                    if (SFCFuction.ToUpper().Trim() == "YES")
                    {
                        //发送sfc1    
                        receive = SFCSplicing(stationName, CN, SFCFuction);
                    }                  
                }

             }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            bool isGet = GetX990(bytepImei, bytepSN, bytepPn, bytepEMID, byteBT_MAC, byteWIFI_MAC, byteCSN, bytepedesMAC, byteerrmessage);
         
            if (isGet)
            {
                //获取数据        

                pImei = System.Text.Encoding.Default.GetString(bytepImei).Replace("\0", "").ToUpper().Trim();
                pSN = System.Text.Encoding.Default.GetString(bytepSN).Replace("\0", "").ToUpper().Trim();
                pPn = System.Text.Encoding.Default.GetString(bytepPn).Replace("\0", "").ToUpper().Trim();
                pEMID = System.Text.Encoding.Default.GetString(bytepEMID).Replace("\0", "").ToUpper().Trim();
                BT_MAC = System.Text.Encoding.Default.GetString(byteBT_MAC).Replace("\0", "").ToUpper().Trim();
                WIFI_MAC = System.Text.Encoding.Default.GetString(byteWIFI_MAC).Replace("\0", "").ToUpper().Trim();
                CSN = System.Text.Encoding.Default.GetString(byteCSN).Replace("\0", "").ToUpper().Trim();
                pedesMAC = System.Text.Encoding.Default.GetString(bytepedesMAC).Replace("\0", "").ToUpper().Trim();
                errmessage = System.Text.Encoding.Default.GetString(byteerrmessage).Replace("\0", "").ToUpper().Trim();
                          

                // 更新UI
                txt_PCB_pSN.Text = pSN;
                txt_PCB_PN.Text = pPn;
                txt_PCB_BT_MAC.Text = BT_MAC;
                txt_PCB_WIFI_MAC.Text = WIFI_MAC;
                txt_PCB_LANMAC.Text = pedesMAC;
                // MessageBox.Show("ok");

                //数据库里的SN与分配给板子的SN不符合，则报错
                if (txt_DB_pSN.Text!= txt_PCB_pSN.Text)
                {
                    gb_Result.BackColor = Color.Red;//红色提示
                }
                 else if (SFCFuction.ToUpper().Trim() == "YES")
                {
                    if (receive.Contains("#OK,UNIT STATUS IS VALID"))
                    {

                        if (true)//结果无错
                        {
                            //打印sn
                            Print(pSN);
                            //Log
                            // createLog(stationName, "-1", pSN, DateTime.Now.ToString(), "Trinity");
                            //发送SFC2 PASS
                            string BomVer = string.Empty;
                            string PTID = string.Empty;
                            string HWID = string.Empty;
                            string NDS = string.Empty;

                            pSN = txt_PCB_pSN.Text.Trim().ToUpper();
                            pPn = txt_PCB_PN.Text.Trim().ToUpper();
                            WIFI_MAC = txt_PCB_WIFI_MAC.Text.Trim().ToUpper();
                            BT_MAC = txt_PCB_BT_MAC.Text.Trim().ToUpper();
                            pedesMAC = txt_PCB_LANMAC.Text.Trim().ToUpper();

                            string dataArgs = "'" + pSN + "'," + "'" + pPn + "'," + "'" + WIFI_MAC + "'," + "'" + BT_MAC + "'," + "'" + pedesMAC + "'," + "'" + CN + "'," + "'" + DateTime.Now.ToString() + "'";

                            string strInsert = string.Format("INSERT INTO [dbo].[t_MIB_UNIT_Info] ([SNCode],[CustomerPN],[WIFIMAC],[BTMAC],[LANMAC],[CNCode],[TestTime]) VALUES ({0})"
                                                                                                           , dataArgs);
                            string strUpdate = "UPDATE [dbo].[t_MIB_UNIT_Info] SET [SNCode]='" + pSN + "',[CustomerPN]='" + pPn + "',[WIFIMAC]='" + WIFI_MAC + "',[BTMAC]='" + BT_MAC + "',[LANMAC]='" + pedesMAC + "' WHERE [CNCode]='" + CN + "'";

                            //如果数据库中有数据，则为更新，否则为插入
                            if (isExistData)
                            {
                                strSql = strUpdate;
                            }
                            else
                            {
                                strSql = strInsert;
                            }

                            SqlCommand cmd = new SqlCommand(strSql, conn);
                            int row = cmd.ExecuteNonQuery();
                            if (row != 0)
                            {
                                lbl_Msg.Text = "插入数据库成功";
                                receive = SFCSplicing(stationName, CN, pSN, computerName, "", "", pedesMAC, WIFI_MAC, BT_MAC, "PASS", SFCFuction);//发送过站信息
                                if (receive.Contains("#OK,UNIT PASS!") || SFCFuction.ToUpper().Trim() == "No")                                                                                                                                                                                                       //判断过战信息
                                {

                                    ClearData();//文本框内容清除
                                    this.BackColor = Color.Green;

                                }
                                else
                                {
                                    ClearData();//文本框内容清除
                                    this.BackColor = Color.Red;
                                    return;
                                }

                            }
                            else
                            {
                                lbl_Msg.Text = "未插入数据库";
                            }


                        }
                        else  //errmessage.Trim() != ""有问题
                        {
                            SFCSplicing(stationName, CN, pSN, computerName, "", "", pedesMAC, WIFI_MAC, BT_MAC, "PASS", SFCFuction);//发送过站信息
                        }

                    }
                    else  //SFC1回复的信息错误
                    {
                        this.BackColor = Color.Red;
                        ClearData();//文本框内容清除
                    }
                }
                else if (SFCFuction.ToUpper().Trim() == "NO")//关SFC测试
                {
                    Print(pSN);
                    //Log
                    // createLog(stationName, "-1", pSN, DateTime.Now.ToString(), "Trinity");
                    //发送SFC2 PASS
                    string BomVer = string.Empty;
                    string PTID = string.Empty;
                    string HWID = string.Empty;
                    string NDS = string.Empty;

                    pSN = txt_PCB_pSN.Text.Trim().ToUpper();
                    pPn = txt_PCB_PN.Text.Trim().ToUpper();
                    WIFI_MAC = txt_PCB_WIFI_MAC.Text.Trim().ToUpper();
                    BT_MAC = txt_PCB_BT_MAC.Text.Trim().ToUpper();
                    pedesMAC = txt_PCB_LANMAC.Text.Trim().ToUpper();

                    string dataArgs = "'" + pSN + "'," + "'" + pPn + "'," + "'" + WIFI_MAC + "'," + "'" + BT_MAC + "'," + "'" + pedesMAC + "'," + "'" + CN + "'," + "'" + DateTime.Now.ToString() + "'";

                    string strInsert = string.Format("INSERT INTO [dbo].[t_MIB_UNIT_Info] ([SNCode],[CustomerPN],[WIFIMAC],[BTMAC],[LANMAC],[CNCode],[TestTime]) VALUES ({0})"
                                                                                                   , dataArgs);
                    string strUpdate = "UPDATE [dbo].[t_MIB_UNIT_Info] SET [SNCode]='" + pSN + "',[CustomerPN]='" + pPn + "',[WIFIMAC]='" + WIFI_MAC + "',[BTMAC]='" + BT_MAC + "',[LANMAC]='" + pedesMAC + "' WHERE [CNCode]='" + CN + "'";

                    //如果数据库中有数据，则为更新，否则为插入
                    if (isExistData)
                    {
                        strSql = strUpdate;
                    }
                    else
                    {
                        strSql = strInsert;
                    }

                    SqlCommand cmd = new SqlCommand(strSql, conn);
                    int row = cmd.ExecuteNonQuery();
                    //***********
                }
               
            }
        }

        public void ClearData()
        {
            txt_CN.Focus();
            txt_CN.SelectAll();

            txt_DB_pSN.Text = string.Empty;
            txt_DB_PN.Text = string.Empty;
            txt_DB_WIFI_MAC.Text = string.Empty;
            txt_DB_BT_MAC.Text = string.Empty;
            txt_DB_LANMAC.Text = string.Empty;

            txt_PCB_pSN.Text = string.Empty;
            txt_PCB_PN.Text = string.Empty;
            txt_PCB_BT_MAC.Text = string.Empty;
            txt_PCB_WIFI_MAC.Text = string.Empty;
            txt_PCB_LANMAC.Text = string.Empty;
            gb_Result.BackColor = Color.White;
            lbl_SN.Text = string.Empty;
            timer1.Enabled = false;

        }






        public string SFCSplicing(string stationName, string CN, string mark)
        {
            if (mark.ToUpper() == "YES")
            {
                sFileSplice.Remove(0, sFileSplice.Length);                                                                                                                                                                                                                              //初始化sfc语句

                sFileSplice.Append("1>>").Append(CN).Append(",G6011184,").Append(stationName).Append(",#");

                receive = sFileSplice.ToString().Trim() + "\r\n";
                //发送和接收语句
                receive = System.Text.Encoding.Default.GetString(GetPortInfo(System.Text.Encoding.Default.GetBytes(receive)));
                return receive;
            }
            else
            {
                return "";
            }
        }

        public string SFCSplicing(string stationName, string CN,string pSN, string computerName, string PDID, string IMEI, string LANMAC, string WIFIMAC, string BTMAC, string ErrorCode, string mark)
        {
            if (mark.ToUpper() == "YES")
            {
                sFileSplice.Remove(0, sFileSplice.Length);
                sFileSplice.Append("2>>").Append(CN).Append(","+pSN+",").Append(PDID + ",").Append(IMEI).Append(",,,,,").Append("Verifone,")
                .Append(stationName + ",").Append(computerName).Append(",,,").Append(LANMAC + ",").Append(WIFIMAC + ",").Append(BTMAC).
                Append(",,,,,,,,,").Append("g6011184").Append(",,,,,,,,,,,,#").Append(ErrorCode);
                receive = sFileSplice.ToString().Trim() + "\r\n";
                //发送和接收语句
                receive = System.Text.Encoding.Default.GetString(GetPortInfo(System.Text.Encoding.Default.GetBytes(receive)));
                return receive;
            }
            else
            {
                return "";
            }
        }

        private void Print(string pSN)
        {//print
           // string SN = pSN;
            lbl = new LabelManager2.Application();
            lbl.Documents.Open(docPath, true);
            LabelManager2.Document document = lbl.ActiveDocument;

            document.Variables.FreeVariables.Item("SN").Value = pSN;
         
            document.PrintDocument(1);
            document.FormFeed();
            lbl.Documents.CloseAll();
            document = null;
            lbl.Quit();
            //createLog("BOX_LABEL", CN, pSN, DateTime.Now.ToString("yyyy-MM-dd"), "Trinity");
            stopProcess("lppa");
        }
        private void stopProcess(string process)
        {
            foreach (Process item in System.Diagnostics.Process.GetProcessesByName(process.Trim()))
            {
                try
                {
                    item.Kill();
                    item.WaitForExit();
                }
                catch
                {
                    MessageBox.Show("can not kill" + process + "process");
                }
            }
        }

        private void createLog(string STATION, string CN, string SN, string Time, string Model)
        {
            sb = new StringBuilder();
            sb.Append("D:\\BOX_LABEL_LOG\\" + "L5_BOX_LABEL_LOG_" + DateTime.Now.ToString("yyyy-MM-dd").Replace("-", "") + ".txt");
            
            if (File.Exists(sb.ToString()))
            {   //log file is exist
                FileStream fileStream = new FileStream(sb.ToString(), FileMode.Append);
                StreamWriter streamWriter = new StreamWriter(fileStream);
                streamWriter.WriteLine(STATION + "," + CN + "," + SN + "," + Time + "," + Model);
                streamWriter.Close();
                streamWriter.Dispose();
                fileStream.Close();
                fileStream.Dispose();
            }
            else
            {   //log file is not exist
                FileStream fileStream = new FileStream(sb.ToString(), FileMode.Append);
                StreamWriter streamWriter = new StreamWriter(fileStream);
                streamWriter.WriteLine("STATION,CN,SN,Time,Model");
                streamWriter.WriteLine(STATION + "," + CN + "," + SN + "," + Time + "," + Model);
                streamWriter.Close();
                streamWriter.Dispose();
                fileStream.Close();
                fileStream.Dispose();
            }
        }

        private byte[] GetPortInfo(byte[] sendDate)
        {   
            int cycleIndex = 0;
            if (mySerialPort == null)
            {
                mySerialPort = new SerialPort(SFC_Port, SFC_BaudRate, Parity.None, 8, StopBits.One); //设置串口
                //设置串口信息的写入长度
                mySerialPort.ReadBufferSize = 1024;                                                                       //设置串口信息的读取长度  
                mySerialPort.WriteBufferSize = 1024;
                //串口打开标志变绿
                lbl_SFC.ForeColor = Color.DarkGreen;
                
            }

            if (!mySerialPort.IsOpen)                                                                                                  //判断串口是否打开
            {
                mySerialPort.Open();
            }

            mySerialPort.Write(sendDate, 0, sendDate.Length);                                          //向串口发送信息

            //接收数据
            while (mySerialPort.BytesToRead == 0)                                                                  //判断是否接收到信息
            {
                cycleIndex++;
                Thread.Sleep(1);
                if (cycleIndex == 10000)
                {
                    break;
                }
            }
            Thread.Sleep(50);                                                                                              //50毫秒内接收完数据
            byte[] recData = new byte[mySerialPort.BytesToRead];                      //接收数组的长度
            mySerialPort.Read(recData, 0, recData.Length);                                    //读取回复信息
            mySerialPort.Close();                                                                                       //关闭串口
            mySerialPort.Dispose();
            return recData;                                                                                                  //返回数据
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //try
            //{
            //    t.Abort();
            //    t.Join();
            //}
            //catch (Exception e33)
            //{                
            //}
        }


        private void yesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (yesToolStripMenuItem.Text.ToUpper().Trim()=="YES")
            {
                panelSFC.Visible = true;
            }
            else if (yesToolStripMenuItem.Text.ToUpper().Trim() == "NO")
            {
                yesToolStripMenuItem.Text = "YES";//open sfc
                lbl_SFC.ForeColor = Color.DarkGreen;
            }
        }

        private void normalPrintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (normalPrintToolStripMenuItem.Text.ToUpper().Trim() == "NORMAL")
            {
                panelPrint.Visible = true;
            }
            else if (normalPrintToolStripMenuItem.Text.ToUpper().Trim() == "REPRINT")
            {
                normalPrintToolStripMenuItem.Text = "NORMAL";//open  normal print
                lblPrintMode.Text = normalPrintToolStripMenuItem.Text + " " + "Mode";
                lblPrintMode.ForeColor = Color.DarkGreen;
            }
            panelSFC.Visible = true;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txt_SFC_Name.Text.Trim().ToUpper()=="G6011184"&&txt_SFC_Pwd.Text.Trim().ToUpper()=="VERIFONE")
            {
                yesToolStripMenuItem.Text = "NO";//close sfc
                lbl_SFC.Text = "SFC Port Close";
                lbl_SFC.ForeColor = Color.Red;
                panelSFC.Visible = false;
            }         
        }
        private void btn_LoginPrint_Click(object sender, EventArgs e)
        {
            if (txt_SFC_Name.Text.Trim().ToUpper() == "G6011184" && txt_SFC_Pwd.Text.Trim().ToUpper() == "VERIFONE")
            {
                normalPrintToolStripMenuItem.Text = "RePrint";//close sfc
                lblPrintMode.Text = normalPrintToolStripMenuItem.Text + " " + "Mode";
                lblPrintMode.ForeColor = Color.Red;
                panelPrint.Visible = false;
            }
        }      

       

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            panelSFC.Visible = false;
            panelPrint.Visible = false;
        }         

        
        
    }
}
