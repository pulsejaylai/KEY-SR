using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.Diagnostics;




namespace SR1000
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string configtxt,pmodel, ticonfig;
        SerialPort sp;
        FileStream fs;
        FileStream savedata;
        FileStream sndata;
        StreamWriter sw2;
        StreamWriter sw3;
        StreamReader sr2;
        StreamWriter checksn;
        Thread thread1, thread2, thread3;
        bool comtrue, teststatue, filenew;
        public bool debug=false;
        SqlConnection con = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        DataTable dt;
         string filepath = "d:\\TelasSNCheck\\M1829";
        string wo, tfilename, tfilename2;
        int x,testspec,fourcount,pan,dpan=1;
        string FTPAddress = "ftp://172.17.11.65"; //ftp服务器地址
        string FTPUsername = "LarryYang";   //用户名
        string FTPPwd = "Welcome@123";        //密码
        int ftpcheck = 0;
        private void UpFile(string localfilepath, string ftpfilepath)
        {
            string LocalPath = localfilepath; //待上传文件
            FileInfo f = new FileInfo(LocalPath);
            string FileName = f.Name;
            string ftpRemotePath = ftpfilepath;
            string FTPPath = FTPAddress + ftpRemotePath + FileName;
            FtpWebRequest reqFtp = (FtpWebRequest)FtpWebRequest.Create(new Uri(FTPPath));
            reqFtp.UseBinary = true;
            reqFtp.Credentials = new NetworkCredential(FTPUsername, FTPPwd); //设置通信凭据
            reqFtp.KeepAlive = false; //请求完成后关闭ftp连接
            reqFtp.Method = WebRequestMethods.Ftp.UploadFile;
            reqFtp.ContentLength = f.Length;
            int buffLength = 2048;
            byte[] buff = new byte[buffLength];
            int contentLen;

            FileStream fs = f.OpenRead();
            try
            {
                Stream strm = reqFtp.GetRequestStream();
                contentLen = fs.Read(buff, 0, buffLength);
                while (contentLen != 0)
                {
                    strm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff, 0, buffLength);
                }
                strm.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                ftpcheck = 1;
                // MessageBox.Show(ex.ToString());
            }


        }

        void Thread2_Test()
        {
            UpFile(tfilename2, "/WO Verification/");//D:\jaylai
                                                    //  UpFile(tfilename, "/WO Verification/");//D:\jaylai

            FileInfo fileInfo2 = new FileInfo(tfilename2); fileInfo2.Attributes |= FileAttributes.ReadOnly; fileInfo2.Attributes |= FileAttributes.Hidden;



        }

        void Thread3_Test()
        {
            UpFile(tfilename, "/WO Verification/");
            FileInfo fileInfo = new FileInfo(tfilename); fileInfo.Attributes |= FileAttributes.ReadOnly; fileInfo.Attributes |= FileAttributes.Hidden;
        }
        

        [DllImport("kernel32.dll")]
        public static extern uint GetTickCount();
        public static void Delay(uint ms)
        {
            uint start = GetTickCount();
            while (GetTickCount() - start < ms)
            {
                System.Windows.Forms.Application.DoEvents();
            }
        }
        void Spec_TransintfEvent(int valuet)
        {
           testspec = valuet;
        }
        private void specToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spec ff = new Spec();
            ff.TransintfEvent += Spec_TransintfEvent;
            DialogResult ddr = ff.ShowDialog();
        }

        private void TEST_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text != "")&&(textBox1.Text.IndexOf("-") != -1))
            {
                fourcount = 0;pan = 0;
                wo = textBox1.Text;
                debug = false;
                dpan = 1;
                TEST.Enabled = false;
                try
                {
                    con.Open();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    this.label1.ForeColor = Color.Red;
                    this.label1.Text = "SQL Err";
                }
                string sqlcomm = "select *from [Tesla].[dbo].[";
                sqlcomm = sqlcomm + pmodel + "]  WHERE WO='" + textBox1.Text + "' AND Testresult='PASS'";
                //  MessageBox.Show(sqlcomm);
                SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                DataSet thisDataset = new DataSet();
                try
                {
                    SqlDap.Fill(thisDataset, pmodel);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }
                dt = thisDataset.Tables[pmodel];
                label6.Text = "已刷数量:" + dt.Rows.Count.ToString();
                string filename = "", filename2 = "";
                string[] lines;
                int total, done, done2;
                FileStream savefile;
                textBox1.ReadOnly = true;
                filename = filepath + "\\" + textBox1.Text.ToUpper() + ".txt";
                filename2 = filepath + "\\" + textBox1.Text.ToUpper() + "_No" + ".txt";
                tfilename = filename; tfilename2 = filename2;
              File.SetAttributes(filename, FileAttributes.Normal);  
                string line;               
                lines = File.ReadAllLines(filename);
                total = int.Parse(lines[0].Substring(3));




                if (dt.Rows.Count < total)
                {
                    byte[] sendData;

                    x = 0;
                    testcount = 0;
                    try
                    {
                        sp.Open();
                        if (sp.IsOpen)
                        {
                            Delay(500);
                            sp.NewLine = "\r";
                            sendData = null;
                            sendData = Encoding.UTF8.GetBytes("LON" + "\r\n");
                            sp.Write(sendData, 0, sendData.Length);

                        }

                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);

                    }
                }
else
                { label1.Visible = true;label1.Text = "工单已满";   }
                FileInfo fileInfo = new FileInfo(filename); fileInfo.Attributes |= FileAttributes.ReadOnly; fileInfo.Attributes |= FileAttributes.Hidden;

            }

            if ((textBox1.Text == "")||(textBox1.Text.IndexOf("-")==-1))
            { MessageBox.Show("工单未填"); }



        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sp.IsOpen)
            {
                byte[] closedata;
                closedata = null;
                closedata = Encoding.UTF8.GetBytes("LOFF" + "\r\n");
                sp.Write(closedata, 0, closedata.Length);
                sp.Close();
            }
            try
            {
                con.Close();
            }
            catch (Exception ex)
            {

                label1.Text = ex.ToString();
            }

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {
                string filename, filename2;
                filename = filepath + "\\" + textBox1.Text.ToUpper() + ".txt";
                filename2 = filepath + "\\" + textBox1.Text.ToUpper() + "_No" + ".txt";
                tfilename = filename; tfilename2 = filename2;
                button2.Enabled = true;
                textBox2.Focus();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            thread2 = new Thread(Thread2_Test);
            thread2.IsBackground = true;
            thread2.Start();
            thread3 = new Thread(Thread3_Test);
            thread3.IsBackground = true;
            thread3.Start();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dataexpert ff = new Dataexpert();
            //  ff.TransintfEvent += Spec_TransintfEvent;
           ff.modelset = pmodel;
            DialogResult ddr = ff.ShowDialog();
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true) {/*MessageBox.Show("True");*/ debug = true; dpan = 0; }
            if (checkBox1.Checked == false) { /*MessageBox.Show("false");*/ debug = false;dpan = 1; }
        }

        private void greateWOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OrderSet order = new OrderSet();
           
            DialogResult ddr = order.ShowDialog();
        }

        private void cToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process proc = Process.Start("D:\\SNCheck\\WindowsFormsApp1.exe");

            proc.WaitForExit();//这句加了就会卡在这里，外部程序退出才能继续执行
            /*
            toolStripMenuItem2.Checked = false;
            cToolStripMenuItem.Checked = true;
            pmodel = "M1817";
            con.Open();
            string sqlcomm = "use Tesla select Name from sysobjects where xtype='u'";
            SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
            DataSet thisDataset = new DataSet();
            try
            {
                SqlDap.Fill(thisDataset);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            // MessageBox.Show("name");
            int xpan = 0, xpan2 = 0;
            string ss, tablename = pmodel;
            do
            {
                ss = thisDataset.Tables[0].Rows[xpan][0].ToString();
                if (tablename == ss) { xpan2 = 1; }
                xpan++;
            } while ((xpan < thisDataset.Tables[0].Rows.Count) && (xpan2 == 0));
            if (xpan2 != 0)
            {
                con.Close();
            }
            if (xpan2 == 0)
            {
               // dt = new DataTable("Table_Testdata");
               // dt.Columns.Add("SN", System.Type.GetType("System.String"));
               // dt.Columns.Add("Match", System.Type.GetType("System.String"));
              //  dt.Columns.Add("Testtime", System.Type.GetType("System.String"));
               // dt.Columns.Add("Testresult", System.Type.GetType("System.String"));
                string sqlcomm2 = "USE Tesla" + "\r\n";
                sqlcomm2 = sqlcomm2 + "CREATE TABLE " + pmodel + "(SN varchar(50) NOT NULL, Match varchar(50) NOT NULL, Readtime varchar(50) NOT NULL, Testtime varchar(50) NOT NULL, Testresult varchar(50) NOT NULL)";
                //   MessageBox.Show(sqlcomm2);
                cmd = new SqlCommand(sqlcomm2, con);
                //执行sql命令
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }
                con.Close();
            }
*/

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            toolStripMenuItem2.Checked = true;
            cToolStripMenuItem.Checked = false;
            pmodel = "M1829";
            con.Open();
           string sqlcomm = "use Tesla select Name from sysobjects where xtype='u'";
            SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
            DataSet thisDataset = new DataSet();
            try
            {
                SqlDap.Fill(thisDataset);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
         //   MessageBox.Show("name");
            int xpan = 0, xpan2 = 0;
            string ss, tablename = pmodel;
            do
            {
                ss = thisDataset.Tables[0].Rows[xpan][0].ToString();
              //  MessageBox.Show(ss);
                if (tablename == ss) { xpan2 = 1; }
                xpan++;
            } while ((xpan < thisDataset.Tables[0].Rows.Count) && (xpan2 == 0));
            if (xpan2 !=0)
            {
                con.Close();
            }
                if (xpan2 == 0)
            {
                //MessageBox.Show("name1");
                /* dt = new DataTable("Table_Testdata");
                dt.Columns.Add("SN", System.Type.GetType("System.String"));
                dt.Columns.Add("Match", System.Type.GetType("System.String"));
                dt.Columns.Add("Testtime", System.Type.GetType("System.String"));*/
                string sqlcomm2 = "USE Tesla" + "\r\n";
                sqlcomm2 = sqlcomm2 + "CREATE TABLE " + pmodel + "(WO varchar(50) NOT NULL,SN varchar(50) NOT NULL, Match varchar(50) NOT NULL, Readtime varchar(50) NOT NULL,  Testtime varchar(50) NOT NULL, Testresult varchar(50) NOT NULL)";
             //   MessageBox.Show(sqlcomm2);
                cmd = new SqlCommand(sqlcomm2, con);
                //执行sql命令
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }
                con.Close();
               // MessageBox.Show("name2");

            }


        }

        void frm_TransfEvent()
        {
            this.TEST.Enabled = true;
        }


        private void comSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ComSet ff = new ComSet();
            ff.TransfEvent += frm_TransfEvent;
            DialogResult ddr = ff.ShowDialog();

        }
        DataGridViewImageColumn img = new DataGridViewImageColumn();
        private void Form1_Load(object sender, EventArgs e)
        {
            System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName("SR1000");//获取指定的进程名   
            if (myProcesses.Length > 1) //如果可以获取到知道的进程名则说明已经启动
            {
                MessageBox.Show("程序已启动！");
                System.Windows.Forms.Application.Exit();              //关闭系统

            }
            this.label1.Font = new Font("隶书", 18, FontStyle.Bold); //第一个是字体，第二个大小，第三个是样式，
            this.label1.ForeColor = Color.Blue;// 颜色 
            this.label2.Font = new Font("隶书", 23, FontStyle.Bold); //第一个是字体，第二个大小，第三个是样式，
            this.label2.ForeColor = Color.Blue;// 颜色 
            string[] p;
            string comstr;
            int pi, i, ii;
            p = new string[6];
            configtxt = "d:\\comconfig.txt";
            pi = 0;
            if (!File.Exists(configtxt))
            {
                fs = new FileStream(configtxt, FileMode.Create);

                sw2 = new StreamWriter(fs);
                sw2.Write("COM1" + "\r\n");
                sw2.Write("115200" + "\r\n");
                sw2.Write("8" + "\r\n");
                sw2.Write("even" + "\r\n");
                sw2.Write("1" + "\r\n");
                sw2.Write("None" + "\r\n");
                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    comstr = sr2.ReadLine();
                    p[pi] = comstr;
                    pi = pi + 1;
                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                //    MessageBox.Show(p[0]);
                sp = new SerialPort(p[0]);
                
            }
            else
            {
                fs = new FileStream(configtxt, FileMode.Open);
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    comstr = sr2.ReadLine();
                    p[pi] = comstr;
                    pi = pi + 1;
                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                //    MessageBox.Show(p[0]);
                sp = new SerialPort(p[0]);
  
            }
            comtrue = true;


                try
                {
                    sp.BaudRate = int.Parse(p[1]);
                    sp.DataBits = int.Parse(p[2]);
                    if (p[3] == "even")
                    {
                        sp.Parity = Parity.Even;
                    }
                    if (p[3] == "odd")
                    {
                        sp.Parity = Parity.Odd;
                    }
                    if (p[3] == "mark")
                    {
                        sp.Parity = Parity.Mark;
                    }
                    if (p[3] == "none")
                    {
                        sp.Parity = Parity.None;
                    }
                    if (p[3] == "space")
                    {
                        sp.Parity = Parity.Space;
                    }
                    if (p[4] == "1")
                    {
                        sp.StopBits = StopBits.One;
                    }
                    if (p[4] == "2")
                    {
                        sp.StopBits = StopBits.Two;
                    }
                    if (p[5] == "none")
                    {
                        sp.Handshake = Handshake.None;
                    }
                    if (p[5] == "send")
                    {
                        sp.Handshake = Handshake.RequestToSend;
                    }
                    if (p[5] == "xonxoff")
                    {
                        sp.Handshake = Handshake.XOnXOff;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                comtrue = false;
                this.label1.ForeColor = Color.Red;
                this.label1.Text = "Com Err";
                TEST.Enabled = false;
                }
            try
            {

                sp.Open();

                //sp.Close();
                // MessageBox.Show("OK");
            }

            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                comtrue = false;
                TEST.Enabled = false;
                this.label1.ForeColor = Color.Red;
                this.label1.Text = "Com Err";
            }
            sp.Close();
if(comtrue==true)
            {
                this.label1.ForeColor = Color.Blue;
                this.label1.Text = "Ready";
            }

            sp.ReceivedBytesThreshold = 1;
            sp.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(sp_DataReceived);

            ticonfig = "d:\\Specconfig.txt";
            if (!File.Exists(ticonfig))
            {
                fs = new FileStream(ticonfig, FileMode.Create);

                sw2 = new StreamWriter(fs);
                sw2.Write("70" + "\r\n");
                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();
                sr2 = new StreamReader(fs);
                string specbuf="";
                while (!sr2.EndOfStream)
                {
                    specbuf = sr2.ReadLine();
                   
                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                testspec = int.Parse(specbuf);


            }
            else
            {
                fs = new FileStream(ticonfig, FileMode.Open);
                sr2 = new StreamReader(fs);
                string specbuf = "";
                while (!sr2.EndOfStream)
                {
                    specbuf = sr2.ReadLine();

                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                testspec = int.Parse(specbuf);

            }



                dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].HeaderText = "SN";
            dataGridView1.Columns[0].Width = 355;
            dataGridView1.Columns[1].HeaderText = "Matching Rate";
            dataGridView1.Columns[1].Width = 105;
            dataGridView1.Columns[2].HeaderText = "ReadTime";
            dataGridView1.Columns[2].Width = 115;
           
            dataGridView1.Columns[3].HeaderText = "LocalTime";
            dataGridView1.Columns[3].Width = 115;
          
            dataGridView1.RowsDefaultCellStyle.Font = new Font("楷体", 15, FontStyle.Bold);
            dataGridView1.ReadOnly = true;
        
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Columns.Insert(0, img);
            img.HeaderText = "Result";
            con.ConnectionString = "server=90EK999ZZVXNUX9\\SQLEXPRESS;database=Tesla;uid=sa;pwd=sqlte";
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
                this.label1.ForeColor = Color.Red;
                this.label1.Text = "SQL Err";
            }
            try
            {
                con.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
                this.label1.ForeColor = Color.Red;
                this.label1.Text = "SQL Err";
                TEST.Enabled = false;
            }






        }
        int testcount ;
        private void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //MessageBox.Show("kaishi");
            string result, result2, comerr,wavepath;
            result = "";
            int ci;
            FileStream fs = null;
            Encoding encoder = Encoding.UTF8;
            Byte[] receivedData = new Byte[sp.BytesToRead];
            byte[] sendData;
            DataTable dt1;
            wavepath = "D:\\photo.wav";
            result = sp.ReadLine();
            //sp.Read(receivedData, 0, receivedData.Length);
           // result = new UTF8Encoding().GetString(receivedData);
            //MessageBox.Show(result);
          /*  if (result=="")
            {
                
                this.Invoke((EventHandler)(delegate
                {
                    this.label1.ForeColor = Color.Blue;
                    this.label1.Text = "Wait";
                }
                           ));
                sp.WriteLine("LOFF");
                Delay(300);
                sp.WriteLine("LON");

            }*/
            if ((result.IndexOf("(P)1071917-00-E") != -1)||(result.IndexOf("(P)1071917-00-F") != -1))
            {
                int mresult,tresult;
                string[] seq;
                string[] seq2;
                string ps;
                string ssresult;
                int xi;
                xi = 0;
                ssresult = "";
                ps = "";
                seq2 = new string[3];
                seq = result.Split(':');
                foreach (string azu in seq)
                {
                    seq2[xi] = azu;
                    xi++;

                }
                int  testresult = 0;
                mresult = int.Parse(seq2[1]);
               tresult= int.Parse(seq2[2].Substring(0, seq2[2].Length-2));
                //  MessageBox.Show(tresult.ToString());
                //MessageBox.Show(testspec.ToString());
                if ((mresult>=testspec)&&(tresult<1200))
                    { testresult = 1;testcount = 0; }
                    else
                    { testcount++; }
if(testresult == 1)
                {

                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + pmodel + "] where SN =" + "'" + seq2[0] + "' AND Testresult='PASS'";
                    //  MessageBox.Show(sqlcomm);
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, pmodel);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }
                    dt = thisDataset.Tables[pmodel];
                    // MessageBox.Show(thisDataset.Tables[0].Rows[0][4].ToString());
                    int panx = 0,writep=1;
                    if(dt.Rows.Count == 0)
                    { writep = 1; }

                    if (dt.Rows.Count != 0)
                    {
                        writep = 0;
                        /*  do
                        {
                            if (thisDataset.Tables[0].Rows[panx][5].ToString() == "PASS")
                            { writep = 0; }
                            panx++;
                        } while ((panx < dt.Rows.Count) && (thisDataset.Tables[0].Rows[0][5].ToString() != "PASS"));
                        */
                    }
                    // if ((dt.Rows.Count != 0) && (thisDataset.Tables[0].Rows[0][5].ToString() == "PASS"))
                   if(writep==0)
                    {
                            this.Invoke((EventHandler)(delegate
                            {

                               // MessageBox.Show(seq2[0]);
                                this.label1.ForeColor = Color.Blue;
                                this.label1.Text = "SN Repeat";
                                this.label2.ForeColor = Color.Blue;
                                this.label2.Text = seq2[0];
                                FileStream fs2 = new FileStream("D:\\trafficlight_green_128px_1187357_easyicon.net.ico", FileMode.Open,
                                                                                                FileAccess.Read);//获取图片文件流
                                Image img = Image.FromStream(fs2); // 文件流转换成Image格式
                                pictureBox1.Image = img;   //给 图片框设置要显示的图片
                                fs2.Close(); // 关闭
                                 
                                             

                            }

                                                                       ));

                     
                     
                        
                        //  if (debug == false)
                     //   MessageBox.Show(dpan.ToString());
                        if (dpan == 1)
                        {
                            // MessageBox.Show(dpan.ToString());
                            
                            sp.Close();
                            sp.Open();
                           // MessageBox.Show(dpan.ToString()+"again");
                            Delay(1000);
                            sp.NewLine = "\r";
                            sendData = null;
                            sendData = Encoding.UTF8.GetBytes("LON" + "\r\n");
                            sp.Write(sendData, 0, sendData.Length);
                        }
                        //if (debug == true)
                        if (dpan == 0)
                        {
                            sp.Close();
                            con.Close();
                            TEST.Enabled = true;
                        }

                    }

                    string snpass = "ok";
                  //  if ((dt.Rows.Count == 0)||((dt.Rows.Count != 0)&& (thisDataset.Tables[0].Rows[0][5].ToString() == "FAIL")))
                  if (writep==1)
                    {
                        //if (debug == false)
                            if (dpan == 1)
                            {

/*
                       
*/
//
//
                            string filename = "", filename2 = "";
                            string[] lines;
                            int total, done, done2;
                            FileStream savefile;
                            textBox1.ReadOnly = true;
                            filename = filepath + "\\" + textBox1.Text.ToUpper() + ".txt";
                            filename2 = filepath + "\\" + textBox1.Text.ToUpper() + "_No" + ".txt";
                            tfilename = filename; tfilename2 = filename2;                           
                            File.SetAttributes(filename2, FileAttributes.Normal); File.SetAttributes(filename, FileAttributes.Normal);
                          //  MessageBox.Show(tfilename);
                           // MessageBox.Show(tfilename2);
                           // StreamReader check = new StreamReader(filename, Encoding.Default);
                            string line;
                            lines = File.ReadAllLines(filename2);
                            done2 = int.Parse(lines[0]) + 1;
                            lines = File.ReadAllLines(filename);
                            total = int.Parse(lines[0].Substring(3));
                            //MessageBox.Show(done2.ToString());
                            //MessageBox.Show(total.ToString());
                            if (done2 > total)
                            {
                                this.Invoke((EventHandler)(delegate
{
    this.label1.ForeColor = Color.Red; this.label1.Text = "数量超出";
}
 ));                             
snpass = "fail"; FileInfo fileInfo2 = new FileInfo(filename2); fileInfo2.Attributes |= FileAttributes.ReadOnly; fileInfo2.Attributes |= FileAttributes.Hidden;
                                FileInfo fileInfo = new FileInfo(filename); fileInfo.Attributes |= FileAttributes.ReadOnly; fileInfo.Attributes |= FileAttributes.Hidden;



                            }
                            if (snpass == "ok")
                            {
                               
                                savefile = new FileStream(filename, FileMode.Append);
                               // MessageBox.Show(seq2[0]);
                                StreamWriter sw3 = new StreamWriter(savefile);
                              //  MessageBox.Show(seq2[0]);
                                sw3.Write(seq2[0] + "," + DateTime.Now.ToString("MM/dd/yyyy") + "," + textBox2.Text + "\r\n");  sw3.Flush(); sw3.Close(); savefile.Close();
                                lines = File.ReadAllLines(filename2);
                                fourcount = fourcount + 1;
                                if (fourcount % 2 == 0) { pan = pan + 1; }
                                done = int.Parse(lines[0]) + 1;
                                label5.Text = "盘数" + ":" + pan.ToString(); label6.Text = "已刷数量" + ":" + done.ToString();
                                savefile = new FileStream(filename2, FileMode.Truncate);
                                sw3 = new StreamWriter(savefile);
                                sw3.Write(done.ToString()); sw3.Flush(); sw3.Close(); savefile.Close();
                                if (done != total)
                                {
                                    FileInfo fileInfo2 = new FileInfo(filename2); fileInfo2.Attributes |= FileAttributes.ReadOnly; fileInfo2.Attributes |= FileAttributes.Hidden;
                                    FileInfo fileInfo = new FileInfo(filename); fileInfo.Attributes |= FileAttributes.ReadOnly; fileInfo.Attributes |= FileAttributes.Hidden;
                                }
                                if (done == total)
                                {                                
                                        thread2 = new Thread(Thread2_Test);
                                        thread2.IsBackground = true;
                                        thread2.Start();
                                        thread3 = new Thread(Thread3_Test);
                                        thread3.IsBackground = true;
                                        thread3.Start();                             
                                }
                                dt1 = new DataTable("Table_Testdata");
                                dt1.Columns.Add("WO", System.Type.GetType("System.String"));
                                dt1.Columns.Add("SN", System.Type.GetType("System.String"));
                                dt1.Columns.Add("Match", System.Type.GetType("System.String"));
                                dt1.Columns.Add("Readtime", System.Type.GetType("System.String"));
                                dt1.Columns.Add("Testtime", System.Type.GetType("System.String"));
                                dt1.Columns.Add("Testresult", System.Type.GetType("System.String"));
                                DataRow dr = dt1.NewRow();
                                dr[0] = wo;
                                dr[1] = seq2[0];
                                dr[2] = seq2[1];
                                dr[3] = tresult.ToString();
                                dr[4] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                dr[5] = "PASS";
                                dt1.Rows.Add(dr);
                                SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null);
                                bulkCopy.DestinationTableName = pmodel;
                                bulkCopy.BatchSize = 1;
                                if (dt1 != null && dt1.Rows.Count != 0)
                                {
                                    try
                                    {
                                        bulkCopy.WriteToServer(dt1);
                                    }
                                    catch (System.Exception ex)
                                    {
                                        MessageBox.Show(ex.ToString());

                                    }


                                }
                                //
                            }







                        }//PASS写入
                        //   if (dt.Rows.Count == 0)
                      //  {
                            this.Invoke((EventHandler)(delegate
                        {
                            this.dataGridView1.Rows.Add();
                            this.dataGridView1.Rows[x].Cells[1].Value = seq2[0];
                            this.dataGridView1.Rows[x].Cells[2].Value = seq2[1];
                            this.dataGridView1.Rows[x].Cells[3].Value = seq2[2];
                            this.dataGridView1.Rows[x].Cells[0].Value = Image.FromFile("D:\\greenled_14px_35209_easyicon.net.ico");
                            this.dataGridView1.Rows[x].Cells[4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Green;
                            this.label1.Text = "PASS";
                            this.label2.ForeColor = Color.Green;
                            this.label2.Text = seq2[0];
                            FileStream fs2 = new FileStream("D:\\trafficlight_green_128px_1187357_easyicon.net.ico", FileMode.Open,
                                                                                                       FileAccess.Read);//获取图片文件流
                            Image img = Image.FromStream(fs2); // 文件流转换成Image格式
                            pictureBox1.Image = img;   //给 图片框设置要显示的图片
                            fs2.Close(); // 关闭
                                         //if(debug==true)
                           


                        }

                                                ));

                        //  if (debug == false)
                        if ((dpan == 1)&&(snpass=="ok"))
                        {
                            x++;
                            System.Media.SoundPlayer player = new System.Media.SoundPlayer(wavepath);
                            player.PlaySync();//
                            sp.Close();
                            sp.Open();
                            Delay(300);

                            sp.NewLine = "\r";
                            sendData = null;
                            sendData = Encoding.UTF8.GetBytes("LON" + "\r\n");
                            sp.Write(sendData, 0, sendData.Length);
                        }
                        // if (debug == true)
                        if (dpan == 0)
                        {
                            sp.Close();
                            con.Close();
                            TEST.Enabled = true;
                        }
                        //  }
                        /*      if (dt.Rows.Count != 0)
                              {
                                  this.Invoke((EventHandler)(delegate
                                  {

                                      this.label1.ForeColor = Color.Blue;
                                      this.label1.Text = "SN Repeat";
                                      this.label2.ForeColor = Color.Blue;
                                      this.label2.Text = seq2[0];
                                      FileStream fs2 = new FileStream("D:\\ball_yellow_128px_15179_easyicon.net", FileMode.Open,
                                                                                                              FileAccess.Read);//获取图片文件流
                                      Image img = Image.FromStream(fs2); // 文件流转换成Image格式
                                      pictureBox1.Image = img;   //给 图片框设置要显示的图片
                                      fs2.Close(); // 关闭

                                  }

                                                                 ));

                              }
      */
                    }
                }
                
if((testcount < 5) && (testresult == 0))
                {
                    sp.Close();
                    sp.Open();
                    Delay(300);
                    sp.NewLine = "\r";
                    sendData = null;
                    sendData = Encoding.UTF8.GetBytes("LON" + "\r\n");
                    sp.Write(sendData, 0, sendData.Length);
                }
                if ((testcount >= 5) && (testresult == 0))
                {
                    testcount = 0;
                  //  MessageBox.Show(result);
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + pmodel + "] where SN ="+"'"+ seq2[0]+"'";
                  //  MessageBox.Show(sqlcomm);
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "TestData");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }
                    dt = thisDataset.Tables["Testdata"];
                    //  MessageBox.Show(dt.Rows.Count.ToString());
                    if (dt.Rows.Count != 0)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            this.dataGridView1.Rows.Add();
                            this.dataGridView1.Rows[x].Cells[1].Value = seq2[0];
                            this.dataGridView1.Rows[x].Cells[2].Value = seq2[1];
                            this.dataGridView1.Rows[x].Cells[3].Value = seq2[2];
                            this.dataGridView1.Rows[x].Cells[0].Value = Image.FromFile("D:\\redled_14px_36067_easyicon.net.ico");
                            this.dataGridView1.Rows[x].Cells[4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                            this.label2.ForeColor = Color.Red;
                            this.label2.Text = seq2[0];
                            FileStream fs2 = new FileStream("D:\\trafficlight_red_128px_1187358_easyicon.net.ico", FileMode.Open,
                                                                           FileAccess.Read);//获取图片文件流
                            Image img = Image.FromStream(fs2); // 文件流转换成Image格式
                            pictureBox1.Image = img;   //给 图片框设置要显示的图片
                            fs2.Close(); // 关闭





                        }

                                               ));

                        x++;
                        sp.Close();
                        con.Close();
                       // Delay(300);
                       
                    }
                        if(dt.Rows.Count==0)
                    {
                        //if (debug == false)
                        if (dpan == 1)
                        {
                            dt1 = new DataTable("Table_Testdata");
                            dt1.Columns.Add("WO", System.Type.GetType("System.String"));
                            dt1.Columns.Add("SN", System.Type.GetType("System.String"));
                            dt1.Columns.Add("Match", System.Type.GetType("System.String"));
                            dt1.Columns.Add("Readtime", System.Type.GetType("System.String"));
                            dt1.Columns.Add("Testtime", System.Type.GetType("System.String"));
                            dt1.Columns.Add("Testresult", System.Type.GetType("System.String"));
                            DataRow dr = dt1.NewRow();
                            dr[0] = wo;
                            dr[1] = seq2[0];
                            dr[2] = seq2[1];
                            dr[3] = tresult.ToString();
                            dr[4] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            dr[5] = "FAIL";
                            dt1.Rows.Add(dr);
                            SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null);
                            bulkCopy.DestinationTableName = pmodel;
                            bulkCopy.BatchSize = 1;
                            if (dt1 != null && dt1.Rows.Count != 0)
                            {
                                try
                                {
                                    bulkCopy.WriteToServer(dt1);
                                }
                                catch (System.Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());

                                }
                            }
                        }

                        this.Invoke((EventHandler)(delegate
                        {
                            this.dataGridView1.Rows.Add();
                            this.dataGridView1.Rows[x].Cells[1].Value = seq2[0];
                            this.dataGridView1.Rows[x].Cells[2].Value = seq2[1];
                            this.dataGridView1.Rows[x].Cells[3].Value = seq2[2];
                            this.dataGridView1.Rows[x].Cells[0].Value = Image.FromFile("D:\\redled_14px_36067_easyicon.net.ico");
                            this.dataGridView1.Rows[x].Cells[4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                            this.label2.ForeColor = Color.Red;
                            this.label2.Text = seq2[0];
                            FileStream fs2 = new FileStream("D:\\trafficlight_red_128px_1187358_easyicon.net.ico", FileMode.Open,
                                                                           FileAccess.Read);//获取图片文件流
                            Image img = Image.FromStream(fs2); // 文件流转换成Image格式
                            pictureBox1.Image = img;   //给 图片框设置要显示的图片
                            fs2.Close(); // 关闭





                        }

                                                ));

                        x++;
                        sp.Close();
                       // sp.Open();
                        // Delay(300);
                        con.Close();
                      //  TEST.Enabled = true;
                        /*  sp.NewLine = "\r";
                        sendData = null;
                        sendData = Encoding.UTF8.GetBytes("LON" + "\r\n");
                        sp.Write(sendData, 0, sendData.Length);*/

                    }



                }
                   /* this.Invoke((EventHandler)(delegate
                {
                    this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[x].Cells[0].Value = seq2[1];
                    this.dataGridView1.Rows[x].Cells[1].Value = seq2[2];
                    this.dataGridView1.Rows[x].Cells[2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


                }

                        ));

                x++;
                sp.Close();
                sp.Open();
                Delay(300);
                
                sp.NewLine = "\r";
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("LON"+"\r\n");
                sp.Write(sendData, 0, sendData.Length);*/
                
            }//result p



        }





    }
}
