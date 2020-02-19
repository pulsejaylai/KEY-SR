using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace SR1000
{
    public partial class Spec : Form
    {
        public Spec()
        {
            InitializeComponent();
        }
        FileStream fs;
        StreamWriter sw2;
        StreamReader sr2;
        string configtxt, comstr;
        public delegate void TransfDelegate(String value);
        public delegate void TransintfDelegate(int valuet);
        private void Spec_Load(object sender, EventArgs e)
        {
            int x = (System.Windows.Forms.SystemInformation.WorkingArea.Width - this.Size.Width) / 2;
            int y = (System.Windows.Forms.SystemInformation.WorkingArea.Height - this.Size.Height) / 1;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = (Point)new Size(x, y);
            textBox1.Focus();
            configtxt = "d:\\Specconfig.txt";
            fs = new FileStream(configtxt, FileMode.Open);
            sr2 = new StreamReader(fs);
            string specbuf = "";
            while (!sr2.EndOfStream)
            {
                specbuf = sr2.ReadLine();

                //  MessageBox.Show(comstr);
            }

            //  MessageBox.Show(comstr);
            fs.Close();
            textBox1.Text = specbuf;
        }
        public event TransfDelegate TransfEvent;
        public event TransintfDelegate TransintfEvent;
        private void button1_Click(object sender, EventArgs e)
        {
if(textBox1.Text=="")
            { MessageBox.Show("Spec Null"); }

if(textBox1.Text != "")
            {
                fs = new FileStream(configtxt, FileMode.Truncate);
                sw2 = new StreamWriter(fs);
                sw2.Write(textBox1.Text + "\r\n");
                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();
                TransintfEvent(int.Parse(textBox1.Text));
                this.Close();
            }



        }









    }
}
