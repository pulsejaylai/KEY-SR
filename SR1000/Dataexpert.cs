using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using System.IO;

namespace SR1000
{
    public partial class Dataexpert : Form
    {
        public Dataexpert()
        {
            InitializeComponent();
        }
        SqlConnection con = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        //   DataTable dt;
        DataSet dsfill=new DataSet();
        private string modelname = "";
        public string modelset
        {
            set
            {
                modelname = value;
            }
        }
        private void Dataexpert_Load(object sender, EventArgs e)
        {
          /*  dataGridView1.ColumnCount = 5;
            dataGridView1.Columns[0].HeaderText = "SN";
            dataGridView1.Columns[0].Width = 255;
            dataGridView1.Columns[1].HeaderText = "Matching Rate";
            dataGridView1.Columns[1].Width = 105;
            dataGridView1.Columns[2].HeaderText = "ReadTime";
            dataGridView1.Columns[2].Width = 255;
            dataGridView1.Columns[3].HeaderText = "LocalTime";
            dataGridView1.Columns[3].Width = 255;
            dataGridView1.Columns[4].HeaderText = "Result";
            dataGridView1.Columns[4].Width = 100;*/
            dataGridView1.ReadOnly = true;

            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
   
            pagerControl1.OnPageChanged += new EventHandler(pagerControl1_OnPageChanged);
            /* con.ConnectionString = "server=90EK999ZZVXNUX9\\SQLEXPRESS;database=Tesla;uid=sa;pwd=sqlte";
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
               
            }
            string sqlcomm = "select *from [Tesla].[dbo].[";
            sqlcomm = sqlcomm + modelname + "]";
            SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
            DataSet thisDataset = new DataSet();
            try
            {
                SqlDap.Fill(thisDataset, "Testdata");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            //this.dataGridView1.DataSource = thisDataset;
            //this.dataGridView1.DataMember = "Testdata";
            this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
            dt = thisDataset.Tables["Testdata"];*/
            // dataGridView1.ItemsSource = dt.DefaultView;
            /*   try
               {
                   con.Close();
               }
               catch (System.Exception ex)
               {
                   MessageBox.Show(ex.ToString());

               }
   */
        }

        private void pagerControl1_OnPageChanged(object sender, EventArgs e)
        {

            LoadData(dsfill);
        }
        private void LoadData(DataSet ds)
        {
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            int rowcount = 0;
            rowcount = ds.Tables[0].Rows.Count;
            pagerControl1.RecordCount = rowcount;
        //    if(pagerControl1.PageIndex==1)
           // {
                dt1 = ds.Tables[0];
                dt2 = dt1.Clone();
            if(pagerControl1.PageIndex < (pagerControl1.PageCount ))
            { 
                for (int i= pagerControl1.PageSize*(pagerControl1.PageIndex-1); i< pagerControl1.PageSize* pagerControl1.PageIndex; i++)
                { dt2.Rows.Add(dt1.Rows[i].ItemArray); }
                this.dataGridView1.DataSource = dt2;
            }
            else
            {
                for (int i = 0; i < rowcount- (pagerControl1.PageSize * (pagerControl1.PageIndex - 1)); i++)
                { dt2.Rows.Add(dt1.Rows[i+ pagerControl1.PageSize * (pagerControl1.PageIndex - 1)].ItemArray); }
                this.dataGridView1.DataSource = dt2;

            }
            /* if ((pagerControl1.PageIndex>1)&& (pagerControl1.PageIndex < pagerControl1.PageCount+1))
            {
                dt1 = ds.Tables[0];
                dt2 = dt1.Clone();
                

            }
            */

        }


        public static bool DataTableToExcel(DataTable dt, string path)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    //  sheet = new ISheet();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                                                           //sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet0的表  
                    int rowCount = dt.Rows.Count;//行数  
                    int columnCount = dt.Columns.Count;//列数  
                                                       //  MessageBox.Show(rowCount.ToString());
                                                       //  MessageBox.Show(columnCount.ToString());
                                                       //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        // MessageBox.Show(DateTime.Now.Month.ToString());
                        //   row = sheet[l].CreateRow(i);
                        // MessageBox.Show(i.ToString());
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                                                     //   MessageBox.Show(dt.Rows[i][j].ToString());
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }

                    }
                    // fs = File.Create(path)
                    using (fs = File.Create(path))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据  
                        result = true;
                    }
                    fs.Close();
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }


        }


        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            con.ConnectionString = "server=90EK999ZZVXNUX9\\SQLEXPRESS;database=Tesla;uid=sa;pwd=sqlte";
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            if ((textBox1.Text != "") && (textBox2.Text != ""))
            {
                string sqlcomm = "select *from [Tesla].[dbo].[";
                sqlcomm = sqlcomm + modelname + "] WHERE WO='" + textBox1.Text + "' AND SN='" + textBox2.Text + "'";
                SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                DataSet thisDataset = new DataSet();
                try
                {
                    SqlDap.Fill(thisDataset, "Testdata");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }

              //  this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                dt = thisDataset.Tables["Testdata"];


            }//!=""
            else
            {
                if (textBox1.Text != "")
                {
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + modelname + "] WHERE WO='" + textBox1.Text + "'";
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "Testdata");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }

                 //   this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                    dt = thisDataset.Tables["Testdata"];
                }
                if (textBox2.Text != "")
                {
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + modelname + "] WHERE SN='" + textBox2.Text + "'";
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "Testdata");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }

                  //  this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                    dt = thisDataset.Tables["Testdata"];

                }
                if ((textBox1.Text == "") && (textBox2.Text == ""))
                {
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + modelname + "]";
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "Testdata");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }

                  //  this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                    dt = thisDataset.Tables["Testdata"];


                }



            }//else

            try
            {
                con.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            DataTableToExcel(dt, "D:\\" + modelname + ".xls");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            con.ConnectionString = "server=90EK999ZZVXNUX9\\SQLEXPRESS;database=Tesla;uid=sa;pwd=sqlte";
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            if ((textBox1.Text!="")&&(textBox2.Text != ""))
            {
                string sqlcomm = "select *from [Tesla].[dbo].[";
                sqlcomm = sqlcomm + modelname + "] WHERE WO='"+ textBox1.Text +"' AND SN='"+textBox2.Text+"'";
                SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                DataSet thisDataset = new DataSet();
                try
                {
                    SqlDap.Fill(thisDataset, "Testdata");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }

                // this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                dsfill = thisDataset;
                LoadData(thisDataset);
                pagerControl1.DrawControl(thisDataset.Tables[0].Rows.Count);
                dt = thisDataset.Tables["Testdata"];


            }//!=""
            else
            {
               if(textBox1.Text != "")
                {
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + modelname + "] WHERE WO='" + textBox1.Text + "'";
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "Testdata");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }

                    //  this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                    dsfill = thisDataset;
                    LoadData(thisDataset);
                    pagerControl1.DrawControl(thisDataset.Tables[0].Rows.Count);
                    dt = thisDataset.Tables["Testdata"];
                }
                if (textBox2.Text != "")
                {
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + modelname + "] WHERE SN='" + textBox2.Text + "'";
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "Testdata");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }

                    // this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                    dsfill = thisDataset;
                    LoadData(thisDataset);
                    pagerControl1.DrawControl(thisDataset.Tables[0].Rows.Count);
                    dt = thisDataset.Tables["Testdata"];

                }
                if ((textBox1.Text == "") && (textBox2.Text == ""))
                {
                    string sqlcomm = "select *from [Tesla].[dbo].[";
                    sqlcomm = sqlcomm + modelname + "]";
                    SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                    DataSet thisDataset = new DataSet();
                    try
                    {
                        SqlDap.Fill(thisDataset, "Testdata");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }

                    //   this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
                    dsfill = thisDataset;
                    LoadData(thisDataset);
                    pagerControl1.DrawControl(thisDataset.Tables[0].Rows.Count);
                    dt = thisDataset.Tables["Testdata"]; 


                }



            }//else

            try
            {
                con.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }



        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            con.ConnectionString = "server=90EK999ZZVXNUX9\\SQLEXPRESS;database=Tesla;uid=sa;pwd=sqlte";
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            string sqlcomm = "select *from [Tesla].[dbo].[";
            sqlcomm = sqlcomm + modelname + "] WHERE WO='" + textBox1.Text + "' AND Testresult='PASS'";
            SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
            DataSet thisDataset = new DataSet();
            try
            {
                SqlDap.Fill(thisDataset, "Testdata");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }

           // this.dataGridView1.DataSource = thisDataset.Tables["Testdata"];
            dt = thisDataset.Tables["Testdata"];

            label3.Text = "数量 : " + dt.Rows.Count.ToString();



        }









    }
}
