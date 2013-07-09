using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using Microsoft.Win32;
using System.Runtime.InteropServices;

using System.Text;
using System.Windows.Forms;

namespace Table_Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            mainfig();
            //rungraph();
         
           
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
           
            mainfig();
      
        }
        private void mainfig()
        {
            GC.Collect();
            timer1.Enabled = false;
            Data.DataLink getmedata = new Data.DataLink();
            getmedata.testingonly();

            if (getmedata.col_9 != "Q")
            {
                Table1_Count_1.Text = null;
                Table2_Count_1.Text = null;
                //label_Hour1.Text = null;
                //label_Hour2.Text = null;

                Table1_Count_1.ForeColor = System.Drawing.Color.DarkBlue;
                Table2_Count_1.ForeColor = System.Drawing.Color.DarkBlue;
                //label_Hour1.ForeColor = System.Drawing.Color.Green;
                //label_Hour2.ForeColor = System.Drawing.Color.Green;
                label_average1.ForeColor = System.Drawing.Color.Green;
                label_average2.ForeColor = System.Drawing.Color.Green;
                label_first1.ForeColor = System.Drawing.Color.Green;
                label_first2.ForeColor = System.Drawing.Color.Green;
                label_last1.ForeColor = System.Drawing.Color.Green;
                label_last2.ForeColor = System.Drawing.Color.Green;
                Table1_Count_1.Text = getmedata.col_2;
                Table2_Count_1.Text = getmedata.col_1;
                //label_Hour1.Text = getmedata.col_4;
                //label_Hour2.Text = getmedata.col_3;
                label_average1.Text = getmedata.averagecut1;
                label_average2.Text = getmedata.averagecut2;
                label_first1.Text = getmedata.starttime1;
                label_first2.Text = getmedata.starttime2;
                label_last1.Text = getmedata.lasttime1;
                label_last2.Text = getmedata.lasttime2;


            }
            else
            {
                Table1_Count_1.ForeColor = System.Drawing.Color.Red;
                Table2_Count_1.ForeColor = System.Drawing.Color.Red;
                //label_Hour1.ForeColor = System.Drawing.Color.Red;
                //label_Hour2.ForeColor = System.Drawing.Color.Red;

                label_average1.ForeColor = System.Drawing.Color.Red;
                label_average2.ForeColor = System.Drawing.Color.Red;
                label_first1.ForeColor = System.Drawing.Color.Red;
                label_first2.ForeColor = System.Drawing.Color.Red;
                label_last1.ForeColor = System.Drawing.Color.Red;
                label_last2.ForeColor = System.Drawing.Color.Red;

            }

            Data.DataLink getmedata2 = new Data.DataLink();
            getmedata2.testingonly2();

            if (getmedata2.col_10 != "Q")
            {
              
                label_Chk_T1.Text = null;
                label_Chk_T2.Text = null;
                //label_Chk_H1.Text = null;
                //label_Chk_H2.Text = null;


                label_Chk_T1.ForeColor = System.Drawing.Color.Purple;
                label_Chk_T2.ForeColor = System.Drawing.Color.Purple;
                //label_Chk_H1.ForeColor = System.Drawing.Color.SaddleBrown;
                //label_Chk_H2.ForeColor = System.Drawing.Color.SaddleBrown;

                label_Chk_T1.Text = getmedata2.col_6;
                label_Chk_T2.Text = getmedata2.col_5;
                //label_Chk_H1.Text = getmedata2.col_8;
                //label_Chk_H2.Text = getmedata2.col_7;
            }
            else
            {

                label_Chk_T1.ForeColor = System.Drawing.Color.Red;
                label_Chk_T2.ForeColor = System.Drawing.Color.Red;
                //label_Chk_H1.ForeColor = System.Drawing.Color.Red;
                //label_Chk_H2.ForeColor = System.Drawing.Color.Red;
            }

            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.returns();

            if (getmedata4.col_11 != "QT")
            {
                label13.Text = null;
                label13.ForeColor = System.Drawing.Color.Black;
                label13.Text = getmedata4.col_11;

            }
            else
            {
                label13.ForeColor = System.Drawing.Color.Red;
            }
            Data.DataLink getmedata5 = new Data.DataLink();
            getmedata5.instantdel();
            if (getmedata5.col_12 != "QT")
            {
                label22.Text = null;
                label22.ForeColor = System.Drawing.Color.Black;
                label22.Text = getmedata5.col_12;

            }
            else
            {
                label22.ForeColor = System.Drawing.Color.Red;
            }
 
            int dt = (int)Convert.ChangeType(Table1_Count_1.Text, typeof(int));
            int dt2 = (int)Convert.ChangeType(label_Chk_T1.Text, typeof(int));
            //label2.Text = null;
            //label2.Text = (dt + dt2).ToString();

            //int dt3 = (int)Convert.ChangeType(label_Hour1.Text, typeof(int));
            //int dt4 = (int)Convert.ChangeType(label_Chk_H1.Text, typeof(int));
            //label4.Text = null;
            //label4.Text = (dt3 + dt4).ToString();

            int dt5 = (int)Convert.ChangeType(Table2_Count_1.Text, typeof(int));
            int dt6 = (int)Convert.ChangeType(label_Chk_T2.Text, typeof(int));
            //label9.Text = null;
            //label9.Text = (dt5 + dt6).ToString();

            //int dt7 = (int)Convert.ChangeType(label_Hour2.Text, typeof(int));
            //int dt8 = (int)Convert.ChangeType(label_Chk_H2.Text, typeof(int));
            //label11.Text = null;
            //label11.Text = (dt7 + dt8).ToString();
            label5.Text = null;
            label12.Text = null;
            label5.Text = (dt + dt5).ToString();
            label12.Text = (dt + dt5 + dt2 + dt6).ToString();
            tables();
            rungraph();
        }

        private void tables()
        {
           

                 Data.DataLink getmedata = new Data.DataLink();
                 DataSet ds = null;
                try{
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView2.Items.Clear();
                    listView2.Columns.Clear();
                    listView1.View = View.Details;
                    listView2.View = View.Details;
                    listView1.GridLines = true;
                    listView2.GridLines = true;
                    listView1.Sorting = SortOrder.Ascending;
                    listView2.Sorting = SortOrder.Ascending;
                    listView1.Columns.Add("Piece", 100, HorizontalAlignment.Center);
                    listView1.Columns.Add("StockCode", 250, HorizontalAlignment.Center);
                    listView1.Columns.Add("User", 250, HorizontalAlignment.Center);
                    listView2.Columns.Add("Piece", 100, HorizontalAlignment.Center);
                    listView2.Columns.Add("StockCode", 250, HorizontalAlignment.Center);
                    listView2.Columns.Add("User", 250, HorizontalAlignment.Center);
                    ds = getmedata.Tables("TABLE1");
                  //  listBox1.DataSource = ds;
                    //listBox1.


                    foreach (DataRow drow in ds.Tables[0].Rows)
                    //  for (int i = 0; i < row.Count; i++)
                    {
                        // DataRow drow = row[i];

                        // Only rows that have not been deleted
                        if (drow.RowState != DataRowState.Deleted)
                        {
                            // Define the list items
                            ListViewItem lvi = new ListViewItem(drow[0].ToString());
                          lvi.SubItems.Add(drow[1].ToString());
                           lvi.SubItems.Add(drow[2].ToString());
                            listView1.Items.Add(lvi);

                        }

                    }

                }
                catch (Exception e)
                {
                    String Str = e.Message;
                }
                Data.DataLink getmedata2 = new Data.DataLink();
                DataSet ds2 = null;
                try
                {
                    ds2 = getmedata.Tables("TABLE2");
                    //  listBox1.DataSource = ds;
                    //listBox1.


                    foreach (DataRow drow2 in ds2.Tables[0].Rows)
                    //  for (int i = 0; i < row.Count; i++)
                    {
                        // DataRow drow = row[i];

                        // Only rows that have not been deleted
                        if (drow2.RowState != DataRowState.Deleted)
                        {
                            // Define the list items
                            ListViewItem lvi2 = new ListViewItem(drow2[0].ToString());
                           lvi2.SubItems.Add(drow2[1].ToString());
                            lvi2.SubItems.Add(drow2[2].ToString());
                            listView2.Items.Add(lvi2);

                        }

                    }

                }
                catch (Exception e)
                {
                    String Str = e.Message;
                } 
        }
     
        //private void rungraph()
        //{

        
        //    Data.DataLink getmedata = new Data.DataLink();
        //    DataSet ds = null;
        //try{
        //    ds = getmedata.graphing();
        //    if (ds.Tables[0].Rows.Count > 0)
        //    {
        //        DataRow row;
        //        row = ds.Tables[0].Rows[0];
        //        chart1.Series[0].Points.Clear();
        //        chart1.Series[1].Points.Clear();
        //        chart1.Series[2].Points.Clear();
        //        chart1.Series[3].Points.Clear();
        //        foreach (DataRow drow in ds.Tables[0].Rows)
        //        //  for (int i = 0; i < row.Count; i++)
        //        {
        //            // DataRow drow = row[i];

        //            // Only row that have not been deleted
        //            if (drow.RowState != DataRowState.Deleted)
        //            {
        //                //drow["product_code"].ToString()
                  
        //                    // chart1.Series.Add(drow["Procedure"].ToString());
        //                    chart1.Series[drow["Procedure"].ToString()].Points.AddXY(drow["Dateme"].ToString(), drow["Values"].ToString());

        //            }
        //        }
        //       }
        //      }
        //       catch (Exception e)
        //      {
        //        String Str = e.Message;
        //      } 
           
        //    timer1.Enabled = true;
        //    timer1.Start();
        //}
        private void rungraph()
        {


            Data.DataLink getmedata = new Data.DataLink();
            DataSet ds = null;
            try
            {
                ds = getmedata.graphing();
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataRow row;
                    row = ds.Tables[0].Rows[0];
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    chart1.Series[2].Points.Clear();
                    chart1.Series[3].Points.Clear();
                    chart1.Series[0].IsValueShownAsLabel = true;
                    chart1.Series[1].IsValueShownAsLabel = true;
                    chart1.Series[2].IsValueShownAsLabel = true;
                    chart1.Series[3].IsValueShownAsLabel = true;
                    chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;
                    chart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                    foreach (DataRow drow in ds.Tables[0].Rows)
                    //  for (int i = 0; i < row.Count; i++)
                    {
                        // DataRow drow = row[i];

                        // Only row that have not been deleted
                        if (drow.RowState != DataRowState.Deleted)
                        {
                            //drow["product_code"].ToString()

                            // chart1.Series.Add(drow["Procedure"].ToString());
                            chart1.Series[drow["Procedure"].ToString()].Points.AddXY(drow["Dateme"].ToString(), drow["Values"].ToString());

                        }
                    }
                }
            }
            catch (Exception e)
            {
                String Str = e.Message;
            }

            timer1.Enabled = true;
            timer1.Start();
        }

    }
}
