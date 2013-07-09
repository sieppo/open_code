using System;
using System.Data;
using System.Data.Odbc;
using System.Data.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;



namespace Table_Project.Data
{
    public class DataLink
    {
        private OdbcConnection cn;
        private OleDbConnection cns;
        private SqlConnection cnn;
        private SqlConnection cnnpl;

        public string col_1 { get; set; }
        public string col_2 { get; set; }
        public string col_3 { get; set; }
        public string col_4 { get; set; }
        public string col_5 { get; set; }
        public string col_6 { get; set; }
        public string col_7 { get; set; }
        public string col_8 { get; set; }
        public string col_9 { get; set; }
        public string col_10 { get; set; }
        public string col_11 { get; set; }
        public string col_12 { get; set; }
        public string starttime1 { get; set; }
        public string lasttime1 { get; set; }
        public string averagecut1 { get; set; }
        public string starttime2 { get; set; }
        public string lasttime2 { get; set; }
        public string averagecut2 { get; set; }


        public void testingonly()

        {
       
           DataSet ds = null;
           DataSet ds2 = null;
           DataTable table = null;
           DataSet dsa = null;
           DataTable tablea = null;

           try
           {

               cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
               cn.Open();
               int hours1 = DateTime.Now.Hour + 1;
               int minutes = DateTime.Now.Minute;// +1;
               int hours = DateTime.Now.Hour;
               string nowme = hours.ToString() + ":00:00";
               string nowme1 = hours1.ToString() + ":00:00";
               DateTime dt = (DateTime)Convert.ChangeType(nowme, typeof(DateTime));
               DateTime dt1 = (DateTime)Convert.ChangeType(nowme1, typeof(DateTime));
               string sam = DateTime.Today.ToString("dd/MM/yyyy");
              // sam = "03/08/2012";

               String selectStra = "select * from hel_monitor where he_date = '" + sam + "' and ((he_string Like 'A%') OR (he_string Like 'E%')) order by he_time";
               OdbcDataAdapter daa = new OdbcDataAdapter(selectStra, cn);
               dsa = new DataSet();
               daa.Fill(dsa, "hel_monitora");
               //int countme = ds.Tables[0].Rows.Count;
               //col_9 = countme.ToString();
               tablea = dsa.Tables["hel_monitora"];
               // Presuming the DataTable has a column named Date.
               //int hours = DateTime.Now.Hour;
               string expressiona;
               expressiona = "he_tableid = 'Y'";
               DataRow[] foundRowsa;

               // Use the Select method to find all rows matching the filter.
               foundRowsa = tablea.Select(expressiona);
               int perhoura = foundRowsa.Length;
               if (perhoura > 0)
               {

                   //col_1 = perhour.ToString();
                   string first_cut = foundRowsa[0].ItemArray[1].ToString();
                   string last_cut = foundRowsa[foundRowsa.Length - 1].ItemArray[1].ToString();
                   DateTime first_cut_1 = (DateTime)Convert.ChangeType(first_cut, typeof(DateTime));
                   DateTime last_cut_1 = (DateTime)Convert.ChangeType(last_cut, typeof(DateTime));

                   int first_cut_hour = first_cut_1.Hour;
                   int last_cut_hour = last_cut_1.Hour;
                   int first_cut_minute = first_cut_1.Minute;
                   int last_cut_minute = last_cut_1.Minute;
                   int worked_hour = (last_cut_hour - first_cut_hour);
                   int worked_minute = (last_cut_minute - first_cut_minute);
                   int hours_to_minutes = (worked_hour * 60) + worked_minute;
                   double aver_cut = ((double)perhoura / hours_to_minutes) * 57;

                   if (aver_cut < 2147483)
                   {
                       //double aver_cut_round = Math.Round(aver_cut, 0);
                       int convert_cuts = Convert.ToInt32(aver_cut);
                       starttime2 = first_cut;
                       lasttime2 = last_cut;
                       averagecut2 = convert_cuts.ToString();
                   }
                   else
                   {
                       starttime2 = first_cut;
                       lasttime2 = last_cut;
                       averagecut2 = perhoura.ToString();
                   }
                   //double aver_cut_round = Math.Round(aver_cut, 0);
                   //int convert_cuts = Convert.ToInt32(aver_cut);
                   //starttime2 = first_cut;
                   //lasttime2 = last_cut;
                   //averagecut2 = convert_cuts.ToString();
               }
               string expressionxa;
               expressionxa = "he_tableid = 'X'";
               DataRow[] foundRowsxa;

               // Use the Select method to find all rows matching the filter.
               foundRowsxa = tablea.Select(expressionxa);
               int perhourxa = foundRowsxa.Length;
               //col_2 = perhourx.ToString();
               if (perhourxa > 0)
               {

                   string first_cutx = foundRowsxa[0].ItemArray[1].ToString();
                   string last_cutx = foundRowsxa[foundRowsxa.Length - 1].ItemArray[1].ToString();
                   DateTime first_cut_1x = (DateTime)Convert.ChangeType(first_cutx, typeof(DateTime));
                   DateTime last_cut_1x = (DateTime)Convert.ChangeType(last_cutx, typeof(DateTime));

                   int first_cut_hourx = first_cut_1x.Hour;
                   int last_cut_hourx = last_cut_1x.Hour;
                   int first_cut_minutex = first_cut_1x.Minute;
                   int last_cut_minutex = last_cut_1x.Minute;
                   int worked_hourx = (last_cut_hourx - first_cut_hourx);
                   int worked_minutex = (last_cut_minutex - first_cut_minutex);
                   int hours_to_minutesx = (worked_hourx * 60) + worked_minutex;
                   double aver_cutx = ((double)perhourxa / hours_to_minutesx) * 57;
                   if (aver_cutx < 2147483)
                   {
                       //double aver_cut_round = Math.Round(aver_cut, 0);
                       int convert_cutsx = Convert.ToInt32(aver_cutx);
                       starttime1 = first_cutx;
                       lasttime1 = last_cutx;
                       averagecut1 = convert_cutsx.ToString();
                   }
                   else
                   {
                       starttime1 = first_cutx;
                       lasttime1 = last_cutx;
                       averagecut1 = perhourxa.ToString();
                   }
               }

               String selectStr = "select * from hel_monitor where he_date = '" + sam + "' and he_string Like 'A%' order by he_time";
                    OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                    ds = new DataSet();
                    da.Fill(ds, "hel_monitor");
                    int countme = ds.Tables[0].Rows.Count;
                    col_9 = countme.ToString();
                    table = ds.Tables["hel_monitor"];
                    // Presuming the DataTable has a column named Date.
                    //int hours = DateTime.Now.Hour;
                    string expression;
                    expression = "he_tableid = 'Y'";
                    DataRow[] foundRows;

                    // Use the Select method to find all rows matching the filter.
                    foundRows = table.Select(expression);
                    int perhour = foundRows.Length;
                    col_1 = perhour.ToString();
                  
                    string expressionx;
                    expressionx = "he_tableid = 'X'";
                    DataRow[] foundRowsx;

                    // Use the Select method to find all rows matching the filter.
                    foundRowsx = table.Select(expressionx);
                    int perhourx = foundRowsx.Length;
                    col_2 = perhourx.ToString();


                  
               //
                   
                    String selectStrx = "select * from hel_monitor where he_time Between '" + dt.TimeOfDay + "' and '" + dt1.TimeOfDay + "' and he_date = '" + sam + "' and he_string Like 'A%'";
                    OdbcDataAdapter da2 = new OdbcDataAdapter(selectStrx, cn);
                    ds2 = new DataSet();
                    da2.Fill(ds2, "hel_monitor");

                    DataTable table2 = ds2.Tables["hel_monitor"];
                    // Presuming the DataTable has a column named Date.
                    //int hours = DateTime.Now.Hour;
                    string expression2;
                    expression2 = "he_tableid = 'Y'";
                    DataRow[] foundRows2;
               
                    // Use the Select method to find all rows matching the filter.
                    foundRows2 = table2.Select(expression2);
                    int perhour2 = foundRows2.Length;
                    col_3 = perhour2.ToString();

                    string expressionx2;
                    expressionx2 = "he_tableid = 'X'";
                    DataRow[] foundRowsx2;

                    // Use the Select method to find all rows matching the filter.
                    foundRowsx2 = table2.Select(expressionx2);
                    int perhourx2 = foundRowsx2.Length;
                    col_4 = perhourx2.ToString();
                   
            //DataRow row = da.SelectCommand.CommandText = "select he_date from hel_monitor where he_time Like '10%'";
 
            cn.Close();
   
          
             }
            catch (Exception e)
            {
          
                cn.Close();
                String Str = e.Message;
                col_9 = "Q";
            }

          
        }


        


        public void testingonly2()
        {
           
         
            DataSet ds = null;
            DataSet ds2 = null;
            DataTable table = null;
            try
            {

                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                string sam = DateTime.Today.ToString("dd/MM/yyyy");
                // sam = "03/08/2012";
                String selectStr = "select * from hel_monitor where he_date = '" + sam + "' and he_string Like 'E%'";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                ds = new DataSet();
                da.Fill(ds, "hel_monitor");
                int countme = ds.Tables[0].Rows.Count;
                col_10 = countme.ToString();
                table = ds.Tables["hel_monitor"];
                // Presuming the DataTable has a column named Date.
                //int hours = DateTime.Now.Hour;
                string expression;
                expression = "he_tableid = 'Y'";
                DataRow[] foundRows;

                // Use the Select method to find all rows matching the filter.
                foundRows = table.Select(expression);
                int perhour = foundRows.Length;
                col_5 = perhour.ToString();
                
                string expressionx;
                expressionx = "he_tableid = 'X'";
                DataRow[] foundRowsx;

                // Use the Select method to find all rows matching the filter.
                foundRowsx = table.Select(expressionx);
                int perhourx = foundRowsx.Length;
                col_6 = perhourx.ToString();
                

                int hours1 = DateTime.Now.Hour; // +1;
                int hours = DateTime.Now.Hour - 1;
                string nowme = hours.ToString() + ":00:00";
                string nowme1 = hours1.ToString() + ":00:00";
                DateTime dt = (DateTime)Convert.ChangeType(nowme, typeof(DateTime));
                DateTime dt1 = (DateTime)Convert.ChangeType(nowme1, typeof(DateTime));
                String selectStrx = "select * from hel_monitor where he_time Between '" + dt.TimeOfDay + "' and '" + dt1.TimeOfDay + "' and he_date = '" + sam + "' and he_string Like 'E%'";
                OdbcDataAdapter da2 = new OdbcDataAdapter(selectStrx, cn);
                ds2 = new DataSet();
                da2.Fill(ds2, "hel_monitor");

                DataTable table2 = ds2.Tables["hel_monitor"];
                // Presuming the DataTable has a column named Date.
                //int hours = DateTime.Now.Hour;
                string expression2;
                expression2 = "he_tableid = 'Y'";
                DataRow[] foundRows2;

                // Use the Select method to find all rows matching the filter.
                foundRows2 = table2.Select(expression2);
                int perhour2 = foundRows2.Length;
                col_7 = perhour2.ToString();
               
                string expressionx2;
                expressionx2 = "he_tableid = 'X'";
                DataRow[] foundRowsx2;

                // Use the Select method to find all rows matching the filter.
                foundRowsx2 = table2.Select(expressionx2);
                int perhourx2 = foundRowsx2.Length;
                col_8 = perhourx2.ToString();
                
                //DataRow row = da.SelectCommand.CommandText = "select he_date from hel_monitor where he_time Like '10%'";
               
                cn.Close();
            }
            catch (Exception e)
            {
            
                cn.Close();
                String Str = e.Message;
                col_10 = "Q";
            }

      
        }

        public void returns()
        {

            DataSet ds = null;
          
            try
            {

                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                string sam = DateTime.Today.ToString("dd/MM/yyyy");
                // sam = "03/08/2012";
                String selectStr = "select pa_tranno from pieceaudit where pa_trandate = '" + sam + "' and pa_narrative LIKE 'RETURN CREATED%' and pa_location = 'BIN17'";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                ds = new DataSet();
                da.Fill(ds, "pieceaudit");
                int countme = ds.Tables[0].Rows.Count;
                col_11 = countme.ToString();
                cn.Close();
            }
            catch (Exception e)
            {
                cn.Close();
                String Str = e.Message;
                col_11 = "QT";
            }

       
        }

        //public void instantdel()
        //{

        //    DataSet ds = null;

        //    try
        //    {

        //        cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
        //        cn.Open();
        //        string sam = DateTime.Today.ToString("dd/MM/yyyy");
        //        // sam = "03/08/2012";
        //        String selectStr = "select pa_tranno from pieceaudit where pa_trandate = '" + sam + "' AND pa_location = 'HJ01A' AND pa_warehouse IN('W', 'C', 'G', '2') WITH (ROWLOCK, UPDLOCK, READPAST)";
        //        OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
        //        ds = new DataSet();
        //        da.Fill(ds, "pieceaudit");
        //        int countme = ds.Tables[0].Rows.Count;
        //        col_12 = countme.ToString();
        //        cn.Close();
        //    }
        //    catch (Exception e)
        //    {
        //        cn.Close();
        //        String Str = e.Message;
        //        col_12 = "QT";
        //    }


        //}
        public void instantdel()
        {

            DataSet ds = null;

            try
            {

                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                string sam = DateTime.Today.ToString("dd/MM/yyyy");
                // sam = "03/08/2012";
                String selectStr = "select pa_tranno from pieceaudit where pa_trandate = '" + sam + "' AND pa_location = 'HJ01A' AND pa_warehouse IN('W', 'C', 'G', '2')";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                ds = new DataSet();
                da.Fill(ds, "pieceaudit");
                int countme = ds.Tables[0].Rows.Count;
                col_12 = countme.ToString();
                cn.Close();
            }
            catch (Exception e)
            {
                cn.Close();
                String Str = e.Message;
                col_12 = "QT";
            }


        }

        public DataSet graphing()
        {
            DataSet ds = null;
            DataSet dsE = null;
            DataSet dsreturn = null;
            DataTable dtreturn = null;
            DataTable tableE = null;
            DataTable table = null;
            try
            {
                dsreturn = new DataSet();
                dtreturn = new DataTable();
                DataColumn column;
                DataRow row;
                column = dtreturn.Columns.Add();
                column.ColumnName = "Procedure";
                column.DataType = typeof(string);

                //column = dtreturn.Columns.Add();
                //column.ColumnName = "Table2cut";
                //column.DataType = typeof(string);

                //column = dtreturn.Columns.Add();
                //column.ColumnName = "Table1chk";
                column.DataType = typeof(string);

                column = dtreturn.Columns.Add();
                column.ColumnName = "Values";
                column.DataType = typeof(string);

                column = dtreturn.Columns.Add();
                column.ColumnName = "Dateme";
                column.DataType = typeof(string);

                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                string sam = DateTime.Today.ToString("dd/MM/yyyy");

                string sam8 = DateTime.Today.AddDays(-8).ToString("dd/MM/yyyy");
                String selectStrx = "select he_tableid, he_date from hel_monitor where he_date Between '" + sam8 + "' and '" + sam + "' and he_string Like 'A%'";
                String selectStrxE = "select he_tableid, he_date from hel_monitor where he_date Between '" + sam8 + "' and '" + sam + "' and he_string Like 'E%'";

                OdbcDataAdapter da = new OdbcDataAdapter(selectStrx, cn);

                ds = new DataSet();
                da.Fill(ds, "hel_monitor");
                table = ds.Tables["hel_monitor"];
                OdbcDataAdapter daE = new OdbcDataAdapter(selectStrxE, cn);
                dsE = new DataSet();
                daE.Fill(dsE, "hel_monitorE");
                tableE = dsE.Tables["hel_monitorE"];

                for (int i = 8; i > -1; i--)
                {

                    string sam1 = DateTime.Today.AddDays(-i).ToString("dd/MM/yyyy");
                    string expression = "he_tableid = 'X' and he_date = '" + sam1 + "'";
                    DataRow[] foundRows;
                    foundRows = table.Select(expression);
                    string toatals = foundRows.Length.ToString();
                    row = dtreturn.NewRow();
                    row["Dateme"] = sam1;
                    row["Values"] = toatals;
                    row["Procedure"] = "Table 1 Cuts";
                    dtreturn.Rows.Add(row);

                    string expression2 = "he_tableid = 'Y' and he_date = '" + sam1 + "'";
                    DataRow[] foundRows2;
                    foundRows2 = table.Select(expression2);
                    string toatals2 = foundRows2.Length.ToString();
                    row = dtreturn.NewRow();
                    row["Dateme"] = sam1;
                    row["Values"] = toatals2;
                    row["Procedure"] = "Table 2 Cuts";
                    dtreturn.Rows.Add(row);

                    string expression3 = "he_tableid = 'X' and he_date = '" + sam1 + "'";
                    DataRow[] foundRows3;
                    foundRows3 = tableE.Select(expression3);
                    string toatals3 = foundRows3.Length.ToString();
                    row = dtreturn.NewRow();
                    row["Dateme"] = sam1;
                    row["Values"] = toatals3;
                    row["Procedure"] = "Table 1 Checks";
                    dtreturn.Rows.Add(row);

                    string expression4 = "he_tableid = 'Y' and he_date = '" + sam1 + "'";
                    DataRow[] foundRows4;
                    foundRows4 = tableE.Select(expression4);
                    string toatals4 = foundRows4.Length.ToString();
                    row = dtreturn.NewRow();
                    row["Dateme"] = sam1;
                    row["Values"] = toatals4;
                    row["Procedure"] = "Table 2 Checks";
                    dtreturn.Rows.Add(row);

                }



                dsreturn.Tables.Add(dtreturn);
                dsreturn.AcceptChanges();

                cn.Close();
            }
            catch (Exception e)
            {

                cn.Close();
                String Str = e.Message;
            }

            return dsreturn;
            
        }
        //for twitter connection
        private void OpenCnn()
        {

            String cnnStr = "Data Source=192.168.20.5;Initial Catalog=Realitex;Persist Security Info=True;User ID=Price_List2;Password=Price_List2010";
         
            cnn = new SqlConnection(cnnStr);
            cnn.Open();
        }
        private void CloseCnn()
        {
            // 5- step five
            cnn.Close();
        }


        //for price list connection

        private void OpenCnnPL()
        {

            String cnnStr = "Data Source=192.168.20.5;Initial Catalog=View_Price_List;Persist Security Info=True;User ID=Price_List2;Password=Price_List2010";
  
            cnnpl = new SqlConnection(cnnStr);
            cnnpl.Open();
        }
        private void CloseCnnPL()
        {
            // 5- step five
            cnnpl.Close();
        }


        private string cust { get; set; }
        private string rep { get; set; }
        private void findcustandadd(string arg1)
        {
            cust = "";
            rep = "";
            DataSet ds = null;
            try
            {

                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                String selectStr = "select oh_custaccref, oh_repref from orderhead, ordersub where os_orderno = oh_orderno and os_invno = '" + arg1 + "' Group By oh_custaccref, oh_repref";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                ds = new DataSet();
                da.Fill(ds, "customer");
                int countme = ds.Tables[0].Rows.Count;
                DataRow dd = ds.Tables[0].Rows[0];
                cust = dd[0].ToString();
                rep = dd[1].ToString();
                cn.Close();
            }
            catch (Exception e)
            {
                cn.Close();
                //cns.Close();
                String Str = e.Message;
              
            }
            cn.Close();


        }

        public void testing(String tablenaming)
        {
            DataSet ds = null;
            DataTable table = null;
            DataTable myTable = null;
            // string tablenaming = "tables_realitex";
            try
            {
                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.ConnectionTimeout = 2140000000;
                cn.Open();
                String selectStr = "select * from " + tablenaming;

                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                da.SelectCommand.CommandTimeout = 2140000000;
                //ds.Rows.Cast<System.Data.DataRow>().Take(n);
                ds = new DataSet();
                da.Fill(ds, 1, 15, tablenaming);
                table = ds.Tables[tablenaming];

                myTable = table.Clone();
                cn.Close();
            }
            catch (Exception err)
            {

                cn.Close();

                myTable.Dispose();
                table.Dispose();
                String Str = err.Message;

                error_messages("testing_P1", Str, tablenaming);

            }
            try
            {
                OpenCnnsql();
                SqlTableCreater ssss = new SqlTableCreater();
                ssss.DestinationTableName = tablenaming;
                messages2("request made for " + tablenaming);
                ssss.Connection = cnnsql;
                ssss.CreateFromDataTable(myTable);
                //messages2("send to bulk upload " + tablenaming);
                //BulkInsertDataTableSQL(tablenaming, table);
                //ssss.Transaction.Commit();
                CloseCnnsql();



            }
            catch (Exception err)
            {
                CloseCnn();


                myTable.Dispose();
                table.Dispose();
                String Str = err.Message;

                error_messages("testing_P2", Str, tablenaming);

            }


            myTable.Dispose();
            table.Dispose();

        }

        public DataTable GetmeTables()
        {

            DataTable table = null;
            string tablenames = null;
            try
            {
                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.ConnectionTimeout = 2140000000;

                cn.Open();
                table = cn.GetSchema("Tables");


                cn.Close();

            }
            catch (Exception err)
            {
                String Str = err.Message;

                error_messages("GetmeTables", Str, tablenames);
                cn.Close();
                //  Console.WriteLine("!!!ERROR!!!- " + tablenames);
                //    MessageBox.Show(err.Message.ToString());
            }
            return table;
        }
        //new info
        private DataSet fileinfo(String file, String sheet)
        {
            DataSet ds = null;
            DataTable table = null;
          
            try
            {
                string excelConnectString = "Provider = Microsoft.Jet.OLEDB.4.0;" +
                "Data Source = " + file + ";" +
                "Extended Properties = Excel 8.0;";

                OleDbConnection objConn = new OleDbConnection(excelConnectString);
                table = objConn.GetSchema("Tables");
                OleDbCommand objCmd = new OleDbCommand("Select * From [" + sheet + "$]", objConn);

                OleDbDataAdapter objDatAdap = new OleDbDataAdapter();
                objDatAdap.SelectCommand = objCmd;
                ds = new DataSet();
                objDatAdap.Fill(ds);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }

            return ds;
        }


        public void BulkInsertDataTable(string tableName, DataTable table)
        {
            String connectionString = "Data Source=192.168.20.5;Initial Catalog=Realitex;Persist Security Info=True;User ID=Price_List2;Password=Price_List2010";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlBulkCopy bulkCopy =
                        new SqlBulkCopy
                        (
                        connection,
                        SqlBulkCopyOptions.TableLock |
                        SqlBulkCopyOptions.FireTriggers |
                        SqlBulkCopyOptions.UseInternalTransaction,
                        null
                        );

                    bulkCopy.DestinationTableName = tableName;
                    bulkCopy.BulkCopyTimeout = 2140000000;
                    connection.Open();


                    bulkCopy.WriteToServer(table);
                    connection.Close();
                }
            }
            catch (Exception err)
            {
                String Str = err.Message;

                error_messages("BulkInsertDataTable", Str, tableName);
                //    MessageBox.Show(err.Message.ToString());
            }
        }

        public void julie()
        {
            DataSet dsreturn = null;
            DataTable dtreturn = null;
            dsreturn = new DataSet();
            dtreturn = new DataTable();
            DataColumn column;
            DataRow row1;
            column = dtreturn.Columns.Add();
            column.ColumnName = "invoice_info";
            column.DataType = typeof(string);

            column = dtreturn.Columns.Add();
            column.ColumnName = "invoice_number";
            column.DataType = typeof(string);

            column = dtreturn.Columns.Add();
            column.ColumnName = "total_amount";
            column.DataType = typeof(string);

            column = dtreturn.Columns.Add();
            column.ColumnName = "to_pay";
            column.DataType = typeof(string);

            column = dtreturn.Columns.Add();
            column.ColumnName = "customer";
            column.DataType = typeof(string);

            column = dtreturn.Columns.Add();
            column.ColumnName = "rep";
            column.DataType = typeof(string);

            try
            {

                DataSet ds = null;

                cns = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='d:\\paid.xls';Extended Properties=Excel 8.0;");
                cns.Open();
                String selectStrnew =  "select * from [PAYABLE$]";

                OleDbDataAdapter da = new OleDbDataAdapter(selectStrnew, cns);

                ds = new DataSet();
                da.Fill(ds, "monitor");

                DataRow row;
                row = ds.Tables[0].Rows[0];
               
                foreach (DataRow drow in ds.Tables[0].Rows)
                //  for (int i = 0; i < row.Count; i++)
                {
                    // DataRow drow = row[i];

                    // Only row that have not been deleted
                    if (drow.RowState != DataRowState.Deleted)
                    {
                        //drow["product_code"].ToString()
                        try
                        {
                           string help1 = drow[0].ToString();
                           string help2 = Decimal.Round(Convert.ToInt32(drow[1].ToString()), 0).ToString();
                           string help3 = drow[2].ToString();
                           string help4 = drow[3].ToString();
                           string help5 = drow[4].ToString();

                           if (help1 != "")
                           {
                               findcustandadd(help2);
                                
                                
                                  row1 = dtreturn.NewRow();
                                  row1["invoice_info"] = help1;
                                  row1["invoice_number"] = help2;
                                  row1["total_amount"] = help3;
                                  row1["to_pay"] = help4;
                                  row1["customer"] = cust;
                                  row1["rep"] = rep;
                                  dtreturn.Rows.Add(row1);

              


                
                           }
                               
                        }
                        catch (Exception e)
                        {
                            String Str = e.Message;
                            cns.Close();
                        }


                    }
                }

               
                dsreturn.Tables.Add(dtreturn);
                dsreturn.AcceptChanges();
                ExcelLibrary.DataSetHelper.CreateWorkbook("payments.xls", dsreturn); 
                cns.Close();
            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());

            }

        }


        public DataSet Tables(string TT)
        {

            DataSet ds = null;

            try
            {

                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                string sam = DateTime.Today.ToString("dd/MM/yyyy");
                // sam = "03/08/2012";

                String selectStr = "SELECT DISTINCT pa_pieceno, pa_stockref, pa_operator from pieceaudit, stockpiece where (pa_location = sp_location and pa_pieceno = sp_pieceno) and pa_trandate = '" + sam + "' and pa_location = '" + TT + "' and pa_reasonref = '000' and NOT((sp_remain - sp_allocated) < '1.5' and sp_datechecked = '" + sam + "')";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);
                ds = new DataSet();
                da.Fill(ds, "pieceaudit");
                //int countme = ds.Tables[0].Rows.Count;
                //col_11 = countme.ToString();
                cn.Close();
            }
            catch (Exception e)
            {
                cn.Close();
                String Str = e.Message;
             
            }
            return ds;

        }

        private void error_messages(string where, string messages, string what)
        {

            //string[] s = System.Web.HttpContext.Current.Request.Url.AbsolutePath.Split(new char[] { '/' });
            //string PageName = s[s.Length - 1];
            //s = System.Web.HttpContext.Current.Request.MapPath(PageName).Split(new char[] { '\\' });
            //string path = s[0] + "\\";
            //for (int i = 1; i < s.Length - 1; i++)
            //{
            //    path = path + s[i] + "\\";
            //}

            DateTime time = DateTime.Now;             // Use current time
            string format = "d M yyyy - h:mm";            // Use this format
            string thetime = time.ToString(format);

            string maakee = thetime + "  - Prosess is: " + where + " " + what + "  -  Error is: " + messages;

            StreamWriter sw = null;

            try
            {

                sw = new StreamWriter("C:\\log.txt", true);

                sw.WriteLine(maakee);
                sw.Flush();

            }
            catch (Exception ex)
            {
                messages2("own  --  error class");
            }
            finally
            {
                if (sw != null)
                {
                    sw.Dispose();
                    sw.Close();
                }
            }


        }
        private void messages2(string messages)
        {


            DateTime time = DateTime.Now;             // Use current time
            string format = "d M yyyy - h:mm";            // Use this format
            string thetime = time.ToString(format);

            string maakee = thetime + "   -   " + messages;

            StreamWriter sw = null;

            try
            {

                sw = new StreamWriter("C:\\log2.txt", true);

                sw.WriteLine(maakee);
                sw.Flush();

            }
            catch (Exception ex)
            {
                error_messages("own", ex.ToString(), "error class");
            }
            finally
            {
                if (sw != null)
                {
                    sw.Dispose();
                    sw.Close();
                }
            }


        }

        private void messageslog(string messages)
        {


            DateTime time = DateTime.Now;             // Use current time
            string format = "d M yyyy - h:mm";            // Use this format
            string thetime = time.ToString(format);

            string maakee = thetime + "   -   " + messages;

            StreamWriter sw = null;

            try
            {

                sw = new StreamWriter("C:\\log3.txt", true);

                sw.WriteLine(maakee);
                sw.Flush();

            }
            catch (Exception ex)
            {
                error_messages("own", ex.ToString(), "error class");
            }
            finally
            {
                if (sw != null)
                {
                    sw.Dispose();
                    sw.Close();
                }
            }


        }
        //Update sql for new company

        //first to sort out customers
        //connection to sql
        private SqlConnection cnnsql;

        private void OpenCnnsql()
        {
            String cnnStr = "Data Source=192.168.20.5;Initial Catalog=twitter;Persist Security Info=True;User ID=Price_List2;Password=Price_List2010";
            cnnsql = new SqlConnection(cnnStr);
            cnnsql.Open();
        }
        private void CloseCnnsql()
        {
            cnnsql.Close();
        }

        //this to repair cockup
        public void matchpieceno10()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM pieceno2";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pieceno2");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {

                if (drow3.RowState != DataRowState.Deleted)
                {



                    try
                    {


                        String selectStrR = "Update orderline Set ol_stockref = '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "', ol_stockdesc = '" + drow3["newdescription"].ToString().Replace("'", "''").Trim() + "' where ol_stockref = '" + drow3["oldcode"].ToString().Trim() +
                            "' and ol_linestatus >= '30' and ol_linestatus < '60'";



                        OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                        cmd2.ExecuteNonQuery();

                    }
                    catch (Exception e)
                    {
                        messageslog(drow3["newcode"].ToString().Replace("'", "''").Trim() + "  " + drow3["oldcode"].ToString().Trim());
                        //cn.Close();
                        //CloseCnnsql();
                        //MessageBox.Show(e.Message.ToString());
                    }

                    //try
                    //{


                    //    String selectStrR2 = "Update ordersub Set os_stockref = '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "', os_stockdesc = '" + drow3["newdescription"].ToString().Replace("'", "''").Trim() + "' where ol_stockref = '" + drow3["oldcode"].ToString().Trim() +
                    //        "' and os_substatus >= '30' and os_substatus < '60'";



                    //    OdbcCommand cmd22 = new OdbcCommand(selectStrR2, cn);
                    //    cmd22.ExecuteNonQuery();

                    //}
                    //catch (Exception e)
                    //{
                    //    cn.Close();
                    //    CloseCnnsql();
                    //    MessageBox.Show(e.Message.ToString());
                    //}



                }
            }

            cn.Close();
            CloseCnnsql();
        }

        //this is to repaire alias
        public void matchpieceno12()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM pieceno3";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pieceno2");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {

                if (drow3.RowState != DataRowState.Deleted)
                {



                    try
                    {


                        String selectStrR = "Update stockalias Set sa_stockref = '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "' where sa_stockref = '" + drow3["oldcode"].ToString().Trim() + "'";



                        OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                        cmd2.ExecuteNonQuery();

                    }
                    catch (Exception e)
                    {
                        cn.Close();
                        CloseCnnsql();
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }

            cn.Close();
            CloseCnnsql();
        }

        //this is to update / transfer new stock codes to pieces 
        public void matchpieceno2()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM pieceno3";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pieceno2");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                
                if (drow3.RowState != DataRowState.Deleted)
                {



                    try
                    {


                        String selectStrR = "Update stockpiece Set sp_stockref = '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "' where sp_stockref = '" + drow3["oldcode"].ToString().Trim() + "'";



                        OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                        cmd2.ExecuteNonQuery();

                    }
                    catch (Exception e)
                    {
                        cn.Close();
                        CloseCnnsql();
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }
          
            cn.Close();
            CloseCnnsql();
        }
        
       //This updates the pricing system for the sql side 
        public void matchpieceno3()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();
                OpenCnnPL();
                String selectStr = "SELECT * FROM pieceno3";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pieceno2");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
               CloseCnnPL();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            //cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            //cn.Open();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {

                if (drow3.RowState != DataRowState.Deleted)
                {



                    try
                    {


                        String selectStrR = "Update Price_Lists Set PL_Stock_Code = '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "', PL_Stock_Description = '" + drow3["newdescription"].ToString().Replace("'", "''").Trim() + "' where PL_Stock_Code = '" + drow3["oldcode"].ToString().Trim() + "'";



                        SqlCommand cmd2 = new SqlCommand(selectStrR, cnnpl);
                        cmd2.ExecuteNonQuery();

                    }
                    catch (Exception e)
                    {
                        CloseCnnPL();
                        CloseCnnsql();
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }

            CloseCnnPL();
            CloseCnnsql();
        }

        public void matchpieceno5()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();
                OpenCnnPL();
                String selectStr = "SELECT * FROM pieceno3";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pieceno2");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                CloseCnnPL();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            //cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            //cn.Open();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {

                if (drow3.RowState != DataRowState.Deleted)
                {



                    try
                    {


                        String selectStrR = "Update Build_Realitex_Table Set Stock_Code = '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "', Stock_Description = '" + drow3["newdescription"].ToString().Replace("'", "''").Trim() + "' where Stock_Code = '" + drow3["oldcode"].ToString().Trim() + "'";



                        SqlCommand cmd2 = new SqlCommand(selectStrR, cnnpl);
                        cmd2.ExecuteNonQuery();

                    }
                    catch (Exception e)
                    {
                        CloseCnnPL();
                        CloseCnnsql();
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }

            CloseCnnPL();
            CloseCnnsql();
        }
        //This updates the pricing system for the realitex side 
        public void matchpieceno4()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM pieceno2";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pieceno2");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                 
                if (drow3.RowState != DataRowState.Deleted)
                {
                    DataSet ds22 = null;
                    DataRow row22 = null;
                        try
                        {
                            String selectStr22 = "SELECT * FROM oplistm WHERE product_code = '" + drow3["oldcode"].ToString().Replace("'", "''").Trim() + "'";
                            OdbcDataAdapter da22 = new OdbcDataAdapter(selectStr22, cn);
                            ds22 = new DataSet();
                            da22.Fill(ds22, "oplistm");
                        }
                        catch (Exception e)
                        {
                           // messageslog(drow3["newcode"].ToString().Replace("'", "''").Trim() + "  " + drow3["oldcode"].ToString().Trim() + "  " + e.Message.ToString());
                            //cn.Close();
                            //CloseCnnsql();
                            //MessageBox.Show(e.Message.ToString());
                        }
                        if (ds22.Tables[0].Rows.Count > 0)

                        {
                      
                            row22 = ds22.Tables[0].Rows[0];

                            foreach (DataRow drow22 in ds22.Tables[0].Rows)
                            {
                                if (RecordExists2("SELECT price_list FROM oplistm WHERE price_list = '" + drow22["price_list"].ToString().Replace("'", "''").Trim() + "' and product_code = '" + drow3["oldcode"].ToString().Replace("'", "''").Trim() + "'"))
                               
                                {
                                try
                                {

                                    String selectStrR2 = "Delete From oplistm where product_code = '" + drow3["oldcode"].ToString().Trim() + "' and price_list = '" + drow22["price_list"].ToString().Replace("'", "''").Trim() + "'";
                                    OdbcCommand cmd22 = new OdbcCommand(selectStrR2, cn);
                                    cmd22.ExecuteNonQuery();

                                    String selectStrR = "INSERT INTO oplistm (price_list, product_code, sequence_number, description, price, new_price, vat_inclusive_flag, unit_qty_per_price, customer_code, currency_code, unit_code, unit_code_group)" +
                                           "VALUES ('" + drow22["price_list"].ToString().Replace("'", "''").Trim() + "', '" + drow3["newcode"].ToString().Replace("'", "''").Trim() + "', '1', '" + drow22["description"].ToString().Replace("'", "''").Trim() +
                                           "', '" + drow22["price"].ToString().Replace("'", "''").Trim() + "', '" + drow22["new_price"].ToString().Replace("'", "''").Trim() + "', 'N', '0', '', '', '', '')";
                                    OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                                    cmd2.ExecuteNonQuery();
                                }
                                catch (Exception e)
                                {
                                    messageslog(drow3["newcode"].ToString().Replace("'", "''").Trim() + "  " + drow3["oldcode"].ToString().Trim() + "  " + drow22["price_list"].ToString().Replace("'", "''").Trim());
                                    //cn.Close();
                                    //CloseCnnsql();
                                    //MessageBox.Show(e.Message.ToString());
                                }
                                }
                                else
                                { }

                            }
                            
                          
                        }
                       
                    }
                   
                    
                }
            

            cn.Close();
            CloseCnnsql();
        }
        //stock price
        public void match3()
        {

            DataSet ds = null;

            try
            {
                //OpenCnnsql();
                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                String selectStr = "SELECT st_stockref, st_cutprice1, st_rollprice1 FROM stock where st_warehouse = 'T'";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);

                ds = new DataSet();
                da.Fill(ds, "stock");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                cn.Close();
                //CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

             DataRow row3;
             //cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
             //cn.Open();
             OpenCnnsql();
             //OpenCnnPL();
            row3 = ds.Tables[0].Rows[0];

            foreach (DataRow drow in ds.Tables[0].Rows)
            {
                //string amount1 = "";

                if (drow.RowState != DataRowState.Deleted)
                {
                   
                        //if (Convert.ToDecimal(drow3["Terms"]) < 1)
                        //{
                        //    //amount1 = (Convert.ToDecimal(Percenteage)).ToString()* Convert.ToDecimal(Percenteage)).ToString("#0.00");
                        //    decimal amount = Convert.ToDecimal(drow3["Terms"]) * 100;
                            try
                            {
                                //String selectStr = "Update Customer_1 Set Cust_Price_List = '" + drow3["NEW"] + "' where Cus_ID = '" + drow3["new_account"].ToString().Trim() + "'";

      //                          String selectStr = "INSERT INTO Price_Lists (PL_Name, PL_Description, PL_Price1, PL_Price2, PL_Stock_Code, PL_Stock_Description, PL_Design_Code, PL_Design_Description, PL_Stock_Type)" +
      //"VALUES ('" + drow["NEW"].ToString().Replace("'", "''") + "', '" + drow["description"].ToString().Replace("'", "''") + "', '" + drow["price"].ToString().Replace("'", "''") + "', '" + drow["new_price"].ToString().Replace("'", "''") +
      //"', '" + drow["product_code"].ToString().Replace("'", "''") + "', '" + drow["st_stockdesc"].ToString().Replace("'", "''") + "', '" + drow["st_designref"].ToString().Replace("'", "''") + "', '" + drow["de_designdesc"].ToString().Replace("'", "''") + "', '5')";


      //                          SqlCommand cmd = new SqlCommand(selectStr, cnnpl);
      //                          cmd.ExecuteNonQuery();
                                //realitex
                                String selectStr2 = "Update tempdes Set cutprice1 = '" + drow["st_cutprice1"] + "', rollprice1 = '" + drow["st_rollprice1"] + "' where stock_code = '" + drow["st_stockref"].ToString().Trim() + "'";




                                //String selectStr2 = "INSERT INTO tempdes (stock_code, stockdescrip, design_code)" +
                                //              "VALUES ('" + drow["st_stockref"].ToString().Replace("'", "''") + "', '" + drow["st_stockdesc"].ToString().Replace("'", "''") + "', '" + drow["st_designref"].ToString().Replace("'", "''") + "')";






                                SqlCommand cmd2 = new SqlCommand(selectStr2, cnnsql);
                                cmd2.ExecuteNonQuery();
                            }
                            catch (Exception e)
                            {
                                cn.Close();
                                CloseCnnPL();
                                //CloseCnnsql();
                                MessageBox.Show(e.Message.ToString());
                            }
                        //}
                   

                }
            }
            CloseCnnsql();
            //CloseCnnPL();
            //cn.Close();
            //CloseCnnsql();
        }

        //priclists
        public void match()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();
                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();
                OpenCnnPL();
                String selectStr = "SELECT * FROM extras";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "extras");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnPL();
                cn.Close();
               CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            //cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            //cn.Open();
            //OpenCnnPL();
            row3 = ds.Tables[0].Rows[0];
             String today = DateTime.Today.ToString("dd/MM/yyyy");   
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                //string amount1 = "";
               

                if (drow3.RowState != DataRowState.Deleted)
                {


                    String selectStr = "SELECT * FROM stock where st_designref = '" + drow3["design"].ToString().Replace("'", "''").Trim() + "'";
                    OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);

                    ds = new DataSet();
                    da.Fill(ds, "stock");
                    DataRow row;
                    if (ds.Tables[0].Rows.Count > 0)
                     {
                    row = ds.Tables[0].Rows[0];

                    foreach (DataRow drow in ds.Tables[0].Rows)
                    {
                        string cutprice = Convert.ToDecimal(drow3["price1"]).ToString("#0.00");
                        string rollprice = Convert.ToDecimal(drow3["price2"]).ToString("#0.00");
                        string designcodeing;
                        //if (drow3["price1"].ToString() == "0.00")
                        //{
                        //    cutprice = drow["cutprice1"].ToString();
                        //}
                        //else
                        //{
                        //    cutprice = drow3["CUTS"].ToString();
                        //}

                        //if (drow3["ROLLS"].ToString() == "0.00")
                        //{
                        //    rollprice = drow["rollprice1"].ToString();
                        //}
                        //else
                        //{
                        //    rollprice = drow3["ROLLS"].ToString();
                        //}
                        if (drow3["design"].ToString().Replace("'", "''").Trim() == "EMT3")
                        { designcodeing = "EMPIRE TWIST 30"; }
                        else if (drow3["design"].ToString().Replace("'", "''").Trim() == "EMT4")
                        { designcodeing = "EMPIRE TWIST 40"; }
                        else
                        { designcodeing = "EMPIRE TWIST 50"; }


                        if (drow.RowState != DataRowState.Deleted)
                        {

                            try
                            {
                                //String selectStr = "Update Customer_1 Set Cust_Price_List = '" + drow3["NEW"] + "' where Cus_ID = '" + drow3["new_account"].ToString().Trim() + "'";
                                if (RecordExists("SELECT PL_Name FROM Price_Lists WHERE PL_Name = '" + drow3["code"].ToString().Replace("'", "''").Trim() + "' and PL_Stock_Code = '" + drow["st_stockref"].ToString().Replace("'", "''").Trim() + "'"))
                                { }
                                else
                                {
                                    String selectStrPL = "INSERT INTO Price_Lists (PL_Name, PL_Description, PL_Price1, PL_Price2, PL_Stock_Code, PL_Stock_Description, PL_Design_Code, PL_Design_Description, PL_Stock_Type, PL_Sequence, PL_Update, PL_Update_Date)" +
            "VALUES ('" + drow3["code"].ToString().Replace("'", "''").Trim() + "', '" + drow3["description"].ToString().Replace("'", "''").Trim() + "', '" + cutprice + "', '" + rollprice +
            "', '" + drow["st_stockref"].ToString().Replace("'", "''").Trim() + "', '" + drow["st_stockdesc"].ToString().Replace("'", "''").Trim() + "', '" + drow["st_designref"].ToString().Replace("'", "''").Trim() + "', '" + designcodeing + "', '5', '1', '1', '" + today + "')";


                                    SqlCommand cmd = new SqlCommand(selectStrPL, cnnpl);
                                    cmd.ExecuteNonQuery();
                                    //realitex
                                    //String selectStr2 = "Update slcustm Set price_list = '" + drow3["NEW"] + "' where customer = '" + drow3["new_account"].ToString().Trim() + "'";

                                }
                                if (RecordExists2("SELECT price_list FROM oplistm WHERE price_list = '" + drow3["code"].ToString().Replace("'", "''").Trim() + "' and product_code = '" + drow["st_stockref"].ToString().Replace("'", "''").Trim() + "'"))
                                { }
                                else
                                {

                                    String selectStrR = "INSERT INTO oplistm (price_list, product_code, sequence_number, description, price, new_price, customer_code, currency_code, unit_code, unit_code_group, vat_inclusive_flag, unit_qty_per_price)" +
                                                  "VALUES ('" + drow3["code"].ToString().Replace("'", "''").Trim() + "', '" + drow["st_stockref"].ToString().Replace("'", "''").Trim() + "', '1', '', '" + cutprice + "', '" + rollprice + "', '', '', '', '', 'N', '0')";






                                    OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                                    cmd2.ExecuteNonQuery();
                                }
                            }
                            catch (Exception e)
                            {
                                cn.Close();
                                CloseCnnPL();
                                CloseCnnsql();
                                MessageBox.Show(e.Message.ToString());
                            }
                        }
                    }
                 }
                    //}


                }
            }
            CloseCnnPL();
            cn.Close();
            CloseCnnsql();
        }

        //pdate print matrix 
       public void match8()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM docs";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "docs");

                CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            OpenCnnPL();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                //string amount1 = "";


                if (drow3.RowState != DataRowState.Deleted)
                {


                    String selectStr = "SELECT customer FROM slcustm where customer Like 'T%'";
                    OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);

                    ds = new DataSet();
                    da.Fill(ds, "slcustm");
                    DataRow row;
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        row = ds.Tables[0].Rows[0];

                        foreach (DataRow drow in ds.Tables[0].Rows)
                        {
                            //string cutprice;
                            //string rollprice;
                            //if (drow3["CUTS"].ToString() == "0.00")
                            //{
                            //    cutprice = drow["cutprice1"].ToString();
                            //}
                            //else
                            //{
                            //    cutprice = drow3["CUTS"].ToString();
                            //}

                            //if (drow3["ROLLS"].ToString() == "0.00")
                            //{
                            //    rollprice = drow["rollprice1"].ToString();
                            //}
                            //else
                            //{
                            //    rollprice = drow3["ROLLS"].ToString();
                            //}

                            if (drow.RowState != DataRowState.Deleted)
                            {

                                try
                                {
                //                    //String selectStr = "Update Customer_1 Set Cust_Price_List = '" + drow3["NEW"] + "' where Cus_ID = '" + drow3["new_account"].ToString().Trim() + "'";
                //                    if (RecordExists("SELECT PL_Name FROM Price_Lists WHERE PL_Name = '" + drow3["NEW"].ToString().Replace("'", "''") + "' and PL_Stock_Code = '" + drow["stock_code"].ToString().Replace("'", "''") + "'"))
                //                    { }
                //                    else
                //                    {
                //                        String selectStrPL = "INSERT INTO Price_Lists (PL_Name, PL_Description, PL_Price1, PL_Price2, PL_Stock_Code, PL_Stock_Description, PL_Design_Code, PL_Design_Description, PL_Stock_Type, PL_Sequence, PL_Update, PL_Update_Date)" +
                //"VALUES ('" + drow3["NEW"].ToString().Replace("'", "''") + "', '', '" + cutprice + "', '" + rollprice +
                //"', '" + drow["stock_code"].ToString().Replace("'", "''") + "', '" + drow["stockdescrip"].ToString().Replace("'", "''") + "', '" + drow["design_code"].ToString().Replace("'", "''") + "', '" + drow["design_descrip"].ToString().Replace("'", "''") + "', '5', '1', '1', '" + today + "')";


                //                        SqlCommand cmd = new SqlCommand(selectStrPL, cnnpl);
                //                        cmd.ExecuteNonQuery();
                //                        //realitex
                //                        //String selectStr2 = "Update slcustm Set price_list = '" + drow3["NEW"] + "' where customer = '" + drow3["new_account"].ToString().Trim() + "'";

                //                    }
                                    if (RecordExists2("SELECT cd_custref FROM custdoc WHERE cd_custref = '" + drow["customer"].ToString().Replace("'", "''") + "' and cd_docno = '" + drow3["cd_docno"].ToString().Replace("'", "''") + "'"))
                                    {

                                        String selectStrR = "UPDATE custdoc SET cd_custref = '" + drow["customer"].ToString().Replace("'", "''") + "', cd_name = '" + drow3["cd_name"].ToString().Replace("'", "''").Trim() +
                                            "', cd_output = '" + drow3["cd_output"].ToString().Replace("'", "''") + "', cd_popup = '" + drow3["cdpopup"].ToString().Replace("'", "''") + "' WHERE cd_custref = '" + drow["customer"].ToString().Replace("'", "''") + "' and cd_docno = '" + drow3["cd_docno"].ToString().Replace("'", "''") + "'"; 
                                                        


                                        OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                                        cmd2.ExecuteNonQuery();
                                    
                                    
                                    }
                                    else
                                    {

                                        String selectStrR = "INSERT INTO custdoc (cd_custref, cd_docno, cd_name, cd_output, cd_popup)" +
                                                      "VALUES ('" + drow["customer"].ToString().Replace("'", "''") + "', '" + drow3["cd_docno"].ToString().Replace("'", "''") + "', '" + drow3["cd_name"].ToString().Replace("'", "''").Trim() + "','" + drow3["cd_output"].ToString().Replace("'", "''") + "','" + drow3["cdpopup"].ToString().Replace("'", "''") + "')";


                                        OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                                        cmd2.ExecuteNonQuery();
                                    }
                                }
                                catch (Exception e)
                                {
                                    cn.Close();
                                    CloseCnnPL();
                                    //CloseCnnsql();
                                    MessageBox.Show(e.Message.ToString());
                                }
                            }
                        }
                    }
                    //}


                }
            }
            CloseCnnPL();
            cn.Close();
            //CloseCnnsql();
        }

        public void matchW2()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM pl1";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "pl1");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            OpenCnnPL();
            row3 = ds.Tables[0].Rows[0];
            String today = DateTime.Today.ToString("dd/MM/yyyy");
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                //string amount1 = "";


                if (drow3.RowState != DataRowState.Deleted)
                {


                    //String selectStr = "SELECT customer FROM slcustm where customer Like 'T%'";
                    //OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);

                    //ds = new DataSet();
                    //da.Fill(ds, "slcustm");
                    //DataRow row;
                    //if (ds.Tables[0].Rows.Count > 0)
                    //{
                    //    row = ds.Tables[0].Rows[0];

                    //    foreach (DataRow drow in ds.Tables[0].Rows)
                    //    {
                    //        //string cutprice;
                    //        //string rollprice;
                    //        //if (drow3["CUTS"].ToString() == "0.00")
                    //        //{
                    //        //    cutprice = drow["cutprice1"].ToString();
                    //        //}
                    //        //else
                    //        //{
                    //        //    cutprice = drow3["CUTS"].ToString();
                    //        //}

                    //        //if (drow3["ROLLS"].ToString() == "0.00")
                    //        //{
                    //        //    rollprice = drow["rollprice1"].ToString();
                    //        //}
                    //        //else
                    //        //{
                    //        //    rollprice = drow3["ROLLS"].ToString();
                    //        //}

                    //        if (drow.RowState != DataRowState.Deleted)
                    //        {

                                try
                                {
                                    //                    //String selectStr = "Update Customer_1 Set Cust_Price_List = '" + drow3["NEW"] + "' where Cus_ID = '" + drow3["new_account"].ToString().Trim() + "'";
                                    //                    if (RecordExists("SELECT PL_Name FROM Price_Lists WHERE PL_Name = '" + drow3["NEW"].ToString().Replace("'", "''") + "' and PL_Stock_Code = '" + drow["stock_code"].ToString().Replace("'", "''") + "'"))
                                    //                    { }
                                    //                    else
                                    //                    {
                                    //                        String selectStrPL = "INSERT INTO Price_Lists (PL_Name, PL_Description, PL_Price1, PL_Price2, PL_Stock_Code, PL_Stock_Description, PL_Design_Code, PL_Design_Description, PL_Stock_Type, PL_Sequence, PL_Update, PL_Update_Date)" +
                                    //"VALUES ('" + drow3["NEW"].ToString().Replace("'", "''") + "', '', '" + cutprice + "', '" + rollprice +
                                    //"', '" + drow["stock_code"].ToString().Replace("'", "''") + "', '" + drow["stockdescrip"].ToString().Replace("'", "''") + "', '" + drow["design_code"].ToString().Replace("'", "''") + "', '" + drow["design_descrip"].ToString().Replace("'", "''") + "', '5', '1', '1', '" + today + "')";


                                    //                        SqlCommand cmd = new SqlCommand(selectStrPL, cnnpl);
                                    //                        cmd.ExecuteNonQuery();
                                    //                        //realitex
                                    //                        //String selectStr2 = "Update slcustm Set price_list = '" + drow3["NEW"] + "' where customer = '" + drow3["new_account"].ToString().Trim() + "'";

                                    //                    }
                                    //if (RecordExists2("SELECT cd_custref FROM custdoc WHERE cd_custref = '" + drow["customer"].ToString().Replace("'", "''") + "' and cd_docno = '" + drow3["cd_docno"].ToString().Replace("'", "''") + "'"))
                                    //{ }
                                    //else
                                    //{

                                        //String selectStrR = "INSERT INTO custdoc (cd_custref, cd_docno, cd_name, cd_output, cd_popup)" +
                                        //              "VALUES ('" + drow["customer"].ToString().Replace("'", "''") + "', '" + drow3["cd_docno"].ToString().Replace("'", "''") + "', '" + drow3["cd_name"].ToString().Replace("'", "''") + "','" + drow3["cd_output"].ToString().Replace("'", "''") + "','" + drow3["cdpopup"].ToString().Replace("'", "''") + "')";


                                    String selectStrR = "Update Customer_1 Set Cust_PL_Description = '" + drow3["name"].ToString().Replace("'", "''").Trim() + "' where Cust_Price_List = '" + drow3["list"].ToString().Trim() + "'";



                                    SqlCommand cmd2 = new SqlCommand(selectStrR, cnnpl);
                                    cmd2.ExecuteNonQuery();

                                    //String selectStrR2 = "Update Price_Lists Set PL_Description = '" + drow3["name"].ToString().Replace("'", "''").Trim() + "' where PL_Name = '" + drow3["list"].ToString().Trim() + "'";



                                    //SqlCommand cmd22 = new SqlCommand(selectStrR2, cnnpl);
                                    //cmd22.ExecuteNonQuery();






                                    //}
                                }
                                catch (Exception e)
                                {
                                    cn.Close();
                                    CloseCnnPL();
                                    CloseCnnsql();
                                    MessageBox.Show(e.Message.ToString());
                                }
                            }
                        }
            //        }
            //        //}


            //    }
            //}
            //CloseCnnPL();
            cn.Close();
            CloseCnnsql();
        }

        public void match99()
        {

            DataSet ds = null;

            try
            {
                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();

                String selectStr = "SELECT * FROM stock where st_warehouse = 'T' AND st_stockdesc LIKE '%ECLIPSE%'";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);

                ds = new DataSet();
                da.Fill(ds, "stock");

                //CloseCnnsql();
            }
            catch (Exception e)
            {
                cn.Close();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;
           
            
            row3 = ds.Tables[0].Rows[0];
           
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                
                if (drow3.RowState != DataRowState.Deleted)
                {

                    try 
                    {

                        String selectStrR = "UPDATE stock SET st_stockdesc = '" + drow3["st_stockdesc"].ToString().Replace("ECLIPSE", "SANDRINGHAM").Trim() + "' WHERE st_warehouse = 'T' AND st_stockref = '" + drow3["st_stockref"].ToString() + "'";



                        OdbcCommand cmd2 = new OdbcCommand(selectStrR, cn);
                        cmd2.ExecuteNonQuery();

                    }
                    catch (Exception e)
                    {
                        cn.Close();
                     
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }
        
            cn.Close();
         
        }
        public void matchP()
        {

            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            try
            {
          
                String selectStr = "UPDATE design SET de_designdesc = REPLACE(de_designdesc, 'ECLIPSE', 'SANDRINGHAM') WHERE de_warehouse = 'T' AND de_designdesc LIKE '%ECLIPSE%'";

                OdbcCommand cmd2 = new OdbcCommand(selectStr, cn);
                cmd2.ExecuteNonQuery();

            }
            catch (Exception e)
            {
                cn.Close();
              
                MessageBox.Show(e.Message.ToString());
            }
            cn.Close();
        }

        //updatepieces
        public void stockpiece()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM stocktw order by rack";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "stocktw");


            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;

            row3 = ds.Tables[0].Rows[0];
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            string newac;
            int i = 460777;
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {

                if (drow3.RowState != DataRowState.Deleted)
                {

                    newac = i.ToString();
               
                    //string thereid = drow3["rollno"].ToString().Trim() + " " + drow3["rack"].ToString().Trim();
                    String daterecived = DateTime.Now.ToString("dd/MM/yyyy");
                    
                    try
                    {
                        String strTable = "Insert into stockpiece";
                        String strFields = " (sp_stockref, sp_pieceno, sp_costsqm, sp_price, sp_width, sp_length, sp_remain, sp_allocated, sp_reserved, sp_datereceived, sp_batch, sp_manrollid, sp_warehouse, sp_dyeno, sp_grade, sp_location, sp_cutroll," +
                        " sp_stockind, sp_checked)";

                        String strValues = " Values ('" + drow3["productcode"].ToString().Trim() + "', '" + newac +
                        "', '0', '0', '" + drow3["width"].ToString().Trim() + "', '" + drow3["length"].ToString().Trim() + "', '" + drow3["length"].ToString().Trim() + "', '0', '0', '" + daterecived + "', '" + drow3["blend"].ToString().Trim() +
                        "', '" + drow3["rollid"].ToString().Trim() + "', 'T', '" + drow3["rack"].ToString().Trim() + "', 'A', 'NA', 'R', '0', 'N')";

                        String insertStr2 = strTable + strFields + strValues;

                        OdbcCommand cmd2 = new OdbcCommand(insertStr2, cn);
                        cmd2.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {
                        cn.Close();
                        CloseCnnsql();
                        MessageBox.Show(e.Message.ToString());
                    }

                    i++;

                   
                }
            }
            cn.Close();
            CloseCnnsql();
        }

        public void stockupdate()
        {

            DataSet ds = null;

            try
            {
                OpenCnnsql();

                String selectStr = "SELECT * FROM stockatlast";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "stockatlast");


            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

            DataRow row3;

            row3 = ds.Tables[0].Rows[0];
            cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
            cn.Open();
            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                if (RecordExists2("SELECT st_stockref FROM stock WHERE st_stockref = '" + drow3["stockcode"].ToString().Trim() + "'"))
                {
                }
                else
                {

                    if (drow3.RowState != DataRowState.Deleted)
                    {
                        string stockdesc = drow3["desc1"].ToString().Trim() + " " + drow3["disc2"].ToString().Trim();
                        if (stockdesc.Length > 30)
                        {
                            stockdesc = stockdesc.Remove(30);
                        }

                        string alpha;
                        if (drow3["desc1"].ToString().Length > 8)
                        {
                            string txtalpha = drow3["desc1"].ToString().Replace("'", "''");
                            alpha = txtalpha.Remove(8);
                        }
                        else
                        {
                            alpha = drow3["desc1"].ToString().Replace("'", "''");
                        }

                        string descrip;
                        if (drow3["desc1"].ToString().Length > 20)
                        {
                            string txtdescrip = drow3["desc1"].ToString().Replace("'", "''");
                            descrip = txtdescrip.Remove(20);
                        }
                        else
                        {
                            descrip = drow3["desc1"].ToString().Replace("'", "''");
                        }

                        string longdescrip = drow3["desc1"].ToString().Trim() + " " + drow3["disc2"].ToString().Trim();
                        if (longdescrip.Length > 40)
                        {
                            longdescrip = longdescrip.Remove(40);
                        }

                        try
                        {
                            String strTable = "Insert into stock";
                            String strFields = " (st_stockref, st_stockdesc, st_vatcode, st_factor, st_stocktype, st_designref, st_location, st_nominalcode, st_unitdesc, st_discount, st_manpur, st_width, st_warehouse)";

                            String strValues = " Values ( '" + drow3["stockcode"].ToString().Trim() + "', '" + stockdesc +
                            "', 'V', '1.00', '5', '" + drow3["design"].ToString().Trim() + "', 'NA', 'T-00-01-TUF', 'LIN.M', 'Y', 'P', '" + drow3["width"].ToString().Trim() + "', 'T')";

                            String insertStr2 = strTable + strFields + strValues;

                            OdbcCommand cmd2 = new OdbcCommand(insertStr2, cn);
                            cmd2.ExecuteNonQuery();
                        }
                        catch (Exception e)
                        {
                            cn.Close();
                            CloseCnnsql();
                            MessageBox.Show(e.Message.ToString());
                        }

                        try
                        {
                            String strTable2 = "Insert into stockm";
                            String strFields2 = " (warehouse, product, alpha, description, long_description, purchase_unit, selling_unit, despatch_units)";

                            String strValues2 = " Values ('T', '" + drow3["stockcode"].ToString().Trim() + "', '" + alpha + "', '" + descrip +
                            "', '" + longdescrip + "', 'ROLL', 'ROLL', 'LINEAR')";

                            String insertStr22 = strTable2 + strFields2 + strValues2;

                            OdbcCommand cmd22 = new OdbcCommand(insertStr22, cn);
                            cmd22.ExecuteNonQuery();
                        }
                        catch (Exception e)
                        {
                            cn.Close();
                            CloseCnnsql();
                            MessageBox.Show(e.Message.ToString());
                        }



                    }
                }
            }
            cn.Close();
            CloseCnnsql();
        }

        public void exelinfo(String file, String sheet)
        {
            DataSet ds = null;
            DataTable table = null;
            DataTable myTable = null;
            try
            {
                string excelConnectString = "Provider = Microsoft.Jet.OLEDB.4.0;" +
                "Data Source = " + file + ";" +
                "Extended Properties = Excel 8.0;";

                OleDbConnection objConn = new OleDbConnection(excelConnectString);
                //table = objConn.GetSchema("Tables");
                OleDbCommand objCmd = new OleDbCommand("Select * From [" + sheet + "$]", objConn);

                OleDbDataAdapter objDatAdap = new OleDbDataAdapter();
                objDatAdap.SelectCommand = objCmd;
                ds = new DataSet();
                objDatAdap.Fill(ds, 1, 1500000, sheet);
                table = ds.Tables[sheet];
                myTable = table.Clone();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
            try
            {
                OpenCnnsql();
                SqlTableCreater ssss = new SqlTableCreater();
                ssss.DestinationTableName = sheet;
                messages2("request made for " + sheet);
                ssss.Connection = cnnsql;
                ssss.CreateFromDataTable(myTable);
                messages2("send to bulk upload " + sheet);
                BulkInsertDataTableSQL(sheet, table);
                ssss.Transaction.Commit();
                CloseCnnsql();



            }
            catch (Exception err)
            {
                CloseCnnsql();


                myTable.Dispose();
                table.Dispose();
                String Str = err.Message;

                error_messages("testing_P2", Str, sheet);

            }


            myTable.Dispose();
            table.Dispose();

            
        }


        private void BulkInsertDataTableSQL(string tableName, DataTable table)
        {
            String connectionString = "Data Source=192.168.20.5;Initial Catalog=twitter;Persist Security Info=True;User ID=Price_List2;Password=Price_List2010";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlBulkCopy bulkCopy =
                        new SqlBulkCopy
                        (
                        connection,
                        SqlBulkCopyOptions.TableLock |
                        SqlBulkCopyOptions.FireTriggers |
                        SqlBulkCopyOptions.UseInternalTransaction,
                        null
                        );

                    bulkCopy.DestinationTableName = tableName;
                    bulkCopy.BulkCopyTimeout = 2140000000;
                    connection.Open();


                    bulkCopy.WriteToServer(table);
                    connection.Close();
                }
            }
            catch (Exception err)
            {
                String Str = err.Message;

                error_messages("BulkInsertDataTable", Str, tableName);
                //    MessageBox.Show(err.Message.ToString());
            }
        }

        private string RemoveSpaceandNewAcc(string inputString)
        {

            string newString = inputString.Replace(" ", "");
            newString = newString.Trim();
            newString = newString.Substring(0, 3);
            return newString;
        }
        public void getsecond()
        {
            DataSet ds = null;

            try
            {

                OpenCnnsql();

                String selectStr = "SELECT InvAddress1, Postcode FROM twitter Group By InvAddress1, Postcode";
                SqlDataAdapter da = new SqlDataAdapter(selectStr, cnnsql);

                ds = new DataSet();
                da.Fill(ds, "twitter");


            }
            catch (Exception e)
            {
                CloseCnnsql();
                MessageBox.Show(e.Message.ToString());
            }

             DataRow row3;
            
            row3 = ds.Tables[0].Rows[0];

            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {

                if (drow3.RowState != DataRowState.Deleted)
                {

                    try
                    {
                        if (drow3["InvAddress1"].ToString().Replace("'", "''") != "")
                        {
                            if (RecordExists("SELECT customer FROM slcustm WHERE name = '" + drow3["InvAddress1"].ToString().Replace("'", "''") + "'"))
                            {
                            }
                            else
                            {

                                string newacc = RemoveSpaceandNewAcc(drow3["InvAddress1"].ToString().Replace("'", "''"));
                                newacc = SumDoWhile("T" + newacc);
                                String selectStr = "Update twitter Set secacc = '" + newacc + "' Where InvAddress1 = '" + drow3["InvAddress1"].ToString().Replace("'", "''") + "'";
                                SqlCommand cmd = new SqlCommand(selectStr, cnnsql);
                                cmd.ExecuteNonQuery();

                                String alpha;
                                if (drow3["InvAddress1"].ToString().Length > 8)
                                {
                                    String txtalpha = drow3["InvAddress1"].ToString().Replace("'", "''");
                                    alpha = txtalpha.Remove(8);
                                }
                                else
                                {
                                    alpha = drow3["InvAddress1"].ToString().Replace("'", "''");
                                }

                                //alpha = alpha.Replace("(", "''");
                                //update sql table customer
                                String strTable = "Insert into customer";
                                String strFields = " (cu_custref, cu_repref, cu_cutroll, cu_delivery, cu_disctype, cu_invdays, cu_settday1, cu_settdisc1, cu_settday2, cu_settdisc2, cu_cutdisc, cu_rolldisc, cu_otherdisc, cu_defaultgrade," +
                                " cu_cutsetup, cu_cutcharge, cu_balorder, cu_balreserved, cu_balforward, cu_disputed, cu_value01, cu_value02, cu_value03, cu_value04, cu_value05, cu_value06, cu_value07, cu_value08, cu_value09, cu_value10," +
                                "  cu_value11, cu_value12, cu_valueytd, cu_valuemonth, cu_initcred, cu_comments, cu_onstop, cu_stopreason, cu_confirmfax, cu_overaction, cu_sequence, cu_value1, cu_surchper1, cu_surchval1, cu_value2, cu_surchper2, cu_surchval2," +
                                " cu_value3, cu_surchper3, cu_surchval3, cu_prosurval, cu_prosurper, cu_showdis, cu_supplement, cu_priority)";

                                String strValues = " Values ( '" + newacc +
                                "', '', 'N', 'N', 'C', '0', '', '0.000', '', '0.000', '0.000', '0.000', '0.000', 'A', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', ''," +
                                " 'N', '', '', '', '0', '0', '0.000', '0', '0', '0.000', '0', '0', '0.000', '0', '0', '0.000', '', '0.000', '0')";

                                String insertStr2 = strTable + strFields + strValues;

                                SqlCommand cmd2 = new SqlCommand(insertStr2, cnnsql);
                                cmd2.ExecuteNonQuery();
                                string now = DateTime.Now.ToString("dd/mm/yyyy");



                                //insert into slcustm table
                                String strTablesl = "Insert into slcustm";
                                String strFieldssl = " (customer, alpha, name, address1, address2, address3, address4, address5, credit_category, export_indicator, cust_disc_code, currency, territory, class, region, statement_customer, invoice_customer," +
                                " analysis_codes1, analysis_codes2, analysis_codes3, analysis_codes4, analysis_codes5, settlement_code, sett_days_code, price_list, letter_code, balance_fwd, credit_limit, ytd_sales," +
                                " ytd_cost_of_sales, cumulative_sales, order_balance, sales_nl_cat, special_price, vat_registration, direct_debit, invoices_printed, consolidated_inv, comment_only_inv, bank_account_no, bank_sort_code," +
                                " bank_name, bank_address1, bank_address2, bank_address3, bank_address4, analysis_code_6, produce_statements, edi_customer, vat_type, language, delivery_method, carrier, vat_reg_number, vat_exe_number," +
                                " paydays1, paydays2, paydays3, bank_branch_code, print_cp_with_stat, payment_method, customer_class, sales_type, cp_lower_value, address6, fax, telex, btx, cp_charge, control_digit, payer, responsibility," +
                                " despatch_held, credit_controller, reminder_letters, severity_days1, severity_days2, severity_days3, severity_days4, severity_days5, severity_days6, delivery_reason, shipper_code1, shipper_code2, shipper_code3," +
                                " shipping_note_ind, account_type, admin_fee, intrest_rate)";
                                String strValuessl = " Values ('" + newacc +
                                "', '" + alpha + "', '" + drow3["InvAddress1"].ToString().Replace("'", "''") + "', '', '', '', ''," +
                                " '', '', '', '', '', '', '', '', 'ZTWI999', '" + newacc + "', '', '', '', '', '', '', '', '', '', 'N', '0', '0', '0', '0', " +
                                "'0', '', '', '', '', 'Y', '', 'N', '', '', '', '', '', '', '', '', 'Y', 'N', '', 'language', '', '', '', '', '', '', '', '', 'N', '', '', '', '0', '" + drow3["Postcode"].ToString().Replace("'", "''") +
                                "', '', '', '', '', '', '', '', 'N', '', 'y', '0', '0', '0', '0', '0', '0', '', '', '', '', 'Y', '', '', '')";
                                String insertStrsl = strTablesl + strFieldssl + strValuessl;

                                SqlCommand cmdsl = new SqlCommand(insertStrsl, cnnsql);
                                cmdsl.ExecuteNonQuery();



                            }
                        }
                    }
                    catch (Exception e)
                    {
                        CloseCnnsql();
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }

        }

        public void doupdateme()
        {
            DataSet ds = null;

            try
            {
                cn = new OdbcConnection("dsn=View;UID=informix;PWD=WIBBLE;");
                cn.Open();

                String selectStr = "SELECT * FROM slcustm where customer Like 'CSCS%'";
                OdbcDataAdapter da = new OdbcDataAdapter(selectStr, cn);

                ds = new DataSet();
                da.Fill(ds, "slcustm");
                

            }
            catch (Exception e)
            {
                cn.Close();
                MessageBox.Show(e.Message.ToString());
            }
            
             DataRow row3;
            
            row3 = ds.Tables[0].Rows[0];

            foreach (DataRow drow3 in ds.Tables[0].Rows)
            {
                string newaccount = drow3["customer"].ToString().Substring(1);
                newaccount = "T" + newaccount;

                if (drow3.RowState != DataRowState.Deleted)
                {
                  
                    try
                    {
                        if (newaccount != "")
                        {
                            //if (RecordExists2("SELECT cu_custref FROM customer WHERE cu_custref = '" + newaccount + "'"))
                            //{
                            //}
                            //else
                            //{
                               
                            //    String strTable = "Insert into customer";
                            //    String strFields = " (cu_custref, cu_repref, cu_routeref, cu_cutroll, cu_delivery, cu_disctype, cu_credterm, cu_condays, cu_invdays, cu_settday1, cu_settdisc1, cu_settday2, cu_settdisc2, cu_cutdisc, cu_rolldisc, cu_otherdisc, cu_deliveryref, cu_defaultgrade," +
                            //    " cu_cutsetup, cu_cutcharge, cu_balorder, cu_balreserved, cu_balforward, cu_disputed, cu_value01, cu_value02, cu_value03, cu_value04, cu_value05, cu_value06, cu_value07, cu_value08, cu_value09, cu_value10," +
                            //    "  cu_value11, cu_value12, cu_valueytd, cu_valuemonth, cu_initcred, cu_comments, cu_onstop, cu_stopreason, cu_confirmfax, cu_overaction, cu_sequence, cu_value1, cu_surchper1, cu_surchval1, cu_value2, cu_surchper2, cu_surchval2," +
                            //    " cu_value3, cu_surchper3, cu_surchval3, cu_prosurval, cu_prosurper, cu_showdis, cu_supplement, cu_priority)";

                            //    String strValues = " Values ( '" + newaccount +
                            //    "' , 'T20',  '" + drow3["cu_routeref"].ToString() + 
                            //    "', 'N', 'N', 'C', 'E30', '0', '0', 'E30', '5.000', '', '0.000', '0.000', '0.000', '0.000', '" + drow3["cu_deliveryref"].ToString() + "', 'A', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', ''," +
                            //    " 'N', '', '', '', '0', '0', '0.000', '0', '0', '0.000', '0', '0', '0.000', '0', '0', '0.000', '', '0.000', '0')";

                            //    String insertStr2 = strTable + strFields + strValues;

                            //    OdbcCommand cmd2 = new OdbcCommand(insertStr2, cn);
                            //    cmd2.ExecuteNonQuery();
                            //}
                            string now = "30/03/2013";

                            if (RecordExists2("SELECT customer FROM slcustm WHERE customer = '" + newaccount + "'"))
                            {
                            }
                            else
                            {

                                //insert into slcustm table
                                String strTablesl = "Insert into slcustm";
                                String strFieldssl = " (customer, alpha, name, address1, address2, address3, address4, address5, credit_category, export_indicator, cust_disc_code, currency, territory, class, region, invoice_customer, statement_customer," +
                                " group_customer, date_created, analysis_codes1, analysis_codes2, analysis_codes3, analysis_codes4, analysis_codes5, reminder_cat, settlement_code, sett_days_code, price_list, letter_code, balance_fwd, credit_limit, ytd_sales," +
                                " ytd_cost_of_sales, cumulative_sales, order_balance, sales_nl_cat, special_price, vat_registration, direct_debit, invoices_printed, consolidated_inv, comment_only_inv, bank_account_no, bank_sort_code," +
                                " bank_name, bank_address1, bank_address2, bank_address3, bank_address4, analysis_code_6, produce_statements, edi_customer, vat_type, language, delivery_method, carrier, vat_reg_number, vat_exe_number," +
                                " paydays1, paydays2, paydays3, bank_branch_code, print_cp_with_stat, payment_method, customer_class, sales_type, cp_lower_value, address6, fax, telex, btx, cp_charge, control_digit, payer, responsibility," +
                                " despatch_held, credit_controller, reminder_letters, severity_days1, severity_days2, severity_days3, severity_days4, severity_days5, severity_days6, delivery_reason, shipper_code1, shipper_code2, shipper_code3," +
                                " shipping_note_ind, account_type, admin_fee, intrest_rate)";
                                String strValuessl = " Values ('" + newaccount +
                                "', '" + drow3["alpha"].ToString().Replace("'", "''") + "', '" + drow3["name"].ToString().Replace("'", "''") + "', '" + drow3["address1"].ToString().Replace("'", "''") + "', '" + drow3["address2"].ToString().Replace("'", "''") + "', '" + drow3["address3"].ToString().Replace("'", "''") + "', '" + drow3["address4"].ToString().Replace("'", "''") + "'," +
                                " '" + drow3["address5"].ToString().Replace("'", "''") + "', '', '', '', '', 'T20', '', '', 'TSCS001', 'ZSCS001', 'TSCSREB', '" + now + "',  '', '', '', '', '', '', '', '', '', '', 'N', '0', '0', '0', '0', " +
                                "'0', '', '', '', '', 'Y', '', 'N', '', '', '', '', '', '', '', '', 'Y', 'N', '', 'language', '', '', '', '', '', '', '', '', 'N', '', '', '', '0', '" + drow3["address6"].ToString().Replace("'", "''") +
                                "', '', '', '', '', '', '', '', 'N', '', 'y', '0', '0', '0', '0', '0', '0', '', '', '', '', 'Y', '', '', '')";
                                String insertStrsl = strTablesl + strFieldssl + strValuessl;

                                OdbcCommand cmdsl = new OdbcCommand(insertStrsl, cn);
                                cmdsl.ExecuteNonQuery();
                            }
                        }

                    }
                    catch (Exception e)
                    {
                        cn.Close();
                       
                        MessageBox.Show(e.Message.ToString());
                    }
                }
            }
            cn.Close();
         

        }

        private string SumDoWhile(string acc)
        {
            string newac;
            int i = 001;
            do
            {
              newac = acc + i.ToString("000");
                i++;
            } while (RecordExists("SELECT customer FROM slcustm WHERE customer = '" + newac + "'"));

            return newac;
        }

        public bool RecordExists2(string _SQL)
        {
            //OpenCnn();

            OdbcDataReader _SqlDataReader = null;
            try
            {
                // Pass the connection to a command object
                OdbcCommand _SqlCommand = new OdbcCommand(_SQL, cn);
                // get query results 
                _SqlDataReader = _SqlCommand.ExecuteReader();

            }
            catch (Exception _Exception)
            {
                // Error occurred while trying to execute reader
                // send error message to console (change below line to customize error handling)
                Console.WriteLine(_Exception.Message);
                return false;
            }
            if (_SqlDataReader != null && _SqlDataReader.Read())
            {
                // close sql reader before exit 
                if (_SqlDataReader != null)
                {
                    _SqlDataReader.Close();
                    _SqlDataReader.Dispose();
                }
                // record found 
                return true;
            }
            else
            {
                // close sql reader before exit  
                if (_SqlDataReader != null)
                {
                    _SqlDataReader.Close();
                    _SqlDataReader.Dispose();
                }

                // record not found
                return false;
            }
            //CloseCnn();
        }


        public bool RecordExists(string _SQL)
        {
            //OpenCnn();

            SqlDataReader _SqlDataReader = null;
            try
            {
                // Pass the connection to a command object
                SqlCommand _SqlCommand = new SqlCommand(_SQL, cnnpl);
                // get query results 
                _SqlDataReader = _SqlCommand.ExecuteReader();

            }
            catch (Exception _Exception)
            {
                // Error occurred while trying to execute reader
                // send error message to console (change below line to customize error handling)
                Console.WriteLine(_Exception.Message);
                return false;
            }
            if (_SqlDataReader != null && _SqlDataReader.Read())
            {
                // close sql reader before exit 
                if (_SqlDataReader != null)
                {
                    _SqlDataReader.Close();
                    _SqlDataReader.Dispose();
                }
                // record found 
                return true;
            }
            else
            {
                // close sql reader before exit  
                if (_SqlDataReader != null)
                {
                    _SqlDataReader.Close();
                    _SqlDataReader.Dispose();
                }

                // record not found
                return false;
            }
            //CloseCnn();
        }
    }
}
