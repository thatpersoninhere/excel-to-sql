using System;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test_App_1
{
    class Program
    {
        static void Main(string[] args)
        {

            var spreadsheetLocation = Path.Combine(Directory.GetCurrentDirectory(), "Chapter 5.xlsx");

            string sexcelconnectionstring = @"provider=microsoft.jet.oledb.4.0;data source=" + spreadsheetLocation + ";extended properties=" + "\"excel 8.0;hdr=yes;\"";

            string myexceldataquery = "SELECT * FROM[SHEET1$]";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(spreadsheetLocation);
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            List<String> columnNames = new List<String>();
            bool stop = false;
            int i = 1;

            while (stop == false)
            {
                columnNames.Add(xlRange.Cells[1, i].Value2.ToString());
                i++;
                if (xlRange.Cells[1, i].Value2 == null)
                    stop = true;
            }

            string connectionString = @"server=localhost;persist security info=True; database=Test;integrated Security=SSPI;";

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand myCommand = new SqlCommand("CREATE TABLE newTable(temp INT)", con))
                {
                    string Result = (string)myCommand.ExecuteScalar(); // returns the first column of the first row
                }
                foreach (var item in columnNames)
                {
                    string com = "ALTER TABLE newTable ADD \"" + item + "\" VARCHAR(50)";
                    Console.WriteLine(com);
                    using (SqlCommand myCommand = new SqlCommand(com, con))
                    {
                        string Result = (string)myCommand.ExecuteScalar(); // returns the first column of the first row
                    }

                }
                using (SqlCommand myCommand = new SqlCommand("ALTER TABLE newTable DROP COLUMN temp", con))
                {
                    string Result = (string)myCommand.ExecuteScalar(); // returns the first column of the first row
                }


                OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
                OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                oledbconn.Open();
                OleDbDataReader dr = oledbcmd.ExecuteReader();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(connectionString);
                bulkcopy.DestinationTableName = "newTable";
                while (i != 0)
                {
                    bulkcopy.WriteToServer(dr);
                    i--;
                }
                dr.Close();
                oledbconn.Close();
            }



        }
    }
}
