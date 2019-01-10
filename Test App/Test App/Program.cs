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
            //change here to specify excel sheet
            string myexceldataquery = "SELECT * FROM[SHEET1$]";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(spreadsheetLocation);
            //always pulls from sheet 1 unless you change this
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
            // change here to specify DB
            string connectionString = @"server=localhost;persist security info=True; database=Test;integrated Security=SSPI;";



            //this creates and adds columns a new table called newTable

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand myCommand = new SqlCommand("CREATE TABLE newTable(temp INT)", con))
                {
                    string Result = (string)myCommand.ExecuteScalar();
                }
                foreach (var item in columnNames)
                {
                    //unfortunatly, i was only able to add data as varchar since it bulk adds any amount and type of columns

                    string com = "ALTER TABLE newTable ADD \"" + item + "\" VARCHAR(50)";
                    Console.WriteLine(com);
                    using (SqlCommand myCommand = new SqlCommand(com, con))
                    {
                        string Result = (string)myCommand.ExecuteScalar();


                    }
                    using (SqlCommand myCommand = new SqlCommand("ALTER TABLE newTable DROP COLUMN temp", con))
                    {
                        string Result = (string)myCommand.ExecuteScalar();
                    }
                    //temp tables can be removed if you feel like doing a ton of string interpolation


                    //this section does all the data adding
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
}
