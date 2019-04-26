using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelinDotNetCoreEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            List<UserDetails> userDetails = ReadFromExcel<List<UserDetails>>("TestDataRead.xlsx");
            Console.WriteLine("ID   " + "Name " + "City " + "Country ");
            foreach (UserDetails details in userDetails)
            {
                Console.WriteLine(details.ID +" "+ details.Name +" "+ details.City+" " +details.Country);
            }
            //WriteToExcel("TestDataWrite.xlsx");


        }

        private static T ReadFromExcel<T>(string path, bool hasHeader = true)
        {
            using (var excelPack = new ExcelPackage())
            {
                //Load excel stream
                using (var stream = File.OpenRead(path))
                {
                    excelPack.Load(stream);
                }

                //Lets Deal with first worksheet.(You may iterate here if dealing with multiple sheets)
                var ws = excelPack.Workbook.Worksheets[0];

                //Get all details as DataTable -because Datatable make life easy :)
                DataTable excelasTable = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    //Get colummn details
                    if (!string.IsNullOrEmpty(firstRowCell.Text))
                    {
                        string firstColumn = string.Format("Column {0}", firstRowCell.Start.Column);
                        excelasTable.Columns.Add(hasHeader ? firstRowCell.Text : firstColumn);
                    }
                }
                var startRow = hasHeader ? 2 : 1;
                //Get row details
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, excelasTable.Columns.Count];
                    DataRow row = excelasTable.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
               //Get everything as generics and let my end user decides on casting to their choice
                var generatedType = JsonConvert.DeserializeObject<List<UserDetails>>(JsonConvert.SerializeObject(excelasTable));
                return (T)Convert.ChangeType(generatedType, typeof(T));
            }
        }
        private static void WriteToExcel(string path)
        {
            //Let use below test data for writing it to excel
            List<UserDetails> persons = new List<UserDetails>()
            {
                new UserDetails() {ID="9999", Name="ABCD", City ="City1", Country="USA"},
                new UserDetails() {ID="8888", Name="PQRS", City ="City2", Country="INDIA"},
                new UserDetails() {ID="7777", Name="XYZZ", City ="City3", Country="CHINA"},
                new UserDetails() {ID="6666", Name="LMNO", City ="City4", Country="UK"},
           };

            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting. 
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));
            FileInfo filePath = new FileInfo(path);
            using (var excelPack = new ExcelPackage(filePath))
            {
                var ws = excelPack.Workbook.Worksheets.Add("WriteTest");
                ws.Cells.LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Light8);
                excelPack.Save();
            }
        }
    }
}
