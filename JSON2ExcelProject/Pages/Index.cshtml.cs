using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace JSON2ExcelProject.Pages
{
    public class IndexModel : PageModel
    {

        public string UserInput { get; set; }

        //properties


        public void OnGet()
        {
            UserInput = "hello2";
            System.Diagnostics.Debug.WriteLine("SomeText");
        }

        //global scope variables

        int counter = 0;
        int rowCounter = 1;
        int columnCounter = 0;


        public void OnPost(string jsontext)
        {
            UserInput = jsontext;

            CreateDataTable(UserInput);

            System.Diagnostics.Debug.WriteLine("SomeText");
            //CreateExcelFile();


        }

        public void CreateDataTable(string UserInput)

        {

            DataTable dataTable = (DataTable)JsonConvert.DeserializeObject(UserInput, (typeof(DataTable)));
            System.Diagnostics.Debug.WriteLine(dataTable.Columns.Count);



            //make new workbook

            // If you are using the Professional version, enter your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            ExcelFile workbook = new ExcelFile();
            ExcelWorksheet worksheet = workbook.Worksheets.Add("Users");


            foreach (DataColumn item in dataTable.Columns)
            {


                worksheet.Cells[0, counter].Value = $"{item.ColumnName}";
                counter++;


            }


            foreach (DataRow row in dataTable.Rows)



                for (int i = 0; i < dataTable.Columns.Count; i++)
                {

                    worksheet.Cells[rowCounter, columnCounter].Value = $"{row[i]}";

                    Console.WriteLine(row[i]);

                    columnCounter++;

                    if ((columnCounter) % dataTable.Columns.Count == 0)
                    {
                        rowCounter++;
                        columnCounter = 0;
                    }

                     workbook.Save("C:/Users/csapct86/Desktop/APItest.xlsx");


                }


        }


    }

}