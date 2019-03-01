using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Read_From_Excel.getExcelFile();
            Console.ReadLine();
        }

    
    }

    //microsoft Excel 14 object in references-> COM tab

        public class Read_From_Excel
        {
            public static void getExcelFile()
            {

                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\KSIUDA\Desktop\SQL\Training List.xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                var strings = new List<Tuple<int, string, string, string>>();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
                for (int row = 3; row <= rowCount; row++)
                {
                    string training = "";
                    string instructor = "";
                    string duration = "";

                    for (int column = 2; column <= 7; column++)
                    {
                       
                        if (column == 2)
                        {
                            if (xlRange.Cells[row, column] != null && xlRange.Cells[row, column].Value2 != null)
                            {
                                int index0 = row - 3; //since we read from the 3rd row and we want to count from 0
                                int index1 = column - 2; //since we read from the 3rd column, and we want to count from 1

                                string value = xlRange.Cells[row, column].Value2.ToString();
                                if (value.ToUpper() != value)
                                {
                                    training = value;
                                }
                            

                            }
                        }
                       
                        if (column == 6)
                        {

                            if (xlRange.Cells[row, column] != null && xlRange.Cells[row, column].Value2 != null)
                            {
                                string value = xlRange.Cells[row, column].Value2.ToString();

                                if (value.ToUpper() != value)
                                {
                                    instructor = value;
                                }
                            

                            }
                        }

                    if (column == 7)
                    {

                        if (xlRange.Cells[row, column] != null && xlRange.Cells[row, column].Value2 != null)
                        {
                            string value = xlRange.Cells[row, column].Value2.ToString();

                            if (value.ToUpper() != value)
                            {
                                duration = value;
                            }


                        }
                    }
                }

                if (instructor != "" || training != "")
                {
                    if (instructor == "")
                    {
                        instructor = "None";
                    }
                    strings.Add(Tuple.Create(strings.Count() + 1, training, instructor, duration));
                }
            }


            foreach (var tuple in strings)
            {
                Console.Write($" {tuple.Item1}, {tuple.Item2}, {tuple.Item3}, {tuple.Item4}");
                Console.WriteLine();
            }


                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
       
    }


}

