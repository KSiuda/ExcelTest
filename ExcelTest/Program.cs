using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

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
                string[,] TrainingDetails = new string[rowCount - 3, 2];

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 3; i <= rowCount; i++)
                {
                    for (int j = 2; j <= 6; j++)
                    {
                        //new line
                    if (j == 2)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            string value = xlRange.Cells[i, j].Value2.ToString();
                            if (value.ToUpper() != value)
                            {
                                TrainingDetails[i - 3, j-2] = value;
                                //Console.WriteLine(value + "\t");
                            }

                        }
                    }
                    if (j == 6)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            string value = xlRange.Cells[i, j].Value2.ToString();
                            if (value.ToUpper() != value)
                            {
                                TrainingDetails[i - 3, j - 5] = value;
                                //Console.WriteLine(value + "\t");
                            }

                        }
                    }
                       
                    }
                }
            for (int i =0; i<TrainingDetails.GetLength(0); i++)
            {
                for (int j=0; j<TrainingDetails.GetLength(1)-1; j++)
                {
                    Console.WriteLine($"Training = {TrainingDetails[i,j]} , Instructor = {TrainingDetails[i,j+1]}");
                }
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

