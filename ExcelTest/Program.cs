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
                var trainings = new List<Tuple<int, string, string, string>>();
                var consultants = new List<Tuple<int,string>>();
                var trainingrecord = new List<Tuple<int, string>>();
                var blanks = new List<int>();
                int blank = 0;


            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int row = 3; row <= rowCount; row++)
                {
                    string training = "";
                    string instructor = "";
                    string duration = "";

                    for (int column = 2; column < 8; column++)
                    {
                       
                        if (column == 2)
                        {
                            if (xlRange.Cells[row, column] != null && xlRange.Cells[row, column].Value2 != null)
                            {

                                string value = xlRange.Cells[row, column].Value2.ToString();
                                if (value.ToUpper() != value)
                                {
                                    training = value;
                                }
                                else
                                {
                                    blanks.Add(row);
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
                               
                                    duration = value;

                            }
                        }


                }

                if (instructor != "" || training != "")
                {
                    if (instructor == "")
                    {
                        instructor = "None";
                    }
                    if(duration == "")
                    {
                        duration = null;
                    }
                    trainings.Add(Tuple.Create(trainings.Count() + 1, training, instructor, duration));
                }
            }




            for (int row = 1; row <= 1; row++)
            {
                string consultant = "";

                for (int column = 10; column <= colCount; column++)
                {
                    string value = xlRange.Cells[row, column].Value2.ToString();

                    consultant = value;
                    consultants.Add(Tuple.Create(consultants.Count() + 1, consultant));
                }
                  
            }

            
            foreach(var consultant in consultants)
            {
                blank = 0;
                for (int row = 3; row <= trainings.Count() + 2 + blanks.Count(); row++)
                {
                    string trainingState = "";
                    int counter = row - 2;

                    for (int column = 10 + consultant.Item1 - 1; column < 10 + consultant.Item1; column++)
                    {
                        if (xlRange.Cells[row, column] != null && xlRange.Cells[row, column].Value2 != null)
                        {

                            string value = xlRange.Cells[row, column].Value2.ToString();
                           
                                trainingState = value;
                                if(trainingState == "")
                                {
                                    blank++;
                                }
                                else
                                {
                                    trainingrecord.Add(Tuple.Create((counter - blank), trainingState));
                                }
                            
                        }
                        else
                        {
                            blank++;
                        }
                    }

                }
            }
               
            




            Console.WriteLine("Trainings:");
            foreach (var training in trainings)
            {
                Console.WriteLine($" {training.Item1}, {training.Item2}, {training.Item3}, {training.Item4}");
               
            }
            Console.WriteLine("Consultants");
            foreach (var consultant in consultants)
            {
                Console.WriteLine($" {consultant.Item1}, {consultant.Item2}");
            }
            Console.WriteLine("Training State for Marcin");
            foreach (var training in trainingrecord)
            {
                Console.Write($" {training.Item1} {training.Item2},");
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

