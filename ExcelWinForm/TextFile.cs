using System;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelWinForm
{
    public static class TextFile
    {
        public static int groupCount { get; set; }
        public static bool isOpenedBook { get; set; }
        public static Workbook ObjWorkBook { get; set; }
        public static Workbooks workbooks { get; set; }


        public static int ReadFromFile(Workbook ObjWorkBook2, Workbooks workbooks2)
        {
            ObjWorkBook = ObjWorkBook2;
            workbooks = workbooks2;
            int count = 0;
            isOpenedBook = false;

            try
            {
                string line = "";
                bool flagReader = false;

                using (StreamReader sr = new StreamReader("Settings.txt"))
                {
                    var fi = new FileInfo("Settings.txt");
                    if (fi.Length == 0)
                    {
                        Extensions.PrintCatchedMessages("Error: File Settings.txt is empty. Check it and try again",
                                                        "Press Enter to end the program");
                        return -1;
                    }



                    File.Delete("pathToFile");

                    while (sr.Peek() != -1)
                    {
                        if (flagReader == false)
                        {
                            if ((char)sr.Peek() == '"')
                            {
                                sr.Read();
                                if (sr.Peek() == ' ')
                                {
                                    sr.Read();
                                }
                                flagReader = true;

                            }
                            else
                            {
                                sr.Read();
                            }
                        }
                        else
                        {


                            if ((char)sr.Peek() == '"')
                            {
                                flagReader = false;
                                sr.Read();



                                if (count == 0)
                                {
                                    try
                                    {
                                        ObjWorkBook = workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory + line, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                                        isOpenedBook = true;


                                    }
                                    catch (COMException)
                                    {
                                        Extensions.PrintCatchedMessages("Error: Wrong Excel file name. Edit and try again",
                                                       "Press Enter to end the program");
                                        return -1;

                                    }
                                    line = "";
                                    count++;
                                }
                                else if (count == 1)
                                {
                                    int number;
                                    bool result = Int32.TryParse(line, out number);

                                    if (result == true && number > 0 && number < 10000)
                                    {
                                        groupCount = Int32.Parse(line);
                                        line = "";
                                        count++;
                                    }
                                    else
                                    {
                                        Extensions.PrintCatchedMessages("Error: Wrong Amount of groups. Edit and try again",
                                                       "Press Enter to end the program");
                                        return -1;
                                    }
                                }
                            }
                            else
                            {
                                line = line + (char)sr.Read();
                            }
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Extensions.PrintCatchedMessages("Error: File Settings.txt wasn't found. Check it and try again",
                                                        "Press Enter to end the program");
                return -1;
            }
            catch (Exception ex)
            {
                Extensions.PrintCatchedMessages("Error ocurred",
                                                ex, "Press Enter to end the program");
                return -1;
            }
            return 0;
        }

    }
}
