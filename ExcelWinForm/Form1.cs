using System;
using System.Collections.Generic;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelWinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public static int staticRow; // число колонок в Excel
        private void button1_Click(object sender, EventArgs e)
        {
            //string fullfilename2007 = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Saved Excel results.xlsx");
            System.Diagnostics.Process[] objProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            //try
            //{
            //    if (System.IO.File.Exists(fullfilename2007)) System.IO.File.Delete(fullfilename2007);
            //}
            //catch (Exception)
            //{

                if (objProcess.Length > 0)
                {
                    System.Collections.Hashtable objHashtable = new System.Collections.Hashtable();

                    // check to kill the right process
                    foreach (System.Diagnostics.Process processInExcel in objProcess)
                    {
                        if (objHashtable.ContainsKey(processInExcel.Id) == false)
                        {
                            processInExcel.Kill();
                        }
                    }
                    objProcess = null;
                }
            //}






            Application ObjWorkExcel = null;
            Workbook ObjWorkBook = null;
            Worksheet ObjWorkSheet = null;
            Workbooks workbooks = null;
            Range lastCell = null;

            int count = 0; // счетчик
            int groupCount = 0; // количество групп
            bool isOpenedBook = false;
            bool isOpenedSheet = false;

            try
            {
                ObjWorkExcel = new Application(); //открыть эксель
                workbooks = ObjWorkExcel.Workbooks; //для избежания двух точек

                // считывание с файла данных
                //int returnedValue = TextFile.ReadFromFile(ObjWorkBook, workbooks);

                //if (returnedValue == -1)
                //{
                //    return;
                //}

                ObjWorkBook = workbooks.Open(txtCurrPath.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                groupCount = Int32.Parse(txtGroupCount.Text);
                //ObjWorkBook = TextFile.ObjWorkBook;
                //workbooks = TextFile.workbooks;
                //groupCount = TextFile.groupCount;
                //isOpenedBook = TextFile.isOpenedBook;



                if (ObjWorkBook == null)
                {
                    Extensions.PrintCatchedMessages("Error in the file Settings.txt: Cannot open Excel file. Check it and try again",
                                                        "Press Enter to end the program");
                    return;
                }
                else if (groupCount == 0)
                {
                    Extensions.PrintCatchedMessages("Error in the file Settings.txt: Wrong amount of groups. Check it and try again",
                                                        "Press Enter to end the program");
                    return;
                }


                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);//1 ячейку
                isOpenedSheet = true;

                int n = 4; // количество колонок без числовых данных в Excel

                List<User> userList = new List<User>(); // список 

                if (ObjWorkSheet.Cells[1, (int)lastCell.Column].Text.ToString() == "")
                {
                    Extensions.PrintCatchedMessages("Your Excel File has one ore more empty columns. Check it and try again",
                                                        "Press Enter to end the program");
                    return;
                }

                string[] programNames = new string[(int)lastCell.Column - n];
                staticRow = (int)lastCell.Row;
                //запись имен из Excel в массив
                ExcelFunctions.ReadExcelNames(lastCell, ObjWorkSheet, programNames, n);
                programNames = ExcelFunctions.programNames;

                User.ZeroCounter();
                //запись результатов из Excel в массив
                ExcelFunctions.ReadExcelData(lastCell, ObjWorkSheet, n);
                lastCell = ExcelFunctions.lastCell;
                ObjWorkSheet = ExcelFunctions.ObjWorkSheet;
                userList = ExcelFunctions.userList;

                #region commentedSort



                #endregion
                //Замена оценок 1-10 на 1-3
                int returnedValue = ExcelFunctions.ChangeMarks(lastCell, userList, n);
                userList = ExcelFunctions.userList;


                int humanCount = 0; // количество человек в Excel
                int remainder = 0; // лишние люди
                int humanPerGroup = 0; // количество человек для каждой программы


                //Подсчет количества групп и человек в них

                humanCount = ((int)lastCell.Row - 1) / groupCount;
                humanPerGroup = humanCount / ((int)lastCell.Column - n);

                if (humanPerGroup == 0)
                {
                    Extensions.PrintCatchedMessages("Error: Not enough people for this amount of groups. Edit and try again",
                                                        "Press Enter to end the program");
                    return;
                }

                if (((int)lastCell.Row - 1) % (((int)lastCell.Column - n) * groupCount) != 0)
                {
                    remainder = ((int)lastCell.Row - 1) % (((int)lastCell.Column - n) * groupCount);
                }

                //список для разных групп
                List<string[,]> reList = new List<string[,]>(); // новый список значений
                //создание списка групп

                if (remainder == 0)
                {
                    for (int i = 0; i < groupCount; i++)
                    {
                        reList.Add(new string[(int)lastCell.Column - n, humanPerGroup]);
                    }
                }
                else
                {
                    for (int i = 0; i < groupCount; i++)
                    {
                        reList.Add(new string[(int)lastCell.Column - n, humanPerGroup + 1]);
                    }
                }

                int[] countArr = new int[(int)lastCell.Column - n]; // подсчет людей для каждой колонки с выбранной оценкой

                List<int[]> groupAmount = new List<int[]>(); // список количества человек для текущей оценки

                for (int i = 0; i < groupCount; i++)
                {
                    groupAmount.Add(new int[(int)lastCell.Column - n]);
                }

                List<List<int>> wrongItemList = new List<List<int>>(); // список ячеек, в которых уже есть значение

                for (int i = 0; i < groupCount; i++)
                {
                    wrongItemList.Add(new List<int>());
                }

                //count = 0;

                //int currValue = 3; // текущая оценка
                //int currGroup = 0; // текущая группа
                int countI = 0; // номер строки
                //bool flag = true; // есть свободные ячейки
                //int id = 0; // идентификатор для Excel
                //int listId = 0; // идентификатор для List
                //int countId = 0; //переменная для подсчета идентификатора
                //int changeCount = 0; // переменная для смены группы
                //bool flagValue = false; // для count == -2

                int[,] arrAllMarks = new int[groupCount, (int)lastCell.Column - n]; // Общее количество оценок для каждой отсортированной колонки


                Extensions.countArr = countArr;
                Extensions.userList = userList;
                Extensions.reList = reList;
                Extensions.arrAllMarks = arrAllMarks;

                Extensions.SortPeople(lastCell, ObjWorkSheet, groupAmount, wrongItemList, humanPerGroup, humanCount, groupCount, remainder);

                countArr = Extensions.countArr;
                userList = Extensions.userList;
                reList = Extensions.reList;
                arrAllMarks = Extensions.arrAllMarks;

                // Если остались неотсортированные люди
                if (remainder != 0)
                {
                    Extensions.SortRemainder(remainder, lastCell, ObjWorkSheet, n, groupCount, wrongItemList, humanPerGroup);

                    userList = Extensions.userList;
                    reList = Extensions.reList;
                    arrAllMarks = Extensions.arrAllMarks;
                }


                // Добавить лист в конце.
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.Add(
                    Type.Missing, ObjWorkBook.Sheets[ObjWorkBook.Sheets.Count],
                    1, Excel.XlSheetType.xlWorksheet);

                ObjWorkSheet.Name = DateTime.Now.ToString("MM-dd-yy; HH-mm-ss");


                //ObjWorkExcel.Visible = true;
                //ObjWorkExcel.UserControl = true;

                count = 0;

                ObjWorkExcel.Columns.ColumnWidth = 40;
                ObjWorkExcel.Columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                ObjWorkExcel.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ObjWorkExcel.Columns.Font.Size = 12;
                //ObjWorkExcel.DisplayFullScreen = true;

                countI = 0;
                // вывод в Excel
                foreach (var item in reList)
                {

                    for (int i = 0; i < (int)lastCell.Column - n; i++)
                    {
                        ObjWorkSheet.Cells[1, i + 1 + count] = programNames[i] + $"\nGroup {countI + 1}";

                        if (remainder == 0)
                        {
                            for (int j = 0; j < humanPerGroup; j++)
                            {
                                ObjWorkSheet.Cells[j + 2, i + 1 + count] = item[i, j];
                            }
                        }
                        else
                        {
                            for (int j = 0; j < humanPerGroup + 1; j++)
                            {
                                ObjWorkSheet.Cells[j + 2, i + 1 + count] = item[i, j];
                            }
                        }
                    }
                    count = count + (int)lastCell.Column - n + 1;
                    countI++;
                }

                //Открытие Excel
                //ObjWorkExcel.Visible = true;
                //ObjWorkExcel.UserControl = true;
                //ObjWorkBook.Save();
                //string fullfilename2007 = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Saved Excel results.xlsx");
                //try
                //{
                //if (System.IO.File.Exists(fullfilename2007)) System.IO.File.Delete(fullfilename2007);
                //}
                //catch (Exception)
                //{
                //    System.Diagnostics.Process[] objProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");

                //    if (objProcess.Length > 0)
                //    {
                //        System.Collections.Hashtable objHashtable = new System.Collections.Hashtable();

                //        // check to kill the right process
                //        foreach (System.Diagnostics.Process processInExcel in objProcess)
                //        {
                //            if (objHashtable.ContainsKey(processInExcel.Id) == false)
                //            {
                //                processInExcel.Kill();
                //            }
                //        }
                //        objProcess = null;
                //    }
                //}

                //ObjWorkBook.SaveAs(fullfilename2007, Excel.XlFileFormat.xlWorkbookDefault);

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                    saveFileDialog.ShowDialog();
                    string filePath = saveFileDialog.FileName;
                    ObjWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault);
                    MessageBox.Show($"Success.Results were saved in the\n{filePath}");
                }
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Success\nResults were saved in the file 'Saved Excel results.xlsx'");
                Console.ResetColor();
                Console.WriteLine("Press Enter to end the program");
                //MessageBox.Show("Success");
            }
            catch (Exception ex)
            {
                Extensions.PrintCatchedMessages("Error ocurred",
                                                ex, "Press Enter to end the program");
                return;
            }
            finally
            {
                #region ClearExcel



                //ObjWorkExcel.Visible = false;
                //ObjWorkExcel.UserControl = false;

                Marshal.ReleaseComObject(ObjWorkExcel);

                if (isOpenedBook == true)
                {
                    ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя

                    Marshal.ReleaseComObject(ObjWorkBook);
                }

                if (isOpenedSheet == true)
                {
                    //ObjWorkExcel.Quit(); // выйти из экселя
                    Marshal.ReleaseComObject(ObjWorkSheet);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(lastCell);
                }



                GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !
                GC.WaitForPendingFinalizers();

                /*System.Diagnostics.Process[]*/
                objProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");

                if (objProcess.Length > 0)
                {
                    System.Collections.Hashtable objHashtable = new System.Collections.Hashtable();

                    // check to kill the right process
                    foreach (System.Diagnostics.Process processInExcel in objProcess)
                    {
                        if (objHashtable.ContainsKey(processInExcel.Id) == false)
                        {
                            processInExcel.Kill();
                        }
                    }
                    objProcess = null;
                }
                Console.ReadLine();
                #endregion

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.ShowDialog();
                string filePath = openFileDialog.FileName;
                txtCurrPath.Text = filePath;
            }
           
        }

        private void txtGroupCount_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            else if (txtGroupCount.TextLength == 0 && e.KeyChar == '0')
            {
                e.Handled = true;
            }
            else if (txtGroupCount.TextLength < 3 || e.KeyChar == '\b')
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }

            //// If you want, you can allow decimal (float) numbers
            //if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            //{
            //    e.Handled = true;
            //}
        }
    }
}
