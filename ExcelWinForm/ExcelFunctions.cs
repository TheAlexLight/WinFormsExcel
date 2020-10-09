using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWinForm
{
    public static class ExcelFunctions
    {
        public static Range lastCell { get; set; }
        public static Worksheet ObjWorkSheet { get; set; }
        public static List<User> userList = new List<User>();
        public static string[] programNames;

        public static void ReadExcelData(Range Lastsell, Worksheet ObjWorksheet, int n)
        {
            
            lastCell = Lastsell;
            ObjWorkSheet = ObjWorksheet;
            

            for (int i = 1; i < (int)lastCell.Row; i++) //по всем рядкам
            {
                string[,] list = new string[1, (int)lastCell.Column - n]; // создаем разные массивы
                for (int j = n; j < (int)lastCell.Column; j++) // по всем колонкам
                {
                    list[0, j - n] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();//считываем текст в строку
                }
                userList.Add(new User(list)); // добавляем в список записи
            }
        }

        public static void ReadExcelNames(Range Lastsell, Worksheet ObjWorksheet, string[] programnames, int n)
        {
            lastCell = Lastsell;
            ObjWorkSheet = ObjWorksheet;
            programNames = programnames;

            for (int i = 0; i < (int)lastCell.Column - n; i++)
            {
                programNames[i] = ObjWorkSheet.Cells[1, i + 5].Text.ToString();
            }
        }

        public static int ChangeMarks(Range Lastsell, List<User> UserList, int n)
        {
            lastCell = Lastsell;
            userList = UserList;

            int count = 0;
            for (int i = 0; i < (int)lastCell.Column - n; i++)
            {
                int countPos = 0;
                foreach (var item in userList)//sortedUsers)
                {
                    bool result = Int32.TryParse(userList[countPos].arr[0, count], out int number);

                    if (result == false)
                    {

                        Console.ForegroundColor = ConsoleColor.Red; // устанавливаем цвет
                        Console.WriteLine($"Error in the Excel file: Wrong data in the cell [{countPos + 1}, {i + n + 1}]. Check it and try again");
                        Console.ForegroundColor = ConsoleColor.Yellow; ;
                        Console.WriteLine("Press Enter to end the program");
                        Console.ResetColor();
                        return -1;
                    }

                    // sortedList[countPos].arr[0, count] = item.arr[0, count];
                    if (Int32.Parse(userList[countPos].arr[0, count]) > 7 && Int32.Parse(userList[countPos].arr[0, count]) < 11)
                    {
                        userList[countPos].arr[0, count] = "3";
                    }
                    else if (Int32.Parse(userList[countPos].arr[0, count]) > 3 && Int32.Parse(userList[countPos].arr[0, count]) < 8)
                    {
                        userList[countPos].arr[0, count] = "2";
                    }
                    else userList[countPos].arr[0, count] = "1";
                    countPos++;
                }
                count++;
            }
            return 0;
        }
    }
}
