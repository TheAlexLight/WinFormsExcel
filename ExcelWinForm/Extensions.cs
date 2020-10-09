using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWinForm
{ 

    static class Extensions
    {
        public static int[] countArr { get; set; }
        public static List<User> userList { get; set; }
        public static List<string[,]> reList { get; set; }
        public static int[,] arrAllMarks { get; set; }

        public static List<List<int>> FindAnotherItem(List<List<int>> list, int wrongItem, int currGroup)
        {
            List<List<int>> list2 = new List<List<int>>(list);

            list2[currGroup].Add(wrongItem);
            return list2;
        }

        public static int BestCount(int[] arr, int length, List<List<int>> list, int currGroup)
        {
            int count = 0;
            //if (list.Count > 0)
            //{
            for (int i = 0; i < list[currGroup].Count; i++)
            {

                foreach (var item in list[currGroup])
                {
                    if (count == item)
                    {
                        count++;
                        break;
                    }
                }
                if (count == length + 1)
                {
                    return -1; // если все поля для этой группы заняты
                }
            }




            bool flag = false;
            for (int i = 0; i < length + 1; i++)
            {
                if (arr[i] != 0)
                {
                    flag = true;
                    break;
                }
            }

            if (flag == false)
            {
                return -2; //если подходящих чисел больше нет
            }

            if (list[currGroup].Count == 0)
            {
                while (true)
                {
                    if (arr[count] == 0)
                    {
                        count++;
                    }
                    else
                    {
                        break;
                    }

                }
            }

            int countId = arr[count];
            for (int i = count; i < length; i++)
            {
                flag = false;
                if (countId > arr[i + 1] && arr[i + 1] != 0)
                {
                    foreach (var item in list[currGroup])
                    {
                        if ((i + 1) == item)
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (flag == false)
                    {
                        countId = arr[i + 1];
                        count = i + 1;
                    }

                }
            }
            if (countId == 0)
            {
                return -2;
            }
            return count;
        }

        public static int BestCount(int[,] arr, int length, List<List<int>> list, int currGroup)
        {
            int count = 0;
            if (list.Count > 0)
            {
                for (int i = 0; i < list[currGroup].Count; i++)
                {

                    foreach (var item in list[currGroup])
                    {
                        if (count == item)
                        {
                            count++;
                            break;
                        }
                    }
                    if (count == length + 1)
                    {
                        return -1; // если все поля для этой группы заняты
                    }
                }
            }

            bool flag = false;

            int countId = arr[currGroup, count];
            for (int i = count; i < length; i++)
            {
                flag = false;
                if (countId > arr[currGroup, i + 1] && arr[currGroup, i + 1] != 0)
                {
                    foreach (var item in list[currGroup])
                    {
                        if ((i + 1) == item)
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (flag == false)
                    {
                        countId = arr[currGroup, i + 1];
                        count = i + 1;
                    }

                }
            }

            return count;
        }

        public static int[] CountMarks(Range lastCell, List<User> userList, int[] countArr, int n, int currValue)
        {
            int count = 0;
            for (int i = 0; i < (int)lastCell.Column - n; i++)
            {
                count = 0;
                foreach (var item in userList)
                {
                    if (item.arr[0, i] == currValue.ToString())
                    {
                        count++;
                    }
                    countArr[i] = count;
                }
            }
            return countArr;
        }

        public static void SortPeople(Range lastCell, Worksheet ObjWorkSheet, List<int[]> groupAmount, List<List<int>> wrongItemList,
                                      int humanPerGroup, int humanCount, int groupCount, int remainder)
        {
            int n = 4;
            int currValue = 3; // текущая оценка
            int currGroup = 0; // текущая группа
            int countI = 0; // номер строки
            bool flag = true; // есть свободные ячейки
            int id = 0; // идентификатор для Excel
            int listId = 0; // идентификатор для List
            int countId = 0; //переменная для подсчета идентификатора
            int changeCount = 0; // переменная для смены группы
            bool flagValue = false; // для count == -2

            while (true)
            {
                //Подсчет количества 1-3 в каждом столбце по всему списку
                countArr = Extensions.CountMarks(lastCell, userList, countArr, n, currValue);

                //выбор следующего столбца
                int count = Extensions.BestCount(countArr, (int)lastCell.Column - n - 1, wrongItemList, currGroup);

                if (count > -1)
                {
                    if ((groupAmount[currGroup][count] < 1 || groupAmount[currGroup][count] < humanPerGroup * 0.33) && reList[currGroup][count, countI] == null)
                    {
                        flag = false;
                        foreach (var item in userList)
                        {
                            if (item.arr[0, count] == currValue.ToString())
                            {
                                id = item.Id;
                                listId = userList.IndexOf(item);
                                flag = true;
                                break;
                            }
                        }
                        if (flag == false)
                        {
                            countId = 0;
                            while (countId < (int)lastCell.Column - n)
                            {
                                foreach (var item2 in userList)
                                {
                                    if (item2.arr[0, countId] == currValue.ToString())
                                    {
                                        id = item2.Id;
                                        listId = userList.IndexOf(item2);
                                        flag = true;
                                        break;
                                    }
                                }
                                if (flag)
                                {
                                    break;
                                }
                                countId++;
                            }
                        }

                        reList[currGroup][count, countI] = ObjWorkSheet.Cells[id + 1, 3].Text.ToString() + $"(№{id + 1})"
                            + " - " + userList[listId].arr[0, count].ToString() + " - " + ObjWorkSheet.Cells[id + 1, 4].Text.ToString();

                        arrAllMarks[currGroup, count] = arrAllMarks[currGroup, count] + Int32.Parse(userList[listId].arr[0, count]);
                        groupAmount[currGroup][count]++;

                        if (groupAmount[currGroup][count] >= humanPerGroup * 0.33)
                        {
                            wrongItemList = Extensions.FindAnotherItem(wrongItemList, count, currGroup);
                        }

                        userList.RemoveAt(listId);
                        if (userList.Count == 0)
                        {
                            break;
                        }
                    }
                    else
                    {
                        wrongItemList = Extensions.FindAnotherItem(wrongItemList, count, currGroup);
                    }
                }
                else
                {
                    if (count == -1) // если все поля для этой группы заняты
                    {
                        if (currGroup == groupCount - 1)
                        {
                            currGroup = 0;
                        }
                        else
                        {
                            currGroup++;
                        }
                        if (changeCount != groupCount - 1)
                        {
                            changeCount++;
                            continue;
                        }
                        else
                        {
                            flag = false;
                        }

                    }
                    else if (count == -2) // если подходящих чисел больше нет
                    {
                        bool checkFlag = false;
                        foreach (var item in reList)
                        {
                            for (int i = 0; i < humanCount / humanPerGroup; i++)
                            {
                                if (item[i, countI] == null)
                                {
                                    if (currValue == 3 || currValue == 2)
                                    {
                                        currValue--;
                                    }
                                    else
                                    {
                                        if (flagValue == false)
                                        {
                                            currValue++;
                                            flagValue = true;
                                        }
                                        else
                                        {
                                            currValue = 3;
                                            flagValue = false;
                                        }

                                    }
                                    checkFlag = true;
                                    flag = true;
                                    changeCount = 0;
                                    break;
                                }
                                if (checkFlag)
                                {
                                    break;
                                }
                            }
                            if (checkFlag)
                            {
                                break;
                            }
                            else
                            {
                                flag = false;
                            }
                        }
                    }
                }

                // если нужно сменить строку
                if (flag == false)
                {
                    changeCount = 0;
                    countI++;
                    flag = true;
                    foreach (var item in wrongItemList)
                    {
                        item.Clear();
                    }

                    if (groupAmount[currGroup][0] >= humanPerGroup * 0.33)
                    {
                        foreach (var item in groupAmount)
                        {
                            for (int i = 0; i < (int)lastCell.Column - n; i++)
                            {
                                item[i] = 0;
                                countArr[i] = 0;

                            }

                        }

                        if (currValue == 3)
                        {
                            currValue = currValue - 2;
                        }
                        else if (currValue == 1)
                        {
                            currValue = currValue + 1;
                        }
                        else
                        {
                            if (userList.Count == 0)
                            {
                                break;
                            }
                        }
                        if (countI == humanPerGroup)
                        {
                            break;
                        }
                    }
                    else if (userList.Count == remainder)
                    {
                        break;
                    }
                }

                if (currGroup == groupCount - 1)
                {
                    currGroup = 0;
                }
                else
                {
                    currGroup++;
                }

            }
        }

        public static void SortRemainder(int remainder_copy, Range lastCell, Worksheet ObjWorkSheet, int n, int groupCount, List<List<int>> wrongItemList, int humanPerGroup)
        {
            int listId = 0;
            int id = 0;
            int currGroup = 0;
            int countJ = 0;
            //int remainder_copy = remainder;

            while (remainder_copy != 0)
            {
                countJ = Extensions.BestCount(arrAllMarks, (int)lastCell.Column - n - 1, wrongItemList, currGroup);

                id = userList[0].Id;
                reList[currGroup][countJ, humanPerGroup] = ObjWorkSheet.Cells[id + 1, 3].Text.ToString() + $"(№{id + 1})" + " - " + userList[0].arr[0, countJ] +
                    " - " + ObjWorkSheet.Cells[id + 1, 4].Text.ToString(); ;

                arrAllMarks[currGroup, countJ] = arrAllMarks[currGroup, countJ] + Int32.Parse(userList[0].arr[0, countJ]);

                wrongItemList = Extensions.FindAnotherItem(wrongItemList, countJ, currGroup);

                userList.RemoveAt(listId);
                remainder_copy = remainder_copy - 1;

                if (currGroup == groupCount - 1)
                {
                    currGroup = 0;
                }
                else
                {
                    currGroup++;
                }
            }
        }

        public static void PrintCatchedMessages(string redStr, string yellowStr)
        {
            Console.ForegroundColor = ConsoleColor.Red; // устанавливаем цвет
            Console.WriteLine(redStr);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(yellowStr);
            Console.ResetColor();
        }
        public static void PrintCatchedMessages(string redStr1, Exception ex, string yellowStr)
        {
            Console.ForegroundColor = ConsoleColor.Red; // устанавливаем цвет
            Console.WriteLine(redStr1);
            Console.WriteLine(ex.Message);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(yellowStr);
            Console.ResetColor();
        }
    }
}
