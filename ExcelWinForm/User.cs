using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWinForm
{
    [Serializable]
    public class User//: ICloneable
    {
        static int counter = 0;
        public int Id { get; private set; }
        public string[,] arr { get; set; }

        public User(string[,] resArr)
        {
            arr = new string[1, Form1/*Program*/.staticRow];
            arr = resArr;
            Id = ++counter;
        }

        public List<User> DeepCopy()
        {
            List<User> tmpList = new List<User>();

            return tmpList;
        }

        public static List<T> Clone<T>(List<T> items)
        {
            using (var stream = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(stream, items);
                stream.Position = 0;
                return (List<T>)formatter.Deserialize(stream);
            };
        }

        public static void ZeroCounter()
        {
            counter = 0;
        }
    }

}
