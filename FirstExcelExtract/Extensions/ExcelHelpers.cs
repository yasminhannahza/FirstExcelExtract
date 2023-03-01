using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelperExtension
{
    public static class ExcelHelpers
    {
        public static List<T> FlattenArray<T>(this Array theData)
        {
            List<T> list = new List<T>();

            foreach (T item in theData)
            {
                list.Add(item);
            }
            return list;
        }
    }
}
