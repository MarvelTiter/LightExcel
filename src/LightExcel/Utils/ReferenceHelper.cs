using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Utils
{
#if DEBUG
    public class ReferenceHelper
#else
    internal class ReferenceHelper
#endif
    {
        const string AZ = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        /// <summary>
        /// (1, 1) => A1 , (2, 2) => B2
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static string ConvertXyToCellReference(int x, int y)
        {
            return $"{ConvertX(x)}{y}";
        }

        public static string ConvertX(int x)
        {
            /**
             * 坐标转换 CellRef (A1,B2,..
             * A=>1, B=>2, C=>3 ....  AA=>27
             * 10进制转26进制
             */
            string xName = String.Empty;
            int mod;
            while (x > 0)
            {
                mod = x % 26;
                if (mod == 0)
                {
                    mod += 26;
                }
                xName = AZ[mod] + xName;
                x = (x - mod) / 26;
            }
            return xName;
        }

        public static (int? X, int? Y) ConvertCellReferenceToXY(string? cellref)
        {
            if (cellref == null) return (null, null);
            var x = GetColumnIndex(cellref);
            var y = GetRowIndex(cellref);
            return (x, y);
        }

        private static int GetRowIndex(string cellref)
        {
            var num = string.Empty;
            foreach (var c in cellref)
            {
                if (Char.IsNumber(c))
                {
                    num += c;
                }
            }
            return int.Parse(num);
        }

        private static int GetColumnIndex(string cellref)
        {
            var x = 0;
            foreach (var c in cellref)
            {
                if (!Char.IsLetter(c))
                {
                    break;
                }
                x = x * 26 + AZ.IndexOf(c);
            }
            return x;
        }
    }
}
