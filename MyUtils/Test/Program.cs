using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            using (NPOIHelper excelHelper = new NPOIHelper(@"F:\隐患点表(6).xlsx"))
            {
                DataTable dt = excelHelper.ExcelToDataTable("Sheet1",true);
                PrintData(dt);
            }
            Console.ReadLine();
        }

        static void PrintData(DataTable data)
        {
            if (data == null) return;
            for (int i = 0; i < data.Rows.Count; ++i)
            {
                for (int j = 0; j < data.Columns.Count; ++j)
                    Console.Write("{0} ", data.Rows[i][j]);
                Console.Write("\n");
            }
        }

    }
}
