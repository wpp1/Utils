using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace Test
{
    class Test
    {
        private static void Test1()
        {
            //1.excel转datatable
            using (NPOIHelper excelHelper = new NPOIHelper(@"F:\隐患点表(6).xlsx"))
            {
                DataTable dt = excelHelper.ExcelToDataTable("Sheet1", true);
                PrintData(dt);
            }

            //2.datatable转excel
            using (NPOIHelper excelHelper = new NPOIHelper(@"F:\test.xlsx"))
            {
                DataTable dt = new DataTable();//创建表
                dt.Columns.Add("ID", typeof(Int32));//添加列
                dt.Columns.Add("Name", typeof(String));
                dt.Columns.Add("Age", typeof(Int32));
                dt.Rows.Add(new object[] { 1, "张三", 20 });//添加行
                dt.Rows.Add(new object[] { 2, "李四", 25 });
                dt.Rows.Add(new object[] { 3, "王五", 30 });
                DataView dv = dt.DefaultView;//获取表视图
                dv.Sort = " ID DESC";//按照ID倒序排序
                dt = dv.ToTable();//转为表
                excelHelper.DataTableToExcel(dt, "s1", true);
            }

            //3.返回字符ascii码字节
            while (true)
            {
                if ((int)ConsoleKey.Enter == 13)
                {
                    string str = Console.ReadLine();
                    Console.WriteLine(StringHelper.GetSubString(str));
                }
            }
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
