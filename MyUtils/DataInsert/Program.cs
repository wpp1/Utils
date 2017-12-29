using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Dml;

namespace DataInsert
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(SqlHelper.GetConnSting());
            //string path = @"F:\AMyTest\MyUtils\DataInsert\Excel\隐患点表.xlsx";
            string path = @"C:\Users\Administrator.MICROSO-8L59UM9\Desktop\隐患点表(7).xlsx";
            DataTable dtExcel = Helper.ExcelHelper.ExcelToDataTable(path, "Sheet3",true);
            for (int i = 0; i < dtExcel.Rows.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("insert into t_PreventionPoint values('"+dtExcel.Rows[i]["隐患点名称"]+"',");
                sb.Append("'" + dtExcel.Rows[i]["责任单位"] + "',");
                sb.Append("'" + dtExcel.Rows[i]["责任联系人"] + "',");
                sb.Append("'" + dtExcel.Rows[i]["责任人电话"] + "',");
                sb.Append("'" + dtExcel.Rows[i]["管理部门"] + "',");
                sb.Append("'" + dtExcel.Rows[i]["管理领导"] + "',");
                sb.Append("'" + dtExcel.Rows[i]["管理领导电话"] + "')");
                SqlHelper.ExecuteNonQuery(SqlHelper.GetConnSting(), CommandType.Text, sb.ToString());
                //自增列id
                string id = SqlHelper.ExecuteDataset(SqlHelper.GetConnSting(), CommandType.Text,
                    "select top 1 ID from t_PreventionPoint order by ID desc").Tables[0].Rows[0][0].ToString();
                string lineName = dtExcel.Rows[i]["影响线路"].ToString();
                if (!string.IsNullOrEmpty(lineName))
                {
                    var line = lineName.Split('、');
                    foreach (var linein in line)
                    {
                        StringBuilder sb2 = new StringBuilder();
                        string lineId = System.Text.RegularExpressions.Regex.Replace(linein, @"[^0-9]+", "")+id;
                        sb2.Append("insert into t_PreventionPointLines values(");
                        sb2.Append("'" + lineId + "',");
                        sb2.Append("'" + linein + "',");
                        sb2.Append("'" + id + "')");
                        SqlHelper.ExecuteNonQuery(SqlHelper.GetConnSting(), CommandType.Text, sb2.ToString());
                    }
                    
                }
                Console.WriteLine(id);
            }


            DataTable dt =  SqlHelper.ExecuteDataset(SqlHelper.GetConnSting(), CommandType.Text, "select * from  t_PreventionPoint").Tables[0];
            Console.WriteLine(dt.Rows.Count);
            Console.ReadLine();
        }
        
    }
}
