using System;
using System.Data;
using DBUtility;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {            
            Console.WriteLine(DbHelperSQL.connectionString);
            DataSet ds = DbHelperSQL.ExcuteDataSet("select * from [dbo].[t_Admin]");
            Console.ReadLine();
        }
    }
}
