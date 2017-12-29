using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    public static class EnumHelper
    {
        /// <summary>
        /// 根据枚举描述返回key
        /// </summary>
        /// <param name="enumType"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static int GetEnumItemByDescription(this Type enumType, string value)
        {
            if (!enumType.IsEnum)
                throw new InvalidOperationException();
            Type typeDescription = typeof(DescriptionAttribute);
            FieldInfo[] fields = enumType.GetFields();
            int key = 0;
            foreach (FieldInfo field in fields)
            {
                if (!field.FieldType.IsEnum)
                    continue;
                if (((DescriptionAttribute)field.GetCustomAttributes(typeDescription, false)[0]).Description == value)
                {
                    key = (int)enumType.InvokeMember(field.Name, BindingFlags.GetField, null, null, null);
                }
            }
            return key;
        }

        /// <summary>
        /// 获取枚举的description属性
        /// </summary>
        /// <param name="e"> </param>
        /// <returns> </returns>
        public static string GetDescription(System.Enum e, int flag = 0)
        {
            Type t = e.GetType();
            try
            {
                FieldInfo fi = t.GetField(System.Enum.GetName(t, e));
                var attrs = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);
                return (attrs.Length > 0 && flag == 0) ? attrs[0].Description : System.Enum.GetName(t, e);
            }
            catch
            {
                return "";
            }
        }


        /// <summary>
        /// 排序方式
        /// </summary>
        [Description("排序方式")]
        public enum OrderTypeEnum
        {
            /// <summary>
            /// 按日期正序
            /// </summary>
            [Description("按日期正序")]
            Asc = 1,
            /// <summary>
            /// 按日期倒序
            /// </summary>
            [Description("按日期倒序")]
            Desc = 2
        }

    }
}
