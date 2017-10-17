using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatchUp
{
    using System.Data;
    using System.Reflection;
    using System.Web.Script.Serialization;
    public static class ExtendClass
    {
        static JavaScriptSerializer js = new JavaScriptSerializer();
        public static string GetJsonStr(this object obj)
        {
            return js.Serialize(obj);
        }

        public static T Deserialize<T>(this string jsonStr)
        {
            return js.Deserialize<T>(jsonStr);
        }
        #region 将实体类集合转换为DataTable + static DataTable ToDataTable<T>(this IList<T> list)
        /// <summary>
        ///将实体类集合转换为DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> list)
        {
            Type elementType = typeof(T);

            var t = new DataTable();

            elementType.GetProperties().ToList().ForEach(propInfo => t.Columns.Add(propInfo.Name, Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType));
            foreach (T item in list)
            {
                var row = t.NewRow();
                elementType.GetProperties().ToList().ForEach(propInfo => row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value);
                t.Rows.Add(row);
            }
            return t;
        }
        #endregion
        #region 将DataTable转换为集合，其中T必须是实体类，不能是匿名类，也就是说你必须给一个实体类给它！ + static IList<T> ConvertToModelList<T>(this DataTable dt) where T : new()
        /// <summary>
        /// 将DataTable转换为集合，其中T必须是实体类，不能是匿名类，也就是说你必须给一个实体类给它！
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static IList<T> ConvertToModelList<T>(this DataTable dt) where T : new()
        {
            IList<T> ts = new List<T>();//定义集合
            Type type = typeof(T); //获得此模型的类型
            string tempName = "";
            foreach (DataRow dr in dt.Rows)
            {
                T t = new T();
                PropertyInfo[] propertys = t.GetType().GetProperties();//获得此模型的公共属性
                foreach (PropertyInfo pi in propertys)
                {
                    tempName = pi.Name;
                    if (dt.Columns.Contains(tempName))
                    {
                        if (!pi.CanWrite)
                            continue;
                        object value = dr[tempName];
                        if (value != DBNull.Value)
                        {
                            pi.SetValue(t, value, null);
                        }
                    }
                }
                ts.Add(t);
            }
            return ts;
        }
        #endregion
    }
}
