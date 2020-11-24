using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;  //Configuration表面配置，组态，构造

namespace DBOper
{
    /// <summary>
    /// SQLHelper数据库操作类
    /// </summary>
    class SQLHelper
    {
        public static string GetSqlConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["ConStr"].ConnectionString;
        }

        public static string GetGUID()
        {
            return Guid.NewGuid().ToString().Replace("-", "").ToUpper();
        }

        //适合增删改操作，返回影响条数（2个参数）
        public static int ExecuteNonQuery(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(GetSqlConnectionString()))
            {
                using (SqlCommand comm = conn.CreateCommand())
                {
                    conn.Open();
                    comm.CommandText = sql;
                    comm.Parameters.AddRange(parameters);
                    return comm.ExecuteNonQuery();
                }
            }
        }

        //适合增删改操作，返回影响条数（1个参数）
        //public static int ExecuteNonQuery(string sql)
        //{
        //    using (SqlConnection conn = new SqlConnection(GetSqlConnectionString()))
        //    {
        //        using (SqlCommand comm = conn.CreateCommand())
        //        {
        //            conn.Open();
        //            comm.CommandText = sql;
        //            return comm.ExecuteNonQuery();
        //        }
        //    }
        //}

        //查询操作，返回查询结果中的第一行第一列的值（2个参数）
        public static object ExecuteScalar(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(GetSqlConnectionString()))
            {
                using (SqlCommand comm = conn.CreateCommand())
                {
                    conn.Open();
                    comm.CommandText = sql;
                    comm.Parameters.AddRange(parameters);
                    return comm.ExecuteScalar();
                }
            }
        }

        //查询操作，返回查询结果中的第一行第一列的值（1个参数）
        //public static object ExecuteScalar(string sql)
        //{
        //    using (SqlConnection conn = new SqlConnection(GetSqlConnectionString()))
        //    {
        //        using (SqlCommand comm = conn.CreateCommand())
        //        {
        //            conn.Open();
        //            comm.CommandText = sql;
        //            return comm.ExecuteScalar();
        //        }
        //    }
        //}

        //Adapter调整，查询操作，返回DataTable（2个参数）
        public static DataTable ExecuteDataTable(string sql, params SqlParameter[]  parameters)
        {
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, GetSqlConnectionString()))
            {
                DataTable dt = new DataTable();
                adapter.SelectCommand.Parameters.AddRange(parameters);
                adapter.Fill(dt);
                return dt;
            }
        }

        //Adapter调整，查询操作，返回DataTable（1个参数）
        //public static DataTable ExecuteDataTable(string sql)
        //{
        //    using (SqlDataAdapter adapter = new SqlDataAdapter(sql, GetSqlConnectionString()))
        //    {
        //        DataTable dt = new DataTable();
        //        adapter.Fill(dt);
        //        return dt;
        //    }
        //}

        public static SqlDataReader ExecuteReader(string sqlText, params SqlParameter[] parameters)
        {
            //SqlDataReader要求，它读取数据的时候有，它独占它的SqlConnection对象，而且SqlConnection必须是Open状态
            SqlConnection conn = new SqlConnection(GetSqlConnectionString());//不要释放连接，因为后面还需要连接打开状态
            SqlCommand cmd = conn.CreateCommand();
            conn.Open();
            cmd.CommandText = sqlText;
            cmd.Parameters.AddRange(parameters);
            //CommandBehavior.CloseConnection当SqlDataReader释放的时候，顺便把SqlConnection对象也释放掉
            return cmd.ExecuteReader(CommandBehavior.CloseConnection);
        }
    }

    /// <summary>
    /// C#SQL注入攻击检查类
    /// </summary>
    public static class SQLInjection
    {
        private const string StrKeyWord = @"select|insert|delete|count(|drop table|update|truncate|asc(|mid(|char(|xp_cmdshell|exec master|netlocalgroup administrators|net user|""|'";

        /// <summary>
        /// 检查文本是否包含SQL关键字
        /// </summary>
        /// <param name="content">被检查的字符串</param> 
        /// <returns>存在SQL关键字返回true，不存在返回false</returns> 
        private static bool CheckKeyWord(string content)
        {
            string word = content;
            string[] patten1 = StrKeyWord.Split('|');
            foreach (string i in patten1)
            {
                if (word.Contains(" " + i) || word.Contains(i + " "))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 检查文本是否注入攻击
        /// </summary>
        /// <param name="content">被检查的字符串</param>
        /// <returns></returns>
        public static bool IsAttack(string content)
        {
            if (String.IsNullOrWhiteSpace(content)) return false;

            //存在单引号且包含SQL命令
            return (content.Contains("'") || CheckKeyWord(content));
        }

        /// <summary>
        /// 移除SQL命令及单引号
        /// </summary>
        /// <param name="content">被检查的字符串</param>
        /// <returns></returns>
        public static string RemoveKeywords(string content)
        {
            if (String.IsNullOrWhiteSpace(content)) return "";

            //替换高危险单引号
            content = content.Replace("'", "");

            string[] patten1 = StrKeyWord.Split('|');
            foreach (string i in patten1)
            {
                content = content.Replace(i, "");
            }

            return content;
        }
    }

    /// <summary>
    /// DataTableToModel 的摘要说明
    /// </summary>
    public static class DataTableToModel
    {
        /// <summary>
        /// DataTable通过反射获取单个对象
        /// </summary>
        public static T ToSingleModel<T>(this DataTable data) where T : new()
        {
            T t = data.GetList<T>(null, true).Single();
            return t;
        }

        /// <summary>
        /// DataTable通过反射获取单个对象
        /// <param name="prefix">前缀</param>
        /// <param name="ignoreCase">是否忽略大小写，默认不区分</param>
        /// </summary>
        public static T ToSingleModel<T>(this DataTable data, string prefix, bool ignoreCase = true) where T : new()
        {
            T t = data.GetList<T>(prefix, ignoreCase).Single();
            return t;
        }

        /// <summary>
        /// DataTable通过反射获取多个对象
        /// </summary>
        /// <typeparam name="type"></typeparam>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<T> ToListModel<T>(this DataTable data) where T : new()
        {
            List<T> t = data.GetList<T>(null, true);
            return t;
        }

        /// <summary>
        /// DataTable通过反射获取多个对象
        /// </summary>
        /// <param name="prefix">前缀</param>
        /// <param name="ignoreCase">是否忽略大小写，默认不区分</param>
        /// <returns></returns>
        private static List<T> ToListModel<T>(this DataTable data, string prefix, bool ignoreCase = true) where T : new()
        {
            List<T> t = data.GetList<T>(prefix, ignoreCase);
            return t;
        }

        private static List<T> GetList<T>(this DataTable data, string prefix, bool ignoreCase = true) where T : new()
        {
            List<T> t = new List<T>();
            int columnscount = data.Columns.Count;
            if (ignoreCase)
            {
                for (int i = 0; i < columnscount; i++)
                    data.Columns[i].ColumnName = data.Columns[i].ColumnName.ToUpper();
            }
            try
            {
                var properties = new T().GetType().GetProperties();
                var rowscount = data.Rows.Count;
                for (int i = 0; i < rowscount; i++)
                {
                    var model = new T();
                    foreach (var p in properties)
                    {
                        var keyName = prefix + p.Name + "";
                        if (ignoreCase)
                            keyName = keyName.ToUpper();
                        for (int j = 0; j < columnscount; j++)
                        {
                            if (data.Columns[j].ColumnName == keyName && data.Rows[i][j] != null)
                            {
                                string pval = data.Rows[i][j].ToString();
                                if (!string.IsNullOrEmpty(pval))
                                {
                                    try
                                    {
                                        // We need to check whether the property is NULLABLE
                                        if (p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                        {
                                            p.SetValue(model, Convert.ChangeType(data.Rows[i][j], p.PropertyType.GetGenericArguments()[0]), null);
                                        }
                                        else
                                        {
                                            p.SetValue(model, Convert.ChangeType(data.Rows[i][j], p.PropertyType), null);
                                        }
                                    }
                                    catch (Exception x)
                                    {
                                        throw x;
                                    }
                                }
                                break;
                            }
                        }
                    }
                    t.Add(model);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return t;
        }
    }
}
