using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrinityMIB
{
    class DBConnection
    {
        #region 公共变量 
        public SqlConnection conn;
        public  string Server = ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "SQL Settings", "Server", "");

        public  string DataBase = ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "SQL Settings", "DataBase", "");

        public  string Uid = ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "SQL Settings", "Uid", "");

        public  string Pwd = ClassLibrary.ReadFilesClassLibrary.ReadINIFiles.INIGetStringValue(".\\AppConfig.ini", "SQL Settings", "Pwd", "");
    
        #endregion

        #region 数据库方法

        #region 数据库连接和关闭
       ///功能： 数据库连接
        public SqlConnection CreateConection()
        {
            string strcon = "Data Source='" + Server + "'; Initial CataLog='" + DataBase + "'; Uid='" + Uid + "';Pwd='" + Pwd + "'";
            conn = new SqlConnection(strcon);                                                                                 //数据库连接
            conn.Open();
            return conn;
        }
      
        /// 功能：关闭数据库
        public void CloseConnection()
        {
            if (conn.State == ConnectionState.Open)                                                                         //判断最近的数据库是否打开
            {
                conn.Close();                                                                                                                        //关闭数据库
                conn.Dispose();                                                                                                                   //释放所有数据库资源
            }
        }
        #endregion

        

        #region 只读的方式读取数据库信息      
        public SqlDataReader GetCom(string SQLstr)
        {
            CreateConection();                                                                               //打开数据库连接
            //创建一个SqlCommand对象，用户执行SQL语句
            SqlCommand my_com = conn.CreateCommand();
            my_com.CommandText = SQLstr;                                                   //获取指定的SQL语句
            SqlDataReader My_read = my_com.ExecuteReader();           //执行SQL语句，生成一个SqlDataReader对象
            return My_read;
        }
        #endregion

        #region 数据库查找     
        public DataTable GetData(string strSql, out string MSG)
        {
            MSG = string.Empty;
            try
            {
                CreateConection();                                                                                                                                                                                                                                             //获取连接
                SqlCommand com = new SqlCommand(strSql, conn);                                                                                                                                                                       //创建一个SqlCommand对象
                SqlDataAdapter da = new SqlDataAdapter(com);                                                                                                                                                                                   //创建适配器
                DataTable dt = new DataTable();                                                                                                                                                                                                                  //创建DataTable容器
                da.Fill(dt);                                                                                                                                                                                                                                                            //填充容器句柄
                da.Dispose();                                                                                                                                                                                                                                                      //释放适配器资源
                CloseConnection();                                                                                                                                                                                                                                          //关闭连接
                return dt;
            }
            catch (Exception ex)
            {
                MSG = "数据库错误信息" + ex.Message;
                return null;
            }
        }


        /// 功能：判断是否含有数据    
        public bool isDBContainValue(string strSql)
        {
            CreateConection();
            SqlCommand com = new SqlCommand(strSql, conn);
            SqlDataReader sdr = com.ExecuteReader();
            sdr.Read();
            if (sdr.HasRows)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region 数据库更新
        public bool UpdateDB(string SqlStr)
        {
            CreateConection();
            SqlCommand com = new SqlCommand(SqlStr, conn);
            int i = com.ExecuteNonQuery();
            if (i > 0)
            {
                CloseConnection();
                return true;
            }
            else
            {
                CloseConnection();
                return false;
            }
        }
        #endregion

        #region 对数据库进行添加、删除、修改的操作
        /// <summary>
        /// 功能：对数据库进行添加、删除、修改的操作
        /// </summary>
        /// <param name="SQLstr">SQL语句</param>
        public void GetSqlCom(string SQLstr)
        {
            CreateConection();                                                                                            //打开与数据库的连接                                                                                                                    //创建一个SqlCommand对象，用于执行SQL语句
            SqlCommand SQLcom = new SqlCommand(SQLstr, conn);
            SQLcom.ExecuteNonQuery();                                                    //执行SQL语句，返回受影响行数
            SQLcom.Dispose();                                                                        //释放所有空间
            CloseConnection();                                                                                   //调用关闭数据库连接方法，关闭数据库连接 
        }

        #region 数据库插入

        public bool SaveDate(string strSql)
        {
            bool savaResult = false;
            try
            {
                //CreateConection();
                //conn.Open();
                SqlCommand com = new SqlCommand(strSql, conn);
                int i = com.ExecuteNonQuery();
                if (i != 0)
                {
                    savaResult = true;
                }
                 
            }
            catch
            {

            }
            return savaResult;
        }
        #endregion


        #region 数据库更新
        public bool UpData(string SqlStr, out string MSG)
        {
            MSG = string.Empty;
            try
            {
                CreateConection();
                SqlCommand com = new SqlCommand(SqlStr, conn);
                int i = com.ExecuteNonQuery();
                if (i > 0)
                {
                    CloseConnection();
                    return true;
                }
                else
                {
                    CloseConnection();
                    return false;
                }
            }
            catch (Exception ex)
            {
                CloseConnection();
                MSG = "数据库更新失败" + ex.Message;
                return false;
            }
        }
        #endregion
        #endregion
        #endregion
    }
}
