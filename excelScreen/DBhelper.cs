using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.IO;

namespace excelScreen
{
    internal class DBHelper
    {
        public static string cour = ConfigurationManager.AppSettings["ConnStr"].ToString();
        public int Dbtran(string sqlstr)
        {
            int flag = 0;
            //string cour = ConfigurationManager.AppSettings["ConnStr"].ToString();
            MySqlConnection con = new MySqlConnection(cour);
            if (con.State == ConnectionState.Closed)
                con.Open();
            MySqlTransaction MyTra = con.BeginTransaction();
            try
            {
                MySqlCommand sc = new MySqlCommand(sqlstr, con);
                sc.Transaction = MyTra;
                flag = sc.ExecuteNonQuery();
                MyTra.Commit();
            }
            catch (Exception ex)
            {
                MyTra.Rollback();
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
            finally
            {
                con.Close();
            }
            return flag;
        }

        public int DbExcuteNonQuery(string sqlStr)
        {
            MySqlConnection con = new MySqlConnection(cour);
            MySqlCommand sc = new MySqlCommand(sqlStr, con);
            int flag = 0;
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                flag = sc.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
                flag = -1;
                // return Convert.ToInt32(ex);
            }
            finally
            {
                sc.Dispose();
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
            return flag;
        }

        public int select(string sqlStr)
        {
            MySqlConnection con = new MySqlConnection(cour);
            MySqlCommand sc = new MySqlCommand(sqlStr, con);
            string flag = null;
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                flag = sc.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open)
                    con.Close();

                //return Convert.ToInt32(ex);
            }
            finally
            {
                sc.Dispose();
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
            return Convert.ToInt32(flag);
        }

        public DataSet select1(string sqlStr)
        {

            MySqlConnection con = new MySqlConnection(cour);
            MySqlDataAdapter sda = new MySqlDataAdapter(sqlStr, con);
            DataSet ds = new DataSet();
            string flag = null;
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                sda.Fill(ds);//
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open)
                    con.Close();

                //return Convert.ToInt32(ex);
            }
            finally
            {
                sda.Dispose();
                if (con.State != ConnectionState.Closed)
                    con.Close();
            }

            return ds;
        }

        public int UpSqlData(string sqlStr,DataTable dt)
        {
            MySqlConnection con = new MySqlConnection(cour);
            MySqlDataAdapter da = new MySqlDataAdapter(sqlStr, con);
            MySqlCommandBuilder sqlCmdBld = new MySqlCommandBuilder(da);
            MySqlCommand sc = new MySqlCommand(sqlStr, con);
            if (con.State == ConnectionState.Closed)
                con.Open();
            MySqlTransaction MyTra = con.BeginTransaction();
            int a = 0;
            try
            {
               // sc.Transaction = MyTra;
                a = da.Update(dt);
                MyTra.Commit();
            }
            catch (Exception e)
            {
                MyTra.Rollback();
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
            finally
            {
                da.Dispose();
                sqlCmdBld.Dispose();
                if (con.State != ConnectionState.Closed)
                    con.Close();
            }

            return a;
          
        }

        public ArrayList FileRE(string fileName)
        {
            string line;
            ArrayList count = new ArrayList();
            if (File.Exists(fileName))
            {
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs, System.Text.Encoding.GetEncoding("UTF-8"));
                while ((line = sr.ReadLine()) != null)
                {
                    count.Add(line);
                }
                sr.Close();
            }
            else
            {
                // Directory.CreateDirectory(fileName);
                File.CreateText(@"C:\config.ini");
                //int i = 0;

                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs, System.Text.Encoding.GetEncoding("UTF-8"));
                while ((line = sr.ReadLine()) != null)
                {
                    count.Add(sr.ReadLine());
                    //count[i] = sr.ReadLine();
                }
                sr.Close();
                fs.Close();
            }
            return count;
        }
    }
}