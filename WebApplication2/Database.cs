//using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace WebApplication2
{
    public class Database
    {
        public static String ConnectionString = "Server=\".\\SQLEXPRESS\";Initial Catalog=\"test2\";User ID=\"sa\";Password=\"Usr123456\"";

        //public static String ConnectionString = "Server=\"tornado-soft.com\";Initial Catalog=\"le-garage\";User ID=\"le-garage\";Password=\"Usr123456\"";


        public static DataTable ReadTable(String tblName, String c, List<SqlParameter> lstParams, out string errMessage)
        {
            SqlDataAdapter adp = new SqlDataAdapter("SELECT * FROM " + tblName + " " + c, ConnectionString);
            if (lstParams != null) adp.SelectCommand.Parameters.AddRange(lstParams.ToArray());
            DataTable tbl = new DataTable();
            try
            {
                adp.Fill(tbl);
                errMessage = "";
                return tbl;
            }
            catch (Exception ex) {
                errMessage = ex.Message;
                ////Tools.SaveError(ex.Message);
                return null;
            }
        }
        public static bool DeleteRow(string tblName, List<String> cols, List<Object> vals, out string errMessage)
        {
            errMessage = string.Empty;

            if (vals == null || cols == null) return false;
            if (vals.Count == 0 || cols.Count == 0) return false;
            try
            {
                SqlConnection cn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("", cn);

                string colValStr = "";

                for (int oCounter = 0; oCounter <= cols.Count - 1; oCounter++)
                {
                    colValStr += cols[oCounter] + " = " + "@param" + (oCounter + 1).ToString();

                    if (oCounter != cols.Count - 1)
                    {
                        colValStr += " AND ";
                    }

                    cmd.Parameters.AddWithValue("@param" + (oCounter + 1).ToString(), vals[oCounter]);
                }

                string qStr = "DELETE FROM " + tblName + " WHERE " + colValStr;
                cmd.CommandText = qStr;

                cn.Open();
                int cnt = cmd.ExecuteNonQuery();
                cn.Close();

                if (cnt > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                errMessage = ex.Message;
                return false;
            }
        }

        public static bool DeleteRow(String tblName, Guid ID, out string errMessage)
        {
            try
            {
                errMessage = "";
                SqlConnection cn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("", cn);

                string qStr = "DELETE FROM " + tblName + " WHERE ID = @id";
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.CommandText = qStr;

                cn.Open();
                int cnt = cmd.ExecuteNonQuery();
                cn.Close();

                if (cnt > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                errMessage = ex.Message;
                //Tools.SaveError(ex.Message);
                return false;
            }
        }

        public static int RowCount(string tblName)
        {
            SqlConnection cn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM " + tblName, cn);

            try
            {
                cn.Open();
                int cnt = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                cn.Close();
                return cnt;
            }
            catch (Exception ex)
            {
                //Tools.SaveError(ex.Message);
                return 0;
            }
        }


        public static DataTable ReadTable(String tblName, [Optional] out string errrMessage)
        {
            errrMessage = "";
            SqlDataAdapter adp = new SqlDataAdapter("SELECT * FROM " + tblName, ConnectionString);

            DataTable tbl = new DataTable();
            try
            {
                adp.Fill(tbl);
                return tbl;
            }
            catch (Exception ex) {
                errrMessage = ex.Message;
                //Tools.SaveError(ex.Message);
                return null;
            }
        }
        public static DataTable ReadTableByQuery(String query, List<SqlParameter> lstPArams, [Optional] out string errrMessage)
        {
            errrMessage = "";
            SqlDataAdapter adp = new SqlDataAdapter(query, ConnectionString);
            if (lstPArams != null) adp.SelectCommand.Parameters.AddRange(lstPArams.ToArray());

            DataTable tbl = new DataTable();
            try
            {
                
                adp.Fill(tbl);
                adp.SelectCommand.Parameters.Clear();
                return tbl;
            }
            catch (Exception ex)
            {
                errrMessage = ex.Message;
                //Tools.SaveError(ex.Message);
                return null;
            }
        }


        public static bool UpdateRow(string tblName, Guid ID, List<String> cols, List<Object> vals, out string errMessage)
        {
            errMessage = string.Empty;
            if (vals == null || cols == null) return false;
            if (vals.Count == 0 || cols.Count == 0) return false;
            try
            {
                SqlConnection cn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("", cn);

                string colValStr = "";

                for (int oCounter = 0; oCounter <= cols.Count - 1; oCounter++)
                {
                    colValStr += cols[oCounter] + " = " + "@param" + (oCounter + 1).ToString();

                    if (oCounter != cols.Count - 1)
                    {
                        colValStr += " , ";
                    }

                    cmd.Parameters.AddWithValue("@param" + (oCounter + 1).ToString(), vals[oCounter]);
                }

                string qStr = "Update " + tblName + " SET " + colValStr + " WHERE ID = @id";
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.CommandText = qStr;

                cn.Open();
                int cnt = cmd.ExecuteNonQuery();
                cn.Close();

                if (cnt > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                errMessage = ex.Message;
                //Tools.SaveError(ex.Message);
                return false;
            }
        }

        public async static Task<bool> InsertRow(string tblName, Guid ID, List<String> cols, List<Object> vals)
        {
            string errMessage = string.Empty;

            if (vals == null || cols == null) return false;
            if (vals.Count == 0 || cols.Count == 0) return false;
            try
            {
                SqlConnection cn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("", cn);

                string colStr = "id";
                string valStr = "@id";
                cmd.Parameters.AddWithValue("@id", ID);

                for (int oCounter = 0; oCounter <= cols.Count - 1; oCounter++)
                {
                    colStr += ",";
                    valStr += ",";

                    colStr += cols[oCounter];
                    valStr += "@param" + (oCounter + 1).ToString();

                    cmd.Parameters.AddWithValue("@param" + (oCounter + 1).ToString(), vals[oCounter]);
                }

                string qStr = "INSERT INTO " + tblName + " (" + colStr + ") VALUES (" + valStr + ")";
                cmd.CommandText = qStr;

                await cn.OpenAsync();
                int cnt = await cmd.ExecuteNonQueryAsync();
                cn.Close();

                if (cnt > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                errMessage = ex.Message;
                return false;
            }
        }

        public static DataRow GetRow(String tblName, Guid ID)
        {
            SqlDataAdapter adp = new SqlDataAdapter("SELECT * FROM " + tblName + " WHERE ID = @id", ConnectionString);
            adp.SelectCommand.Parameters.Add(new SqlParameter("@id", ID));
            DataTable tbl = new DataTable();
            adp.Fill(tbl);
            if (tbl != null && tbl.Rows.Count > 0)
            {
                return tbl.Rows[0];
            }
            else
            {
                return null;
            }
        }

        public static DataRow FindRow(string tblName, string colName, object v)
        {
            try
            {
                SqlDataAdapter adp = new SqlDataAdapter("SELECT * FROM " + tblName + " WHERE " + colName + " LIKE @v", ConnectionString);
                adp.SelectCommand.Parameters.Add(new SqlParameter("@v", v));
                DataTable tbl = new DataTable();
                adp.Fill(tbl);
                if (tbl != null && tbl.Rows.Count > 0)
                {
                    return tbl.Rows[0];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                //Tools.SaveError(ex.Message);
                return null;
            }

        }



        public static bool check_rating(string tblName, string colName, string colName2, string user_id, string src_id)
        {
            try
            {
                SqlDataAdapter adp = new SqlDataAdapter($"SELECT {colName} FROM {tblName}  WHERE EXISTS (SELECT {colName} FROM {tblName}  WHERE {colName} = @value and {colName2} = {src_id})", ConnectionString);
                adp.SelectCommand.Parameters.Add(new SqlParameter("@user_id", user_id));
                adp.SelectCommand.Parameters.Add(new SqlParameter("@src_id", src_id));
                DataTable tbl = new DataTable();
                adp.Fill(tbl);
                if (tbl != null && tbl.Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                //Tools.SaveError(ex.Message);
                return false;
            }

        }

        public static bool checkRating(string user_id, string src_id)
        {
            try
            {
                //Database.checkRating(user_id.ToString(), src_id.ToString())
                SqlDataAdapter adp = new SqlDataAdapter($"SELECT COUNT(*) FROM rating  WHERE user_id = @user_id and src_id = @src_id", ConnectionString);
                adp.SelectCommand.Parameters.Add(new SqlParameter("@user_id", user_id));
                adp.SelectCommand.Parameters.Add(new SqlParameter("@src_id", src_id));
                DataTable tbl = new DataTable();
                adp.Fill(tbl);
                if (tbl != null && tbl.Rows[0]?[0].ToString() == "0")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                //Tools.SaveError(ex.Message);
                return false;
            }

        }


        public static object ReadValue(string tblName, string colName, Guid id)
        {
            SqlConnection cn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand("SELECT " + colName + " FROM " + tblName + " WHERE ID = @id", cn);
            cmd.Parameters.AddWithValue("@id", id);

            try
            {
                cn.Open();
                object prop_value = cmd.ExecuteScalar();
                cn.Close();
                return prop_value;
            }
            catch
            {
                if (cn.State == ConnectionState.Open) cn.Close();
                return null;
            }

        }

        public static int Get_Users_Count()
        {
            SqlConnection cn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand("SELECT  Counter FROM  Counter_TBL", cn);
            try
            {
                cn.Open();

                object prop_value = cmd.ExecuteScalar();
                cn.Close();
                return Convert.ToInt32(prop_value);
            }
            catch (Exception ex)
            {
                if (cn.State == ConnectionState.Open) cn.Close();
                return 0;
            }

        }

        public static bool Set_Users_Count(int count)
        {
            SqlConnection cn = new SqlConnection(ConnectionString);
            string command = count == 1 ? "INSERT INTO  Counter_TBL (Counter) VALUES(1)" : "UPDATE Counter_TBL SET Counter = " + count;
            SqlCommand cmd = new SqlCommand(command, cn);
            try
            {
                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
                return true;
            }
            catch
            {
                if (cn.State == ConnectionState.Open) cn.Close();
                return false;
            }

        }
        public static String ReadProp(int id)
        {
            SqlConnection cn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand("SELECT Prop FROM Prop WHERE ID = @id", cn);
            cmd.Parameters.AddWithValue("@id", id);

            String v = "";

            try
            {
                cn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.Read())
                {
                    v = rd["Prop"].ToString();
                }
                rd.Close();
                cn.Close();
                return v;
            }
            catch (Exception ex)
            {
                if (cn.State == ConnectionState.Open) cn.Close();
                //Tools.SaveError(ex.Message);
                return v;
            }

        }
        public static bool WriteProp(int id, String v, out string errMessage)
        {
            SqlConnection cn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand("IF NOT EXISTS (SELECT ID FROM Prop WHERE ID = @id) INSERT INTO Prop (ID, Prop) VALUES (@id, @v) ELSE UPDATE Prop SET Prop = @v WHERE ID = @id", cn);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.Parameters.AddWithValue("@v", v);
            try
            {
                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
                errMessage = "";
                return true;
            }
            catch (Exception ex)
            {
                if (cn.State == ConnectionState.Open) cn.Close();
                errMessage = ex.Message;
                //Tools.SaveError(ex.Message);
                return false;
            }

        }
        public static DataTable ConverSQLQueryPage(string sql, List<SqlParameter> lstPArams, string order_col, int page_number, int per_page_number, out string msg, out int count)
        {
            count = 0;
            string countquery = $"SELECT COUNT(*) COUNTER FROM ({ sql}) AS C ";
           // List<SqlParameter> lstPArams1 = new List<SqlParameter>();
           // lstPArams1.AddRange(lstPArams.ToArray());
           
            DataTable data = ReadTableByQuery(countquery, lstPArams, out msg);
            if (data != null)
            {
                var counter = data.Rows[0]["COUNTER"];
                count = Convert.ToInt32(counter);
            }
            string Sql_Statment = "DECLARE @page int = " + page_number + " ";
            Sql_Statment += "DECLARE @per_page int = " + per_page_number + " ";
            Sql_Statment += "SELECT * FROM (   " + sql + " ) AS t  ";
            Sql_Statment += "ORDER BY " + order_col + " OFFSET @per_page * (@page - 1) ROWS ";
            Sql_Statment += " FETCH NEXT @per_page ROWS ONLY";
            DataTable d1 = ReadTableByQuery(Sql_Statment, lstPArams, out msg);
            //count = d1.Rows.Count;
            return d1;

        }
    }

}
