using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class DBManage
    {
        public static void createOrInsertCmd(SQLiteConnection conn, string sqlCmd) //CREATE語法或INSERT語法
        {
            try
            {
                conn.Open();
                SQLiteCommand sql = new SQLiteCommand(sqlCmd, conn);
                sql.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {

            }
        }

        public static Queue selectCmd(SQLiteConnection conn, string sqlCmd, string col) //SELECT語法
        {
            Queue returnQueue = new Queue();
            try
            {
                conn.Open();

                using (SQLiteCommand sql = new SQLiteCommand(sqlCmd, conn))
                {
                    using (SQLiteDataReader dataReader = sql.ExecuteReader())
                    {
                        dataReader.NextResult();
                        while (dataReader.Read())
                        {
                            string sqlResult = dataReader[col].ToString();
                            returnQueue.Enqueue(sqlResult);
                        }
                        dataReader.Close();
                    }
                }
                conn.Close();
            }
            catch (Exception ex)
            {

            }

            return returnQueue;
        }

        public static DataTable getTable(SQLiteConnection conn, string sql)
        {
            DataTable dataTable = new DataTable();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);

            conn.Open();
            da.Fill(dataTable);
            conn.Close();
            da.Dispose();
            return dataTable;
        }
    }
}
