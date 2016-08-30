using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Login : Form
    {
        public SQLiteConnection conn = new SQLiteConnection(@"Data Source=" + Constants.LoginDatabaseFile);
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Resize(object sender, EventArgs e)
        {
            this.Width = 819;
            this.Height = 547;
        }

        private void Login_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
            LoginLib.checkLoginDB(Constants.LoginDatabaseFile, conn);
        }

        private void LoginButton_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == LoginLib.checkPassword(conn))
            {
                main frm = new main();
                this.Hide();
                frm.Show();               
            }
            else
            {
                MessageBox.Show("密碼錯誤");
            }
        }
    }

    public static class LoginLib
    {
        public static void checkLoginDB(string dbFile, SQLiteConnection conn)
        {
            if (!File.Exists(dbFile))
            {
                SQLiteConnection.CreateFile(dbFile);
                conn.Open();
                // Create Table
                try
                {
                    String commandText = "CREATE TABLE passwordTable(ID INTEGER PRIMARY KEY AUTOINCREMENT,loginpassword VARCHAR(25))";
                    SQLiteCommand cmd = new SQLiteCommand(commandText, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                // Insert Default Password
                try
                {
                    SQLiteCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "INSERT INTO passwordTable(loginpassword) VALUES ('123')";
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                conn.Close();
            }
        }

        public static String checkPassword(SQLiteConnection conn)
        {
            DataTable returnTable = new DataTable();
            String dbPassword = "";
            conn.Open();
            // Select Password
            try
            {
                SQLiteCommand cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT loginpassword FROM passwordTable";
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                adapter.Fill(returnTable);
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    using (DataTable dt = new DataTable())
                    {
                        dt.Load(dr);
                        DataRow row = dt.Rows[0];
                        dbPassword = Convert.ToString(row["loginpassword"]);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            conn.Close();
            return dbPassword;
        }
    }
}
