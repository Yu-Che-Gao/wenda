using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class 更改密碼 : Form
    {
        public 更改密碼()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            新增紀錄表 frm2 = new 新增紀錄表();
            this.Hide();
            frm2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            查詢 frm3 = new 查詢();
            frm3.FormClosed += new FormClosedEventHandler(frm3_FormClosed);

            每月統計 frm4 = new 每月統計();
            frm4.FormClosed += new FormClosedEventHandler(frm4_FormClosed);

            更改密碼 frm5 = new 更改密碼();
            frm5.FormClosed += new FormClosedEventHandler(frm5_FormClosed);

            frm3.Show();
            this.Hide();
        }

        void frm3_FormClosed(object sender, FormClosedEventArgs e)
        {
            新增紀錄表 frm2 = new 新增紀錄表();
            frm2.Show();
        }

        void frm4_FormClosed(object sender, FormClosedEventArgs e)
        {
            新增紀錄表 frm2 = new 新增紀錄表();
            frm2.Show();
        }

        void frm5_FormClosed(object sender, FormClosedEventArgs e)
        {
            新增紀錄表 frm2 = new 新增紀錄表();
            frm2.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            查詢 frm3 = new 查詢();
            frm3.FormClosed += new FormClosedEventHandler(frm3_FormClosed);

            每月統計 frm4 = new 每月統計();
            frm4.FormClosed += new FormClosedEventHandler(frm4_FormClosed);

            更改密碼 frm5 = new 更改密碼();
            frm5.FormClosed += new FormClosedEventHandler(frm5_FormClosed);

            frm4.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == getOld())
            {
                if (textBox2.Text == textBox3.Text)
                {
                    insertNew(textBox2.Text);
                    新增紀錄表 frm2 = new 新增紀錄表();
                    this.Hide();
                    frm2.Show();
                }
                else
                {
                    MessageBox.Show("密碼確認輸入錯誤");
                }
            }
            else
            {
                MessageBox.Show("舊密碼輸入錯誤");
            }

            
        }

        private string getOld()
        {
            SQLiteConnection conn = new SQLiteConnection(@"Data Source=database.dat");
            DataTable returnTable = new DataTable();
            string password1 = "";
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT password1 FROM passwordTable";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            using (SQLiteDataReader dr = cmd.ExecuteReader())
            {
                using (DataTable dt = new DataTable())
                {
                    dt.Load(dr);
                    for (int i1 = 0; i1 < dt.Rows.Count; i1++)
                    {
                        DataRow row = dt.Rows[i1];
                        password1 = Convert.ToString(row["password1"]);
                    }
                }
            }

            conn.Close();
            return password1;

        }

        private void insertNew(string newPassword)
        {
            SQLiteConnection conn = new SQLiteConnection(@"Data Source=database.dat");
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE passwordTable SET password1='"+newPassword+"'";
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void 更改密碼_Load(object sender, EventArgs e)
        {
            button4.Enabled = false;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
