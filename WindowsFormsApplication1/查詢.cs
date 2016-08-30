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
    public partial class 查詢 : Form
    {
        public SQLiteConnection conn = new SQLiteConnection(@"Data Source=database1.dat");

        public 查詢()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("全部");
            comboBox1.Items.Add("僅顯示受傷狀況");
            comboBox1.Items.Add("僅顯示耗材狀況");
            comboBox1.SelectedItem="全部";
            DataTable dt = get("全部");
            dataGridView1.DataSource = dt;
            dataGridView1.Height = Screen.PrimaryScreen.WorkingArea.Height - 100;
            dataGridView1.Width = Screen.PrimaryScreen.WorkingArea.Width;
            button2.Enabled = false;
        }

        private DataTable get(String selectitem)
        {
            DataTable returnTable = new DataTable();
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            if (selectitem == "全部")
            {
                cmd.CommandText = "SELECT * FROM member";
                
            }
            else if (selectitem == "僅顯示受傷狀況")
            {
                cmd.CommandText = "SELECT member.日期,schoolTeam.校名,teamMember.隊伍名,member.隊員,member.受傷部位,member.傷側,member.受傷種類,member.受傷分類,member.處置,member.備註 FROM schoolTeam,teamMember,member where schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員";
            }
            else if (selectitem == "僅顯示耗材狀況")
            {
                cmd.CommandText = "SELECT 日期,schoolTeam.校名,teamMember.隊伍名,member.隊員,member.白貼0_5,member.白貼1_5,member.輕彈1,member.輕彈2,member.輕彈3,member.強彈1,member.強彈2,member.強彈3,member.墊片1_4,member.機能貼布,member.KG3,member.膠膜,member.內膜,member.備註 FROM schoolTeam,teamMember,member where schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員";
            }
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            cmd.ExecuteNonQuery();
            conn.Close();
            return returnTable;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable dt2 = get(comboBox1.Text);
            dataGridView1.DataSource = dt2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
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

        private void button4_Click(object sender, EventArgs e)
        {
            查詢 frm3 = new 查詢();
            frm3.FormClosed += new FormClosedEventHandler(frm3_FormClosed);

            每月統計 frm4 = new 每月統計();
            frm4.FormClosed += new FormClosedEventHandler(frm4_FormClosed);

            更改密碼 frm5 = new 更改密碼();
            frm5.FormClosed += new FormClosedEventHandler(frm5_FormClosed);

            frm5.Show();
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

        private void Form3_Closed(object sender, FormClosedEventArgs e)
        {

        }
    }
}
