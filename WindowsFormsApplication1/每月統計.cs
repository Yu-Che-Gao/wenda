using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class 每月統計 : Form
    {
        public SQLiteConnection conn = new SQLiteConnection(@"Data Source=database1.dat");

        public 每月統計()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {           
            comboBox1.Items.Add("全部");
            comboBox2.Items.Add("全部");
            comboBox3.Items.Add("耗材");
            comboBox3.Items.Add("處置");
            comboBox1.SelectedItem = "全部";
            comboBox2.SelectedItem = "全部";
            comboBox3.SelectedItem = "耗材";
            DataTable dt = getTable(comboBox1.Text,comboBox2.Text,comboBox3.Text);
            dataGridView1.DataSource = dt;
            dataGridView1.Height = Screen.PrimaryScreen.WorkingArea.Height - 100;
            dataGridView1.Width = Screen.PrimaryScreen.WorkingArea.Width;
            getSchool();
            getTeam("全部");
            button3.Enabled = false;
        }


        private DataTable get()
        {
            DataTable returnTable = new DataTable();
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();

            cmd.CommandText = "SELECT  strftime('%m',日期) as 月份,sum(白貼0_5) as 白貼0_5,sum(白貼1_5) as 白貼1_5,sum(輕彈1) as 輕彈1,sum(輕彈2) as 輕彈2,sum(輕彈3) as 輕彈3,sum(強彈1) as 強彈1,sum(強彈2) as 強彈2,sum(強彈3) as 強彈3,sum(墊片1_4) as 墊片1_4,sum(機能貼布) as 機能貼布,sum(KG3) as KG3,sum(膠膜) as 膠膜,sum(內膜) as 內膜,sum(備註) as 備註 FROM member group by  strftime('%m',日期) ";

            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            cmd.ExecuteNonQuery();
            conn.Close();
            return returnTable;
        }

        private void getSchool()
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();

            //cmd.CommandText = "SELECT  strftime('%m',日期) as 月份,sum(白貼0_5) as 白貼0_5,sum(白貼1_5) as 白貼1_5,sum(輕彈1) as 輕彈1,sum(輕彈2) as 輕彈2,sum(輕彈3) as 輕彈3,sum(強彈1) as 強彈1,sum(強彈2) as 強彈2,sum(強彈3) as 強彈3,sum(墊片1_4) as 墊片1_4,sum(機能貼布) as 機能貼布,sum(KG3) as KG3,sum(膠膜) as 膠膜,sum(內膜) as 內膜,sum(備註) as 備註 FROM member group by  strftime('%m',日期) ";
            cmd.CommandText = "SELECT 校名 FROM schoolName";
            using (SQLiteDataReader dr = cmd.ExecuteReader())
            {
                using (DataTable dt = new DataTable())
                {
                    dt.Load(dr);
                    for (int i1 = 0; i1 < dt.Rows.Count; i1++)
                    {
                        DataRow row = dt.Rows[i1];
                        comboBox1.Items.Add(row["校名"]);
                    }
                }
            }
            conn.Close();
        }

        private void getTeam(string schoolName)
        {
            conn.Open();
            comboBox2.Items.Clear();
            comboBox2.Text = "全部";
            comboBox2.Items.Add("全部");
            SQLiteCommand cmd = conn.CreateCommand();
            if (schoolName!="全部")
            {
                cmd.CommandText = "SELECT 隊伍名 FROM schoolTeam WHERE 校名='" + schoolName + "'";
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    using (DataTable dt = new DataTable())
                    {
                        dt.Load(dr);
                        for (int i1 = 0; i1 < dt.Rows.Count; i1++)
                        {
                            DataRow row = dt.Rows[i1];
                            comboBox2.Items.Add(row["隊伍名"]);
                        }
                    }
                }
            }
            conn.Close();
        }

        private DataTable getTable(string schoolName,string teamName,string selectThing)
        {
            DataTable returnTable = new DataTable();
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "";
            if (selectThing == "耗材")
            {
                if (schoolName == "全部")
                {
                    cmd.CommandText = "SELECT strftime('%m',member.日期) as 月份,sum(member.白貼0_5) as 白貼0_5,sum(member.白貼1_5) as 白貼1_5,sum(member.輕彈1) as 輕彈1,sum(member.輕彈2) as 輕彈2,sum(member.輕彈3) as 輕彈3,sum(member.強彈1) as 強彈1,sum(member.強彈2) as 強彈2,sum(member.強彈3) as 強彈3,sum(member.墊片1_4) as 墊片1_4,sum(member.機能貼布) as 機能貼布,sum(member.KG3) as KG3,sum(member.膠膜) as 膠膜,sum(member.內膜) as 內膜 FROM schoolTeam,teamMember,member WHERE schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員 GROUP BY strftime('%m',member.日期)";
                }
                else
                {
                    if (teamName == "全部")
                    {
                        cmd.CommandText = "SELECT schoolTeam.校名,strftime('%m',member.日期) as 月份,sum(member.白貼0_5) as 白貼0_5,sum(member.白貼1_5) as 白貼1_5,sum(member.輕彈1) as 輕彈1,sum(member.輕彈2) as 輕彈2,sum(member.輕彈3) as 輕彈3,sum(member.強彈1) as 強彈1,sum(member.強彈2) as 強彈2,sum(member.強彈3) as 強彈3,sum(member.墊片1_4) as 墊片1_4,sum(member.機能貼布) as 機能貼布,sum(member.KG3) as KG3,sum(member.膠膜) as 膠膜,sum(member.內膜) as 內膜 FROM schoolTeam,teamMember,member WHERE schoolTeam.校名='" + schoolName + "' AND schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員 GROUP BY strftime('%m',member.日期)";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT schoolTeam.校名,teamMember.隊伍名,strftime('%m',member.日期) as 月份,sum(member.白貼0_5) as 白貼0_5,sum(member.白貼1_5) as 白貼1_5,sum(member.輕彈1) as 輕彈1,sum(member.輕彈2) as 輕彈2,sum(member.輕彈3) as 輕彈3,sum(member.強彈1) as 強彈1,sum(member.強彈2) as 強彈2,sum(member.強彈3) as 強彈3,sum(member.墊片1_4) as 墊片1_4,sum(member.機能貼布) as 機能貼布,sum(member.KG3) as KG3,sum(member.膠膜) as 膠膜,sum(member.內膜) as 內膜 FROM schoolTeam,teamMember,member WHERE schoolTeam.校名='" + schoolName + "' AND teamMember.隊伍名='" + teamName + "' AND schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員 GROUP BY strftime('%m',member.日期)";
                    }
                }
            }else if (selectThing == "處置")
            {
                if (schoolName == "全部")
                {
                    cmd.CommandText = "SELECT strftime('%m',member.日期) as 月份,sum(member.處置 like '%冰敷%') as 冰敷次數,sum(member.處置 like '%熱敷%') as 熱敷次數,sum(member.處置 like '%貼紮%') as 貼紮次數,sum(member.處置 like '%外傷%') as 外傷次數 FROM schoolTeam,teamMember,member WHERE schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員 GROUP BY strftime('%m',member.日期)";
                }
                else
                {
                    if (teamName == "全部")
                    {
                        cmd.CommandText = "SELECT schoolTeam.校名,strftime('%m',member.日期) as 月份,sum(member.處置 like '%冰敷%') as 冰敷次數,sum(member.處置 like '%熱敷%') as 熱敷次數,sum(member.處置 like '%貼紮%') as 貼紮次數,sum(member.處置 like '%外傷%') as 外傷次數 FROM schoolTeam,teamMember,member WHERE schoolTeam.校名='" + schoolName + "' AND schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員 GROUP BY strftime('%m',member.日期)";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT schoolTeam.校名,teamMember.隊伍名,strftime('%m',member.日期) as 月份,sum(member.處置 like '%冰敷%') as 冰敷次數,sum(member.處置 like '%熱敷%') as 熱敷次數,sum(member.處置 like '%貼紮%') as 貼紮次數,sum(member.處置 like '%外傷%') as 外傷次數 FROM schoolTeam,teamMember,member WHERE schoolTeam.校名='" + schoolName + "' AND teamMember.隊伍名='" + teamName + "' AND schoolTeam.隊伍名=teamMember.隊伍名 AND teamMember.隊員=member.隊員 GROUP BY strftime('%m',member.日期)";
                    }
                }
            }
                

            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            
            cmd.ExecuteNonQuery();
            conn.Close();
            return returnTable;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            getTeam(comboBox1.Text);
            dataGridView1.DataSource = null;
            DataTable data = getTable(comboBox1.Text,comboBox2.Text,comboBox3.Text);
            dataGridView1.DataSource = data;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable data = getTable(comboBox1.Text, comboBox2.Text, comboBox3.Text);
            dataGridView1.DataSource = data;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable data = getTable(comboBox1.Text, comboBox2.Text, comboBox3.Text);
            dataGridView1.DataSource = data;
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

        private void button5_Click(object sender, EventArgs e)
        {
            if (codeboolisExcelInstalled())
            {
                button5.Enabled = false;
                button5.Text = "匯出處理中...";
                exportToExcel();
                button5.Text = "Excel匯出";
                button5.Enabled = true;
            }
            else
            {
                MessageBox.Show("已偵測到您的系統未正確安裝Office 2007或更高的版本", "警告");
            }
        }

        private bool codeboolisExcelInstalled()
        {
            Type type = Type.GetTypeFromProgID("Excel.Application");
            return type != null;
        }


        private DataSet GetDataSetFromDataGridView(DataGridView ucgrd)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            for (int j = 0; j < ucgrd.Columns.Count; j++)
            {
                dt.Columns.Add(ucgrd.Columns[j].HeaderCell.Value.ToString());
            }

            for (int j = 0; j < ucgrd.Rows.Count; j++)
            {
                DataRow dr = dt.NewRow();
                for (int i = 0; i < ucgrd.Columns.Count; i++)
                {
                    if (ucgrd.Rows[j].Cells[i].Value != null)
                    {
                        dr[i] = ucgrd.Rows[j].Cells[i].Value.ToString();
                    }
                    else
                    {
                        dr[i] = "";
                    }
                }
                dt.Rows.Add(dr);
            }
            ds.Tables.Add(dt);

            return ds;
        }

        private void exportToExcel()
        {
            //Print using Ofice InterOp
            Excel.Application excel = new Excel.Application();

            var workbook = (Excel._Workbook)(excel.Workbooks.Add(Missing.Value));
            var dataset = GetDataSetFromDataGridView(dataGridView1);

            for (var i = 0; i < dataset.Tables.Count; i++)
            {

                if (workbook.Sheets.Count <= i)
                {
                    workbook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing);
                }

                //NOTE: Excel numbering goes from 1 to n
                var currentSheet = (Excel._Worksheet)workbook.Sheets[i + 1];

                for (var j = 0; j <= this.dataGridView1.ColumnCount - 1; j++)
                {
                    currentSheet.Cells[1, j+1] = dataGridView1.Columns[j].HeaderText;
                }

                for (var y = 1; y < dataset.Tables[i].Rows.Count; y++)
                {
                    for (var x = 0; x < dataset.Tables[i].Rows[y].ItemArray.Count(); x++)
                    {
                        currentSheet.Cells[y + 1, x + 1] = dataset.Tables[i].Rows[y-1].ItemArray[x];
                    }
                }
            }

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.OverwritePrompt = false;
            saveFileDialog1.Filter = "Excel活頁簿(*.xlsx)|*.xlsx|Excel 97-2003活頁簿(*.xls)|*.xls";
            if ((saveFileDialog1.ShowDialog() == DialogResult.OK))
            {
                workbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing);
                workbook.Close();
                MessageBox.Show("匯出完成");
            }
            else
            {
                MessageBox.Show("匯出已取消");
            }
        }

        private void Form4_Closed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
