using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Data.SQLite;


namespace WindowsFormsApplication1
{

    public partial class 新增紀錄表 : Form
    {
        public SQLiteConnection conn = new SQLiteConnection(@"Data Source=database1.dat");
        String schoolnameadd;
        String teamnameadd;
        String membernameadd;
        String injurySide;
        String injuryCategory;
        String injuryKind;
        String injuryHandle;
        private bool flag = false;
        public delegate string insertHandler();

        public 新增紀錄表()
        {
            InitializeComponent();
            this.InputLanguageChanged += new InputLanguageChangedEventHandler(languageChange);
        }

        private void languageChange(Object sender,InputLanguageChangedEventArgs e)
        {
            textBox1.ImeMode = System.Windows.Forms.ImeMode.OnHalf;  // 將控制項的ImeMode設為OnHalf 
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            schoolnameadd = comboBox1.Text;
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox3.Items.Clear();
            get2();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            teamnameadd = comboBox2.Text;
            comboBox3.Text = "";
            get3();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string value = "";
            if (InputBox("請輸入學校名", "新增學校:", ref value)== DialogResult.OK)
            {
                    if(value=="")
                    {
                        MessageBox.Show("請勿輸入空值!!");
                    }
                    else
                    {
                        schoolnameadd = value;
                        insert(schoolnameadd);
                        get1();
                        comboBox1.SelectedItem = schoolnameadd;
                    }
            }
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string value = "";
            if (InputBox("請輸入隊名", "新增隊伍:", ref value) == DialogResult.OK)
            {
                if (value == "")
                {
                    MessageBox.Show("請勿輸入空值!!");
                }
                else
                {
                    teamnameadd = value;
                    insert2(teamnameadd);
                    get2();
                    comboBox2.SelectedItem = teamnameadd;
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string value = "";
            if (InputBox("請輸入隊員", "新增隊員:", ref value) == DialogResult.OK)
            {
                if (value == "")
                {
                    MessageBox.Show("請勿輸入空值!!");
                }
                else
                {
                    membernameadd = value;
                    insert3(membernameadd);
                    get3();
                    comboBox3.SelectedItem = membernameadd;
                }
            }
        }

        private void get1()
        {
            DataTable returnTable = new DataTable();
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 校名 FROM schoolName";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            comboBox1.Items.Clear();
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

        private void get2()
        {
            DataTable returnTable = new DataTable();
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 隊伍名 FROM schoolTeam WHERE 校名='" + schoolnameadd + "'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            comboBox2.Items.Clear();
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
            conn.Close();
        }

        private void get3()
        {
            DataTable returnTable = new DataTable();
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 隊員 FROM teamMember WHERE 隊伍名='" + teamnameadd + "'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            comboBox3.Items.Clear();
            using (SQLiteDataReader dr = cmd.ExecuteReader())
            {
                using (DataTable dt = new DataTable())
                {
                    dt.Load(dr);
                    for (int i1 = 0; i1 < dt.Rows.Count; i1++)
                    {
                        DataRow row = dt.Rows[i1];
                        comboBox3.Items.Add(row["隊員"]);
                    }
                }
            }
            conn.Close();
        }

        private Boolean checkRepeat1(string school)
        {
            DataTable returnTable = new DataTable();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 校名 FROM schoolName WHERE 校名='"+school+"'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            comboBox1.Items.Clear();
            using (SQLiteDataReader dr = cmd.ExecuteReader())
            {
                using (DataTable dt = new DataTable())
                {
                    dt.Load(dr);
                    if (dt.Rows.Count == 0)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
        }

        private Boolean checkRepeat2(string school,string team)
        {
            DataTable returnTable = new DataTable();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 隊伍名 FROM schoolTeam WHERE 校名='" + school + "' and 隊伍名='" + team + "'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            comboBox3.Items.Clear();
            using (SQLiteDataReader dr = cmd.ExecuteReader())
            {
                using (DataTable dt = new DataTable())
                {
                    dt.Load(dr);
                    if (dt.Rows.Count == 0)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
        }

        private Boolean checkRepeat3(string team,string member)
        {
            DataTable returnTable = new DataTable();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 隊員 FROM teamMember WHERE 隊伍名='"+team+"' and 隊員='"+member+"'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(returnTable);
            comboBox3.Items.Clear();
            using (SQLiteDataReader dr = cmd.ExecuteReader())
            {
                using (DataTable dt = new DataTable())
                {
                    dt.Load(dr);
                    if(dt.Rows.Count==0)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
        }

        private void insert(string schoolname)
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@schoolName", schoolname);
            if(!checkRepeat1(schoolname))
            {
                cmd.CommandText = "INSERT INTO schoolName(校名) VALUES (@schoolName)";
                cmd.ExecuteNonQuery();
            }
            else
            {
                MessageBox.Show("此筆資料已輸入過!!");
            }
            
            conn.Close();
        }

        private void insert2(string teamName)
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@schoolName", schoolnameadd);
            cmd.Parameters.AddWithValue("@teamName", teamName);
            if(!checkRepeat2(schoolnameadd,teamName))
            {
                cmd.CommandText = "INSERT INTO schoolTeam(校名,隊伍名) VALUES (@schoolName,@teamName)";
                cmd.ExecuteNonQuery();
            }
            else
            {
                MessageBox.Show("此筆資料已輸入過!!");
            }
            conn.Close();
        }

        private void insert3(string memberName)
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@teamName", teamnameadd);
            cmd.Parameters.AddWithValue("@memberName", memberName);
            if(!checkRepeat3(teamnameadd,memberName))
            {
                cmd.CommandText = "INSERT INTO teamMember(隊伍名,隊員) VALUES (@teamName,@memberName)";
                cmd.ExecuteNonQuery();
            }
            else
            {
                MessageBox.Show("此筆資料已輸入過!!");
            }
            conn.Close();
        }

        private void insert4()
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@date1", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@memberName", comboBox3.Text);
            cmd.Parameters.AddWithValue("@injuryPlace", textBox1.Text);
            cmd.Parameters.AddWithValue("@injurySide", injurySide);
            cmd.Parameters.AddWithValue("@injuryCategory", injuryCategory);
            cmd.Parameters.AddWithValue("@injuryKind", injuryKind);
            cmd.Parameters.AddWithValue("@injuryHandle", injuryHandle);
            cmd.Parameters.AddWithValue("@White05", numericUpDown1.Value);
            cmd.Parameters.AddWithValue("@White15", numericUpDown2.Value);
            cmd.Parameters.AddWithValue("@Light1", numericUpDown3.Value);
            cmd.Parameters.AddWithValue("@Light2", numericUpDown4.Value);
            cmd.Parameters.AddWithValue("@Light3", numericUpDown5.Value);
            cmd.Parameters.AddWithValue("@Strong1", numericUpDown6.Value);
            cmd.Parameters.AddWithValue("@Strong2", numericUpDown7.Value);
            cmd.Parameters.AddWithValue("@Strong3", numericUpDown8.Value);
            cmd.Parameters.AddWithValue("@Piece14", numericUpDown9.Value);
            cmd.Parameters.AddWithValue("@muscle", numericUpDown10.Value);
            cmd.Parameters.AddWithValue("@KG3", numericUpDown11.Value);
            cmd.Parameters.AddWithValue("@Glue", numericUpDown12.Value);
            cmd.Parameters.AddWithValue("@Inside", numericUpDown13.Value);
            cmd.Parameters.AddWithValue("@ps", textBox2.Text);

            cmd.CommandText = "INSERT INTO member(日期,隊員,受傷部位,傷側,受傷種類,受傷分類,處置,白貼0_5,白貼1_5,輕彈1,輕彈2,輕彈3,強彈1,強彈2,強彈3,墊片1_4,機能貼布,KG3,膠膜,內膜,備註) VALUES (@date1,@memberName,@injuryPlace,@injurySide,@injuryCategory,@injuryKind,@injuryHandle,@White05,@White15,@Light1,@Light2,@Light3,@Strong1,@Strong2,@Strong3,@Piece14,@muscle,@KG3,@Glue,@Inside,@ps)";
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void createNewTable()
        {
            conn.Open();
            string commandText = "CREATE TABLE schoolName(ID INTEGER PRIMARY KEY AUTOINCREMENT,校名 VARCHAR(25))";
            SQLiteCommand cmd = new SQLiteCommand(commandText, conn);
            cmd.ExecuteNonQuery();
            string commandText2 = "CREATE TABLE schoolTeam(ID INTEGER PRIMARY KEY AUTOINCREMENT,校名 VARCHAR(25),隊伍名 VARCHAR(25))";
            SQLiteCommand cmd2 = new SQLiteCommand(commandText2, conn);
            cmd2.ExecuteNonQuery();
            string commandText3 = "CREATE TABLE teamMember(ID INTEGER PRIMARY KEY AUTOINCREMENT,隊伍名 VARCHAR(25),隊員 VARCHAR(25))";
            SQLiteCommand cmd3 = new SQLiteCommand(commandText3, conn);
            cmd3.ExecuteNonQuery();
            string commandText4 = "CREATE TABLE member(ID INTEGER PRIMARY KEY AUTOINCREMENT,日期 DATETIME,隊員 VARCHAR(25),受傷部位 VARCHAR(30),傷側 VARCHAR(4),受傷種類 VARCHAR(12),受傷分類 VARCHAR(4),處置 VARCHAR(16),白貼0_5 TEXT,白貼1_5 TEXT,輕彈1 TEXT,輕彈2 TEXT,輕彈3 TEXT,強彈1 TEXT,強彈2 TEXT,強彈3 TEXT,墊片1_4 TEXT,機能貼布 TEXT,KG3 TEXT,膠膜 TEXT,內膜 TEXT,備註 VARCHAR(100))";
            SQLiteCommand cmd4 = new SQLiteCommand(commandText4, conn);
            cmd4.ExecuteNonQuery();
            conn.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
           
            if (!File.Exists("database1.dat"))
            {
                SQLiteConnection.CreateFile("database1.dat");
                createNewTable();
            }
            else
            {
                get1();
            }

            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            button1.Enabled = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                injurySide += "左";
            }
            if (checkBox2.Checked == true)
            {
                injurySide += "右";
            }
            if (checkBox3.Checked == true)
            {
                injuryCategory += "骨折";
            }
            if (checkBox4.Checked == true)
            {
                injuryCategory += "脫臼";
            }
            if (checkBox5.Checked == true)
            {
                injuryCategory += "拉傷";
            }
            if (radioButton1.Checked == true)
            {
                injuryKind = "新傷";
            }
            if (radioButton2.Checked == true)
            {
                injuryKind = "舊傷";
            }
            if (radioButton3.Checked == true)
            {
                injuryKind = "預防";
            }
            if (checkBox6.Checked == true)
            {
                injuryHandle += "外傷";
            }
            if (checkBox7.Checked == true)
            {
                injuryHandle += "熱敷";
            }
            if (checkBox8.Checked == true)
            {
                injuryHandle += "冰敷";
            }
            if (checkBox9.Checked == true)
            {
                injuryHandle += "貼紮";
            }
            if (comboBox1.Text != "" && comboBox2.Text != "" && comboBox3.Text != "" && textBox1.Text != "")
            {
                insert4();
                clearInput();
                this.Hide();
                查詢 frm3 = new 查詢();
                frm3.Show();
            }
            else
            {
                MessageBox.Show("尚未填入完整表格");
                if(comboBox1.Text=="")
                {
                    label2.ForeColor = Color.Red;
                    label2.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Bold);
                }
                else
                {
                    label2.ForeColor = Color.Black;
                    label2.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Regular);
                }
                
                if(comboBox2.Text=="")
                {
                    label3.ForeColor = Color.Red;
                    label3.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Bold);
                }
                else
                {
                    label3.ForeColor = Color.Black;
                    label3.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Regular);
                }

                if(comboBox3.Text=="")
                {
                    label7.ForeColor = Color.Red;
                    label7.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Bold);
                }
                else
                {
                    label7.ForeColor = Color.Black;
                    label7.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Regular);
                }

                if(textBox1.Text=="")
                {
                    label4.ForeColor = Color.Red;
                    label4.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Bold);
                }
                else
                {
                    label4.ForeColor = Color.Black;
                    label4.Font = new Font(new FontFamily(label1.Font.Name), 12, FontStyle.Regular);
                }
            }
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

        void frm3_FormClosed(object sender,FormClosedEventArgs e)
        {
            this.Show();
        }

        void frm4_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Show();
        }

        void frm5_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Show();
        }

        private void clearInput()
        {
            DateTime today = DateTime.Today;
            dateTimePicker1.Value = today;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            textBox1.Text = "";
            textBox2.Text = "";
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            numericUpDown1.Value = 0;
            numericUpDown2.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown4.Value = 0;
            numericUpDown5.Value = 0;
            numericUpDown6.Value = 0;
            numericUpDown7.Value = 0;
            numericUpDown8.Value = 0;
            numericUpDown9.Value = 0;
            numericUpDown10.Value = 0;
            numericUpDown11.Value = 0;
            numericUpDown12.Value = 0;
            numericUpDown13.Value = 0;
            get1();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            delete1();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            get1();
        }

        private void delete1()
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@schoolName", comboBox1.Text);
            cmd.CommandText = "DELETE FROM schoolName WHERE 校名=@schoolName";
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            delete2();
            comboBox3.Items.Clear();
            get2();
        }

        private void delete2()
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@teamName", comboBox2.Text);
            cmd.CommandText = "DELETE FROM schoolTeam WHERE 隊伍名=@teamName";
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void delete3()
        {
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@member", comboBox3.Text);
            cmd.CommandText = "DELETE FROM teamMember WHERE 隊員=@member";
            cmd.ExecuteNonQuery();
            conn.Close();
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

        private void Form2_Closed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void 新增紀錄表_Shown(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            textBox.ImeMode = System.Windows.Forms.ImeMode.OnHalf;  // 將控制項的ImeMode設為OnHalf 
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;
            
            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            delete3();
            get3();
        }

        private void numericUpDown1_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        private void numericUpDown2_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown2.Value);
            //if (isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown3_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown3.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown4_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown4.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown5_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown5.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown11_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown11.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown13_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown13.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown6_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown6.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown7_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown7.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown8_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown8.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown9_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown9.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown10_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown10.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown12_KeyDown(object sender, KeyEventArgs e)
        {
            //String temp = Convert.ToString(numericUpDown12.Value);
            //if (!isFullWord(temp))
            //{
            //    MessageBox.Show("請勿輸入全形字!!");
            //}
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        static public string TransStringType(string oriString, int transType)
        {
            string value = string.Empty;

            try
            {
                if (transType != 0 && transType != 1)
                {
                    value = oriString;
                }

                if (transType == 0)
                {
                    value = Microsoft.VisualBasic.Strings.StrConv(oriString, Microsoft.VisualBasic.VbStrConv.Wide, 0);
                }
                else if (transType == 1)
                {
                    value = Microsoft.VisualBasic.Strings.StrConv(oriString, Microsoft.VisualBasic.VbStrConv.Narrow, 0);
                }
            }
            catch (Exception ex)
            {
                value = oriString;
            }

            return value;
        }

        public static bool isFullWord(string words)
        {
            bool result = false;
            string pattern = @"^[\u4E00-\u9fa5]+$";
            foreach (char item in words)
            {
                //以Regex判斷是否為中文字，中文字視為全形
                if (!System.Text.RegularExpressions.Regex.IsMatch(item.ToString(), pattern))
                {
                    //以16進位值長度判斷是否為全形字
                    if (string.Format("{0:X}", Convert.ToInt32(item)).Length != 2)
                    {
                        result = true;
                        break;
                    }
                }
            }
            return result;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown11_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown13_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown9_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown10_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {
            string tempString = string.Empty;
            foreach (char str in (((NumericUpDown)sender).Value).ToString())
            {
                tempString += TransStringType(str.ToString(), 1);
            }

            ((NumericUpDown)sender).Value = Convert.ToDecimal(tempString);
        }

        private void numericUpDown1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag==false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
            
        }

        private void numericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
            
        }

        private void numericUpDown2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }

        private void numericUpDown12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (flag == false)
            {
                flag = true;
            }
            else
            {
                if (isFullWord(e.KeyChar.ToString()) && e.KeyChar != (char)Keys.Back)
                {
                    MessageBox.Show("請勿輸入全形字");
                }
                flag = false;
            }
        }



        //public void isFullWord(string word)
        //{
        //    string pattern = @"^[\u4E00-\u9fa5]+$";
        //    foreach (char item in word)
        //    {
        //        //以Regex判斷是否為中文字，中文字視為全形 
        //        if (!System.Text.RegularExpressions.Regex.IsMatch(item.ToString(), pattern))
        //        {
        //            //以16進位值長度判斷是否為全形字
        //            if (string.Format("{0:X}", Convert.ToInt32(item)).Length != 2)
        //            {
        //                MessageBox.Show("請勿輸入全形字!!");
        //                break;
        //            }
        //        }
        //    }
        //  }

    }
}
