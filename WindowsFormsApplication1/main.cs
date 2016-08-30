using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.SQLite;
using System.IO;
using System.Collections;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;
using Microsoft.VisualBasic;

namespace WindowsFormsApplication1
{
    public partial class main : Form
    {
        //當日用量
        Label[] labelToday1;
        Label[] labelToday2;

        //新增
        Label[] labelArray;
        CheckBox[] checkBoxArray;
        Button[] deleteButtonArray;

        //進貨紀錄
        Label[] labelArray2;
        TextBox[] textBoxArray2;
        Button[] negativeButtonArray2;
        Button[] positiveButtonArray2;
        Button[] deleteButtonArray2;

        //受傷部位
        ComboBox[] comboBoxPosArray;
        CheckBox[] checkBoxLeftArray;
        CheckBox[] checkBoxRightArray;
        Button[] deletePosArray;
        Button[] negativeButtonArray;
        Button[] positiveButtonArray;
        Button[] deleteButtonArrayHandle;
        Label[] leftLabel;
        Label[] rightLabel;
        Label[] labelPosArray;
        Label[] No;
        TextBox[] textBoxPosArray;
        int comboBoxPosFlag = 0;


        string temp;

        int objectNameCount;
        int handleCount;
        int kindCount;
        int categoryCount;
        int sideCount;
        int posCount;

        bool threadFlag_InitFHSDb = false;
        // P變數
        SQLiteConnection conn = new SQLiteConnection("Data Source=" + Constants.FHSDatabaseFile);
        SQLiteConnection connToday = new SQLiteConnection("Data Source=" + Constants.FHSTodayDatabaseFile);
        SQLiteConnection connLogin = new SQLiteConnection("Data Source=" + Constants.LoginDatabaseFile);
        public main()
        {
            InitializeComponent();
        }

        public void alertThread()
        {
            threadFlag_InitFHSDb = true;
            SQLiteConnection.CreateFile(Constants.FHSDatabaseFile);

            DBManage.createOrInsertCmd(conn, "CREATE TABLE team(teamID INTEGER PRIMARY KEY AUTOINCREMENT,teamName VARCHAR(25))");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE member(memberID INTEGER PRIMARY KEY AUTOINCREMENT,name VARCHAR(25),teamID INTEGER)");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE injurySide(sideID INTEGER PRIMARY KEY AUTOINCREMENT,side VARCHAR(5))");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injurySide(side) VALUES ('左')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injurySide(side) VALUES ('右')");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE injuryPos(posID INTEGER PRIMARY KEY AUTOINCREMENT,pos VARCHAR(25))");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('頭部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('顏面')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('眼')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('鼻')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('耳')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('嘴')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('頸部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('胸部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('腹部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('上背')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('下背')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('肩關節')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('上臂')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('手肘')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('前臂')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('腕關節')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('掌部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('指關節')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('髖關節')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('臀部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('鼠蹊部')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('大腿')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('膝關節')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('脛骨')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('小腿')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('踝關節')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('跟腱')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('足背')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('足底')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('趾關節')");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE injuryKind(kindID INTEGER PRIMARY KEY AUTOINCREMENT,kind VARCHAR(25))");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryKind(kind) VALUES ('扭傷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryKind(kind) VALUES ('拉傷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryKind(kind) VALUES ('挫傷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryKind(kind) VALUES ('脫臼')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryKind(kind) VALUES ('骨折')");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE injuryCategory(categoryID INTEGER PRIMARY KEY AUTOINCREMENT,category VARVHAR(25))");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryCategory(category) VALUES ('預防')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryCategory(category) VALUES ('新傷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryCategory(category) VALUES ('舊傷')");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE injuryHandle(handleID INTEGER PRIMARY KEY AUTOINCREMENT,handle VARCHAR(25))");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('貼紮')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('冰敷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('熱敷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('外傷')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('冷熱交替')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('肌內效貼布')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('生動乳膠')");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE object(objectID INTEGER PRIMARY KEY AUTOINCREMENT,objectName TEXT)");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('內膜')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('白貼0.5吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('白貼1吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('白貼1.5吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('輕彈1吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('輕彈2吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('強彈1吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('強彈2吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('強彈3吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('墊片')");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE insertLog(insertID INTEGER PRIMARY KEY AUTOINCREMENT,insertTime TEXT, contestName TEXT,teamID TEXT,memberID TEXT,psText TEXT)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE logPos(PID INTEGER PRIMARY KEY AUTOINCREMENT, posID INTEGER, insertID INTEGER, side TEXT)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE logSide(SID INTEGER PRIMARY KEY AUTOINCREMENT, sideID INTEGER, insertID INTEGER)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE logKind(KID INTEGER PRIMARY KEY AUTOINCREMENT, kindID INTEGER, insertID INTEGER)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE logCategory(CID INTEGER PRIMARY KEY AUTOINCREMENT, categoryID INTEGER, insertID INTEGER)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE logHandle(HID INTEGER PRIMARY KEY AUTOINCREMENT, handleID INTEGER, insertID INTEGER, handleCount TEXT)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE logObject(OID INTEGER PRIMARY KEY AUTOINCREMENT, objectID INTEGER, insertID INTEGER, number INTEGER)");

            DBManage.createOrInsertCmd(conn, "CREATE TABLE insertGoodLog(IGID INTEGER PRIMARY KEY AUTOINCREMENT,insertMonth TEXT)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE goodLog(GID INTEGER PRIMARY KEY AUTOINCREMENT,goodID TEXT,number INTEGER,IGID INTEGER)");
            DBManage.createOrInsertCmd(conn, "CREATE TABLE good(goodID INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT)");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('1.5吋白貼')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('1吋白貼')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('0.5吋白貼')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('內膜')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('2吋輕彈')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('2吋強彈')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('3吋強彈')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('蕾絲墊片')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('肌效貼布')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('墊片1/2')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('墊片1/4')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('墊片1/8')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('黏著劑')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('去黏著劑')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('冷凍劑')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('生理食鹽水')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('凡士林')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('2nd Skin')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('優碘')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('kg3')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('透氣膠帶0.5吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('透氣膠帶1吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('免縫膠帶')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('消炎藥膏')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('大ok蹦')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('小ok蹦')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('紗布3*3')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('棉花棒')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('止血棒')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('彈蹦4吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('彈蹦4吋加長')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('彈蹦6吋')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('彈蹦6吋加長')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('三角巾')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('大裁刀')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('小裁刀')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('Y字剪')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('一字剪')");
            DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('鐵片固定')");
            threadFlag_InitFHSDb = false;
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("您確定要關閉程式?", "警告", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                e.Cancel = true;
            }
        }

        private void main_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void main_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void main_Resize(object sender, EventArgs e)
        {
            this.Width = 1024;
            this.Height = 647;
        }

        private void main_Load(object sender, EventArgs e)
        {
            // check fhs.dat
            if (!File.Exists(Constants.FHSDatabaseFile))
            {
                alertForm alertForm = new alertForm();
                alertForm.setTitle("環境設定");
                alertForm.setContent("第一次啟動，環境設定中......\r\n請勿關閉本視窗，並耐心等候。");
                alertForm.Show();

                Thread thread = new Thread(alertThread);
                thread.Start();
                while (threadFlag_InitFHSDb) ;
                //thread.Join();

                alertForm.Hide();
            }

            //UI初始化
            psTextBox.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            contestTextBox.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox4.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox5.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox6.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox7.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox8.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox9.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
            textBox10.ImeMode = System.Windows.Forms.ImeMode.OnHalf;

            comboBoxTab2.SelectedItem = "全部";

            comboBoxTeamTab3.Items.Add("全部");
            comboBoxTeamTab3.SelectedItem = "全部";

            comboBoxKindTab3.Items.Add("耗材");
            comboBoxKindTab3.Items.Add("處置");
            comboBoxKindTab3.Items.Add("部位");
            comboBoxKindTab3.SelectedItem = "耗材";

            textBoxSchoolYear.Text = (Int32.Parse(DateTime.Now.Year.ToString()) - 1912).ToString();

            // 功能面初始化
            refreshComboBoxTeam(conn, "");

            DataTable memberName = DBManage.getTable(conn, "SELECT name FROM member WHERE teamID=(SELECT teamID FROM team WHERE teamName='" + teamComboBox.Text + "')");
            memberComboBox.Items.Clear();
            for (int i = 0; i < memberName.Rows.Count; i++)
            {
                memberComboBox.Items.Add(memberName.Rows[i]["name"]);
            }

            Queue tmp;

            DataTable pos = DBManage.getTable(conn, "SELECT pos FROM injuryPos ORDER BY posID");
            myPos(panelPos, pos);
            tmp = DBManage.selectCmd(conn, "SELECT handle FROM injuryHandle ORDER BY handleID", "handle");
            tmp = DBManage.selectCmd(conn, "SELECT objectName FROM object ORDER BY objectID", "objectName");
            myObject(objectPanel, tmp);

            DataTable goodName = DBManage.getTable(conn, "SELECT name FROM good ORDER BY goodID");
            panelInsertGood = myGood(panelInsertGood, goodName);
            getDataGridView1();

            if (!File.Exists(Constants.FHSTodayDatabaseFile))
            {
                SQLiteConnection.CreateFile(Constants.FHSTodayDatabaseFile);
                DBManage.createOrInsertCmd(connToday, "CREATE TABLE todayTable(Today TEXT,t1 TEXT,t2 TEXT,t3 TEXT,t4 TEXT,t5 TEXT,t6 TEXT,t7 TEXT)");
            }
        }


        private void insertTeamButton_Click(object sender, EventArgs e)
        {
            string val = "";
            if (Lib.InputBox("請輸入隊伍名", "新增隊伍:", ref val) == DialogResult.OK)
            {
                if (val == "")
                {
                    MessageBox.Show("請勿輸入空值");
                }
                else
                {
                    DBManage.createOrInsertCmd(conn, "INSERT INTO team(teamName) VALUES ('" + val + "')");
                    refreshComboBoxTeam(conn, val);
                }
            }
        }
        /*
            刷新TabC1的teamComboBox
            comboBoxTeamTab3
        */
        public void refreshComboBoxTeam(SQLiteConnection conn, string val)
        {
            DataTable teamName = DBManage.getTable(conn, "SELECT teamName FROM team");
            teamComboBox.Items.Clear();
            comboBoxTeamTab3.Items.Clear();
            comboBoxTeamTab3.Items.Add("全部");
            for (int i = 0; i < teamName.Rows.Count; i++)
            {
                teamComboBox.Items.Add(teamName.Rows[i]["teamName"]);
                comboBoxTeamTab3.Items.Add(teamName.Rows[i]["teamName"]);
            }
            comboBoxTeamTab3.Text = "全部";
            if (val != null && val != "")
                teamComboBox.Text = val;
        }

        private void insertMemberButton_Click(object sender, EventArgs e)
        {
            if (teamComboBox.Text != "")
            {
                string value = "";
                if (Lib.InputBox("請輸入人員名", "新增人員:", ref value) == DialogResult.OK)
                {
                    if (value == "")
                    {
                        MessageBox.Show("請勿輸入空值");
                    }
                    else
                    {
                        DataTable teamID = DBManage.getTable(conn, "SELECT teamID FROM team WHERE teamName='" + teamComboBox.Text + "'");
                        DBManage.createOrInsertCmd(conn, "INSERT INTO member(name,teamID) VALUES ('" + value + "','" + teamID.Rows[0]["teamID"] + "')");
                        DataTable memberName = DBManage.getTable(conn, "SELECT name FROM member WHERE teamID=(SELECT teamID FROM team WHERE teamName='" + teamComboBox.Text + "')");
                        memberComboBox.Items.Clear();
                        for (int i = 0; i < memberName.Rows.Count; i++)
                        {
                            memberComboBox.Items.Add(memberName.Rows[i]["name"]);
                        }
                        memberComboBox.Text = value;
                    }
                }
            }
        }

        private void deleteTeamButton_Click(object sender, EventArgs e)
        {
            if (teamComboBox.Text != "")
            {
                if (MessageBox.Show("您確定要刪除嗎?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string value = teamComboBox.Text;
                    DBManage.createOrInsertCmd(conn, "DELETE FROM member WHERE teamID=(SELECT teamID FROM team WHERE teamNAME='" + value + "')");
                    DBManage.createOrInsertCmd(conn, "DELETE FROM team WHERE teamName='" + value + "'");
                    DataTable teamName = DBManage.getTable(conn, "SELECT teamName FROM team");

                    teamComboBox.Items.Clear();
                    for (int i = 0; i < teamName.Rows.Count; i++)
                    {
                        teamComboBox.Items.Add(teamName.Rows[i]["teamName"]);
                    }
                    teamComboBox.Text = "";
                    memberComboBox.Items.Clear();
                }
            }
        }

        private void deleteMemberButton_Click(object sender, EventArgs e)
        {
            if (teamComboBox.Text != "" && memberComboBox.Text != "")
            {
                if (MessageBox.Show("您確定要刪除嗎?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string value = memberComboBox.Text;
                    DataTable teamID = DBManage.getTable(conn, "SELECT teamID FROM team WHERE teamName='" + teamComboBox.Text + "'");
                    DBManage.createOrInsertCmd(conn, "DELETE FROM member WHERE name='" + value + "' and teamID=" + teamID.Rows[0]["teamID"]);
                    DataTable memberName = DBManage.getTable(conn, "SELECT name FROM member WHERE teamID=(SELECT teamID FROM team WHERE teamName='" + teamComboBox.Text + "')");

                    memberComboBox.Items.Clear();
                    for (int i = 0; i < memberName.Rows.Count; i++)
                    {
                        memberComboBox.Items.Add(memberName.Rows[i]["name"]);
                    }
                    memberComboBox.Text = "";
                }
            }

        }

        private void myComboBox(ref string value, ComboBox myCombo, int type, string col, string table)
        {
            Queue tmp = null;
            switch (type)
            {
                case 1: //insert
                    DBManage.createOrInsertCmd(conn, "INSERT INTO " + table + "(" + col + ") VALUES ('" + value + "')");
                    tmp = DBManage.selectCmd(conn, "SELECT " + col + " FROM " + table, col);
                    setComboBox(myCombo, tmp);
                    myCombo.SelectedItem = value;
                    break;

                case 2: //delete
                    DBManage.createOrInsertCmd(conn, "DELETE FROM " + table + " WHERE " + col + " = '" + value + "'");
                    tmp = DBManage.selectCmd(conn, "SELECT " + col + " FROM " + table, col);
                    setComboBox(myCombo, tmp);
                    break;
            }
        }

        private void myObject(Panel myPanel, Queue queue)
        {
            labelArray = new Label[queue.Count - 1];
            checkBoxArray = new CheckBox[queue.Count - 1];
            string[] labelText = new string[queue.Count - 1];
            deleteButtonArray = new Button[queue.Count - 1];
            labelText = getLabelText(queue);
            objectPanel.Controls.Clear();

            for (int signal = 0; signal < labelArray.Length; signal++)
            {
                labelArray[signal] = new Label();
                labelArray[signal].Text = labelText[signal];
                labelArray[signal].Font = new Font("微軟正黑體", 12);
                labelArray[signal].Top = signal * 30;
                objectPanel.Controls.Add(labelArray[signal]);

                checkBoxArray[signal] = new CheckBox();
                checkBoxArray[signal].Top = signal * 30;
                checkBoxArray[signal].Left = labelArray[signal].Width + 2;
                checkBoxArray[signal].Width = 25;
                objectPanel.Controls.Add(checkBoxArray[signal]);

                deleteButtonArray[signal] = new Button();
                deleteButtonArray[signal].Font = new Font("微軟正黑體", 12);
                deleteButtonArray[signal].Top = signal * 30;
                deleteButtonArray[signal].Visible = true;
                deleteButtonArray[signal].Height = 28;
                deleteButtonArray[signal].AutoSize = true;
                deleteButtonArray[signal].Show();
                deleteButtonArray[signal].Left = labelArray[signal].Width + checkBoxArray[signal].Width + 5;
                deleteButtonArray[signal].Visible = false;
                deleteButtonArray[signal].Text = "刪除耗材";
                myPanel.Controls.Add(deleteButtonArray[signal]);
                deleteButtonArray[signal].Click += new System.EventHandler(delegate(object sender, EventArgs args) { Button temp = (Button)sender; if (MessageBox.Show("確定要刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) { deleteContestButtonAction(temp); } });
            }

        }

        private void myPos(Panel myPanel, DataTable queue)
        {
            negativeButtonArray = new Button[(queue.Rows.Count - 1) * (comboBoxPosFlag + 1)];
            positiveButtonArray = new Button[(queue.Rows.Count - 1) * (comboBoxPosFlag + 1)];
            deleteButtonArrayHandle = new Button[(queue.Rows.Count - 1) * (comboBoxPosFlag + 1)];

            No = new Label[comboBoxPosFlag + 1];
            comboBoxPosArray = new ComboBox[comboBoxPosFlag + 1];
            deletePosArray = new Button[comboBoxPosFlag + 1];
            leftLabel = new Label[comboBoxPosFlag + 1];
            rightLabel = new Label[comboBoxPosFlag + 1];
            checkBoxLeftArray = new CheckBox[comboBoxPosFlag + 1];
            checkBoxRightArray = new CheckBox[comboBoxPosFlag + 1];

            DataTable handle = DBManage.getTable(conn, "SELECT * FROM injuryHandle");
            labelPosArray = new Label[handle.Rows.Count];
            textBoxPosArray = new TextBox[handle.Rows.Count * (comboBoxPosFlag + 1)];
            panelPos.Controls.Clear();

            for (int i = 0; i < comboBoxPosFlag + 1; i++)
            {
                No[i] = new Label();
                No[i].Font = new Font("微軟正黑體", 12);
                No[i].Text = (i + 1).ToString() + ".";
                No[i].Top = (i * (labelPosArray.Length + 1)) * 30;
                No[i].Width = 23;
                panelPos.Controls.Add(No[i]);

                comboBoxPosArray[i] = new ComboBox();
                comboBoxPosArray[i].Font = new Font("微軟正黑體", 12);
                comboBoxPosArray[i].DropDownStyle = ComboBoxStyle.DropDownList;
                comboBoxPosArray[i].Width = 100;
                setComboBoxPos(comboBoxPosArray[i], queue);
                comboBoxPosArray[i].Top = (i * (labelPosArray.Length + 1)) * 30;
                comboBoxPosArray[i].Left = 30;
                panelPos.Controls.Add(comboBoxPosArray[i]);

                deletePosArray[i] = new Button();
                deletePosArray[i].Text = "刪除部位";
                deletePosArray[i].Font = new Font("微軟正黑體", 12);
                deletePosArray[i].Height = 28;
                deletePosArray[i].Width = comboBoxPosArray[i].Width;
                deletePosArray[i].Show();
                deletePosArray[i].Top = comboBoxPosArray[i].Bottom + 5;
                deletePosArray[i].Left = comboBoxPosArray[i].Left;
                deletePosArray[i].Visible = false;
                panelPos.Controls.Add(deletePosArray[i]);
                deletePosArray[i].Click += new System.EventHandler(delegate(object sender, EventArgs args)
                {
                    Button temp = (Button)sender;
                    if (comboBoxPosArray[Array.IndexOf(deletePosArray, temp)].Text != "")
                    {
                        if (MessageBox.Show("您確定要刪除嗎?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            string value = comboBoxPosArray[Array.IndexOf(deletePosArray, temp)].Text;
                            DBManage.createOrInsertCmd(conn, "DELETE FROM injuryPos WHERE pos='" + value + "'");
                            DataTable injuryPos = DBManage.getTable(conn, "SELECT pos FROM injuryPos");
                            myPos(panelPos, injuryPos);
                            button8.Text = "刪除部位";
                            button9.Text = "刪除處置";
                        }
                    }
                });

                leftLabel[i] = new Label();
                leftLabel[i].Text = "左";
                leftLabel[i].Font = new Font("微軟正黑體", 12);
                leftLabel[i].Width = 20;
                leftLabel[i].Left = comboBoxPosArray[i].Left;
                leftLabel[i].Top = deletePosArray[i].Bottom + 5;
                panelPos.Controls.Add(leftLabel[i]);

                checkBoxLeftArray[i] = new CheckBox();
                checkBoxLeftArray[i].Font = new Font("微軟正黑體", 12);
                checkBoxLeftArray[i].Width = 25;
                checkBoxLeftArray[i].Left = leftLabel[i].Left + 5;
                checkBoxLeftArray[i].Top = leftLabel[i].Bottom;
                panelPos.Controls.Add(checkBoxLeftArray[i]);

                rightLabel[i] = new Label();
                rightLabel[i].Text = "右";
                rightLabel[i].Width = 20;
                rightLabel[i].Font = new Font("微軟正黑體", 12);
                rightLabel[i].Left = leftLabel[i].Right + 5;
                rightLabel[i].Top = deletePosArray[i].Bottom + 5;
                panelPos.Controls.Add(rightLabel[i]);

                checkBoxRightArray[i] = new CheckBox();
                checkBoxRightArray[i].Font = new Font("微軟正黑體", 12);
                checkBoxRightArray[i].Width = 25;
                checkBoxRightArray[i].Left = checkBoxLeftArray[i].Right;
                checkBoxRightArray[i].Top = rightLabel[i].Bottom;
                panelPos.Controls.Add(checkBoxRightArray[i]);


                for (int signal = 0; signal < labelPosArray.Length; signal++)
                {
                    labelPosArray[signal] = new Label();
                    labelPosArray[signal].Text = handle.Rows[signal]["handle"].ToString();
                    labelPosArray[signal].Font = new Font("微軟正黑體", 12);
                    labelPosArray[signal].Left = deletePosArray[i].Right + 5;
                    labelPosArray[signal].Top = (i * (labelPosArray.Length + 1) + signal) * 30;
                    labelPosArray[signal].Width = 100;
                    panelPos.Controls.Add(labelPosArray[signal]);

                    textBoxPosArray[labelPosArray.Length * i + signal] = new TextBox();
                    textBoxPosArray[labelPosArray.Length * i + signal].Text = "0";
                    textBoxPosArray[labelPosArray.Length * i + signal].Width = 30;
                    textBoxPosArray[labelPosArray.Length * i + signal].Font = new Font("微軟正黑體", 12);
                    textBoxPosArray[labelPosArray.Length * i + signal].Left = labelPosArray[i].Right + 2;
                    textBoxPosArray[labelPosArray.Length * i + signal].Top = (i * (labelPosArray.Length + 1) + signal) * 30;
                    panelPos.Controls.Add(textBoxPosArray[labelPosArray.Length * i + signal]);

                    negativeButtonArray[signal + i * labelPosArray.Length] = new Button();
                    negativeButtonArray[signal + i * labelPosArray.Length].Font = new Font("微軟正黑體", 12);
                    negativeButtonArray[signal + i * labelPosArray.Length].Top = (i * (labelPosArray.Length + 1) + signal) * 30;
                    negativeButtonArray[signal + i * labelPosArray.Length].Visible = true;
                    negativeButtonArray[signal + i * labelPosArray.Length].Height = 28;
                    negativeButtonArray[signal + i * labelPosArray.Length].Width = 50;
                    negativeButtonArray[signal + i * labelPosArray.Length].Show();
                    negativeButtonArray[signal + i * labelPosArray.Length].Left = textBoxPosArray[i].Right + 2;
                    negativeButtonArray[signal + i * labelPosArray.Length].Text = "-";
                    panelPos.Controls.Add(negativeButtonArray[signal + i * labelPosArray.Length]);
                    negativeButtonArray[signal + i * labelPosArray.Length].Click += new System.EventHandler(delegate(object sender, EventArgs args)
                    {
                        Button temp = (Button)sender;
                        if (textBoxPosArray[Array.IndexOf(negativeButtonArray, temp)].Text == "") { textBoxPosArray[Array.IndexOf(negativeButtonArray, temp)].Text = "0"; }
                        if (Int32.Parse(textBoxPosArray[Array.IndexOf(negativeButtonArray, temp)].Text) > 0) { textBoxPosArray[Array.IndexOf(negativeButtonArray, temp)].Text = (Int32.Parse(textBoxPosArray[Array.IndexOf(negativeButtonArray, temp)].Text) - 1).ToString(); }
                    });

                    positiveButtonArray[signal + i * labelPosArray.Length] = new Button();
                    positiveButtonArray[signal + i * labelPosArray.Length].Font = new Font("微軟正黑體", 12);
                    positiveButtonArray[signal + i * labelPosArray.Length].Top = (i * (labelPosArray.Length + 1) + signal) * 30;
                    positiveButtonArray[signal + i * labelPosArray.Length].Visible = true;
                    positiveButtonArray[signal + i * labelPosArray.Length].Height = 28;
                    positiveButtonArray[signal + i * labelPosArray.Length].Width = 50;
                    positiveButtonArray[signal + i * labelPosArray.Length].Show();
                    positiveButtonArray[signal + i * labelPosArray.Length].Left = negativeButtonArray[i].Right + 2;
                    positiveButtonArray[signal + i * labelPosArray.Length].Text = "+";
                    panelPos.Controls.Add(positiveButtonArray[signal + i * labelPosArray.Length]);
                    positiveButtonArray[signal + i * labelPosArray.Length].Click += new System.EventHandler(delegate(object sender, EventArgs args)
                    {
                        Button temp = (Button)sender;
                        if (textBoxPosArray[Array.IndexOf(positiveButtonArray, temp)].Text == "") { textBoxPosArray[Array.IndexOf(positiveButtonArray, temp)].Text = "0"; }
                        textBoxPosArray[Array.IndexOf(positiveButtonArray, temp)].Text = (Int32.Parse(textBoxPosArray[Array.IndexOf(positiveButtonArray, temp)].Text) + 1).ToString();
                    });

                    deleteButtonArrayHandle[signal + i * labelPosArray.Length] = new Button();
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Font = new Font("微軟正黑體", 12);
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Top = (i * (labelPosArray.Length + 1) + signal) * 30;
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Visible = true;
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Height = 28;
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].AutoSize = true;
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Show();
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Left = positiveButtonArray[i].Right + 2;
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Visible = false;
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Text = "刪除處置";
                    panelPos.Controls.Add(deleteButtonArrayHandle[signal + i * labelPosArray.Length]);
                    deleteButtonArrayHandle[signal + i * labelPosArray.Length].Click += new System.EventHandler(delegate(object sender, EventArgs args) { Button temp = (Button)sender; if (MessageBox.Show("確定要刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) { deleteContestButtonHandleAction(temp); } });
                }
            }

        }

        private void deleteContestButtonHandleAction(Button temp)
        {
            DataTable handleNumberDelete = DBManage.getTable(conn, "SELECT * FROM injuryHandle");
            int num = handleNumberDelete.Rows.Count;
            string sql = "SELECT handleID FROM injuryHandle WHERE handle='" + labelPosArray[Array.IndexOf(deleteButtonArrayHandle, temp) % num].Text + "'";
            string handleID = DBManage.getTable(conn, sql).Rows[0]["handleID"].ToString();
            DBManage.createOrInsertCmd(conn, "DELETE FROM injuryHandle WHERE handleID=" + handleID);
            DBManage.createOrInsertCmd(conn, "DELETE FROM logHandle WHERE handleID=" + handleID);
            DataTable pos = DBManage.getTable(conn, "SELECT pos FROM injuryPos");
            myPos(panelPos, pos);
            button8.Text = "刪除部位";
            button9.Text = "刪除處置";
        }

        private void deleteContestButtonAction(Button temp)
        {
            string objectID = DBManage.getTable(conn, "SELECT objectID FROM object WHERE ObjectName='" + labelArray[Array.IndexOf(deleteButtonArray, temp)].Text + "'").Rows[0]["objectID"].ToString();
            DBManage.createOrInsertCmd(conn, "DELETE FROM object WHERE objectID=" + objectID);
            DBManage.createOrInsertCmd(conn, "DELETE FROM logObject WHERE objectID=" + objectID);
            Queue objectName = DBManage.selectCmd(conn, "SELECT objectName FROM object", "objectName");
            myObject(objectPanel, objectName);
            deleteUsedSupplies.Text = "刪除使用耗材類別";
        }

        private void numberOnly(object sender, KeyPressEventArgs e)
        {
            TextBox text = (TextBox)sender;
            text.Text = Strings.StrConv(text.Text, VbStrConv.Narrow, 0);

            if (e.KeyChar == (Char)48 || e.KeyChar == (Char)49 ||
               e.KeyChar == (Char)50 || e.KeyChar == (Char)51 ||
               e.KeyChar == (Char)52 || e.KeyChar == (Char)53 ||
               e.KeyChar == (Char)54 || e.KeyChar == (Char)55 ||
               e.KeyChar == (Char)56 || e.KeyChar == (Char)57 ||
               e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void insertPosOfInjuryButton_Click(object sender1, EventArgs e)
        {
            DataTable tmp;
            tmp = DBManage.getTable(conn, "SELECT pos FROM injuryPos ORDER BY posID");
            DataTable registerHandle = DBManage.getTable(conn, "SELECT * FROM injuryHandle");

            string[] newComboBox = new string[comboBoxPosArray.Length];
            for (int i = 0; i < newComboBox.Length; i++)
            {
                newComboBox[i] = comboBoxPosArray[i].Text;
                Console.WriteLine(newComboBox[i]);
            }

            bool[] newCheckBoxLeft = new bool[checkBoxLeftArray.Length];
            for (int i = 0; i < newCheckBoxLeft.Length; i++)
            {
                newCheckBoxLeft[i] = checkBoxLeftArray[i].Checked;
            }

            bool[] newCheckBoxRight = new bool[checkBoxRightArray.Length];
            for (int i = 0; i < newCheckBoxRight.Length; i++)
            {
                newCheckBoxRight[i] = checkBoxRightArray[i].Checked;
            }

            string[] newTextBox = new string[textBoxPosArray.Length];
            for (int i = 0; i < newTextBox.Length; i++)
            {
                newTextBox[i] = textBoxPosArray[i].Text;
            }

            comboBoxPosFlag++;
            myPos(panelPos, tmp);

            for (int i = 0; i < newComboBox.Length; i++)
            {
                comboBoxPosArray[i].Text = newComboBox[i];
            }

            for (int i = 0; i < newCheckBoxLeft.Length; i++)
            {
                checkBoxLeftArray[i].Checked = newCheckBoxLeft[i];
            }

            for (int i = 0; i < newCheckBoxRight.Length; i++)
            {
                checkBoxRightArray[i].Checked = newCheckBoxRight[i];
            }

            for (int i = 0; i < newTextBox.Length; i++)
            {
                textBoxPosArray[i].Text = newTextBox[i];
            }

            button8.Text = "刪除部位";
            button9.Text = "刪除處置";
        }

        private void insertHandleButton_Click(object sender, EventArgs e)
        {
            string value = "";
            if (Lib.InputBox("請輸入新處置", "新增處置:", ref value) == DialogResult.OK)
            {
                DBManage.createOrInsertCmd(conn, "INSERT INTO injuryHandle(handle) VALUES ('" + value + "')");
                DataTable insertID = DBManage.getTable(conn, "SELECT insertID FROM insertLog ORDER BY insertID DESC");
                DataTable handleID = DBManage.getTable(conn, "SELECT handleID FROM injuryHandle ORDER BY handleID DESC");
                for (int i = 0; i < insertID.Rows.Count; i++)
                {
                    string sql = "INSERT INTO logHandle(handleID,insertID,handleCount) VALUES (" + (Int32.Parse(handleID.Rows[0]["handleID"].ToString())) + "," + Int32.Parse(insertID.Rows[i]["insertID"].ToString()) + ",0)";
                    Console.WriteLine(sql);
                    DBManage.createOrInsertCmd(conn, sql);
                }

                DataTable tmp;
                tmp = DBManage.getTable(conn, "SELECT pos FROM injuryPos ORDER BY posID");
                myPos(panelPos, tmp);
                button9.Text = "刪除處置";
            }
        }

        private string[] getLabelText(Queue queue)
        {
            queue.Dequeue(); //釋放空值
            string[] returnText = new string[queue.Count];
            int i = 0;
            while (queue.Count > 0)
            {
                returnText[i] = queue.Dequeue().ToString();
                i++;
            }
            return returnText;
        }

        private CheckedListBox setCheckedListBox(CheckedListBox myCheckedList, Queue queue)
        {
            myCheckedList.Items.Clear();
            queue.Dequeue(); //釋放空值
            while (queue.Count > 0)
            {
                myCheckedList.Items.Add(queue.Dequeue());
            }
            return myCheckedList;
        }

        ComboBox setComboBoxPos(ComboBox myComboBox, DataTable queue)
        {
            myComboBox.Items.Clear();
            //queue.Dequeue(); //釋放空值
            for (int i = 0; i < queue.Rows.Count; i++)
            {
                myComboBox.Items.Add(queue.Rows[i]["pos"]);
            }
            return myComboBox;
        }

        private ComboBox setComboBox(ComboBox myComboBox, Queue queue)
        {
            myComboBox.Items.Clear();
            while (queue.Count > 0)
            {
                myComboBox.Items.Add(queue.Dequeue());
            }
            return myComboBox;
        }

        private void insertDbButton_Click(object sender, EventArgs e)
        {
            if (teamComboBox.Text == "" || memberComboBox.Text == "")
            {
                MessageBox.Show("尚未填入完整表格");

                if (teamComboBox.Text == "")
                {
                    teamLabel.ForeColor = Color.Red;
                }
                else
                {
                    teamLabel.ForeColor = Color.Black;
                }

                if (memberComboBox.Text == "")
                {
                    memberLabel.ForeColor = Color.Red;
                }
                else
                {
                    memberLabel.ForeColor = Color.Black;
                }
            }
            else
            {
                if (MessageBox.Show("您確定要新增?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    DBManage.createOrInsertCmd(conn, "INSERT INTO insertLog(insertTime, contestName, teamID, memberID, psText) VALUES ('" + dateTimePicker.Text + "','" + contestTextBox.Text + "','" + teamComboBox.Text + "','" + memberComboBox.Text + "','" + psTextBox.Text + "')");
                    DataTable insertID = DBManage.getTable(conn, "SELECT insertID FROM insertLog ORDER BY insertID DESC");
                    DataTable tmp = null;

                    CheckBox[] kindCheckBoxArray = new CheckBox[5];
                    kindCheckBoxArray[0] = kindCheckBox1;
                    kindCheckBoxArray[1] = kindCheckBox2;
                    kindCheckBoxArray[2] = kindCheckBox3;
                    kindCheckBoxArray[3] = kindCheckBox4;
                    kindCheckBoxArray[4] = kindCheckBox5;
                    string[] kindStringArray = { "扭傷", "拉傷", "挫傷", "脫臼", "骨折" };

                    for (int i = 0; i < kindCheckBoxArray.Length; i++)
                    {
                        if (kindCheckBoxArray[i].Checked)
                        {
                            tmp = DBManage.getTable(conn, "SELECT kindID FROM injuryKind WHERE kind='" + kindStringArray[i] + "'");
                            DBManage.createOrInsertCmd(conn, "INSERT INTO logKind(kindID, insertID) VALUES (" + tmp.Rows[0]["kindID"] + "," + insertID.Rows[0]["insertID"] + ")");
                        }
                    }

                    RadioButton[] radioButtonArray = new RadioButton[3];
                    radioButtonArray[0] = kindRadio1;
                    radioButtonArray[1] = kindRadio2;
                    radioButtonArray[2] = kindRadio3;
                    string[] categoryStringArray = { "預防", "新傷", "舊傷" };

                    for (int i = 0; i < radioButtonArray.Length; i++)
                    {
                        if (radioButtonArray[i].Checked)
                        {
                            tmp = DBManage.getTable(conn, "SELECT CategoryID FROM injuryCategory WHERE category='" + categoryStringArray[i] + "'");
                            DBManage.createOrInsertCmd(conn, "INSERT INTO logCategory(categoryID, insertID) VALUES (" + tmp.Rows[0]["categoryID"] + "," + insertID.Rows[0]["insertID"] + ")");
                            break;
                        }
                    }

                    DataTable pos = DBManage.getTable(conn, "SELECT * FROM injuryPos");
                    for (int i = 0; i < pos.Rows.Count; i++)
                    {
                        for (int j = 0; j < comboBoxPosArray.Length; j++)
                        {
                            if (pos.Rows[i]["pos"].ToString() == comboBoxPosArray[j].Text)
                            {
                                tmp = DBManage.getTable(conn, "SELECT posID FROM injuryPos WHERE pos='" + comboBoxPosArray[j].Text + "'");
                                string checkTmp = "";
                                if (checkBoxLeftArray[j].Checked == true)
                                {
                                    checkTmp += "左";
                                }
                                if (checkBoxRightArray[j].Checked == true)
                                {
                                    checkTmp += "右";
                                }
                                DBManage.createOrInsertCmd(conn, "INSERT INTO logPos(posID, insertID, side) VALUES (" + tmp.Rows[0]["posID"] + "," + insertID.Rows[0]["insertID"] + ",'" + checkTmp + "')");
                            }
                        }
                    }

                    int[] sumTemp = new int[labelPosArray.Length];
                    for (int i = 0; i < textBoxPosArray.Length; i++)
                    {
                        DataTable tmpHandleNumber = DBManage.getTable(conn, "SELECT * FROM injuryHandle");
                        tmp = DBManage.getTable(conn, "SELECT handleID FROM injuryHandle WHERE handle='" + labelPosArray[i % (tmpHandleNumber.Rows.Count)].Text + "'");
                        sumTemp[i % (tmpHandleNumber.Rows.Count)] += Int32.Parse(textBoxPosArray[i].Text);
                        DBManage.createOrInsertCmd(conn, "INSERT INTO logHandle(handleID, insertID, handleCount) VALUES (" + tmp.Rows[0]["handleID"] + "," + insertID.Rows[0]["insertID"] + ",'" + textBoxPosArray[i].Text + "')");
                    }

                    for (int i = 0; i < labelArray.Length; i++)
                    {
                        tmp = DBManage.getTable(conn, "SELECT objectID FROM object WHERE objectName='" + labelArray[i].Text + "'");
                        if (checkBoxArray[i].Checked == true)
                        {
                            DBManage.createOrInsertCmd(conn, "INSERT INTO logObject(objectID, insertID, number) VALUES (" + tmp.Rows[0]["objectID"] + "," + insertID.Rows[0]["insertID"] + ",1)");
                        }
                        else
                        {
                            DBManage.createOrInsertCmd(conn, "INSERT INTO logObject(objectID, insertID, number) VALUES (" + tmp.Rows[0]["objectID"] + "," + insertID.Rows[0]["insertID"] + ",0)");
                        }
                    }

                    dataGridViewStatistic = getStatisticView(dataGridViewStatistic, comboBoxKindTab3.Text, comboBoxTeamTab3.Text);
                    MessageBox.Show("輸入完成");

                    teamComboBox.SelectedIndex = -1;
                    memberComboBox.SelectedIndex = -1;
                    contestTextBox.Text = "";
                    psTextBox.Text = "";
                    kindCheckBox1.Checked = false;
                    kindCheckBox2.Checked = false;
                    kindCheckBox3.Checked = false;
                    kindCheckBox4.Checked = false;
                    kindCheckBox5.Checked = false;
                    kindRadio1.Checked = false;
                    kindRadio2.Checked = false;
                    kindRadio3.Checked = false;

                    comboBoxPosFlag = 0;
                    teamLabel.ForeColor = Color.Black;
                    memberLabel.ForeColor = Color.Black;

                    DataTable posIni = DBManage.getTable(conn, "SELECT pos FROM injuryPos");
                    myPos(panelPos, posIni);
                    Queue objectIni = DBManage.selectCmd(conn, "SELECT objectName FROM object", "objectName");
                    myObject(objectPanel, objectIni);
                    dateTimePicker.Text = DateTime.Now.ToString("yyyy/MM/dd");
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage2)
            {
                textBoxSchoolYear.ImeMode = System.Windows.Forms.ImeMode.Hangul;
                textBoxSchoolYear.Text = (Int32.Parse(DateTime.Now.Year.ToString()) - 1912).ToString();
            }
        }

        private void comboBoxTeamTab3_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridViewStatistic = getStatisticView(dataGridViewStatistic, comboBoxKindTab3.Text, comboBoxTeamTab3.Text);
        }

        private void comboBoxKindTab3_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridViewStatistic = getStatisticView(dataGridViewStatistic, comboBoxKindTab3.Text, comboBoxTeamTab3.Text);
        }

        private DataGridView getStatisticView(DataGridView rawView, string selectedItem, string teamItem)
        {
            rawView.Rows.Clear();
            rawView.ColumnHeadersDefaultCellStyle.Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            rawView.DefaultCellStyle.Font = new Font("微軟正黑體", 12);
            rawView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            if (teamItem == "全部")
            {
                if (selectedItem == "耗材")
                {
                    DataTable objectName = DBManage.getTable(conn, "SELECT objectName FROM object ORDER BY objectID");
                    int objectNameCount = objectName.Rows.Count;

                    rawView.ColumnCount = 1 + objectNameCount;
                    rawView.Columns[0].HeaderText = "月份";

                    for (int i = 1; i < 1 + objectNameCount; i++)
                    {
                        rawView.Columns[i].HeaderText = objectName.Rows[i - 1]["objectName"].ToString();
                    }

                    for (int i = 0; i < rawView.ColumnCount; i++)
                    {
                        rawView.Columns[i].Width = 130;
                    }

                    DataTable dateInsertID = DBManage.getTable(conn, "SELECT substr(insertTime,1,7) AS time FROM logObject,insertLog WHERE logObject.insertId=insertLog.insertId GROUP BY substr(insertTime,1,7)");
                    for (int i = 0; i < dateInsertID.Rows.Count; i++)
                    {
                        rawView.Rows.Add(dateInsertID.Rows[i]["time"]);
                    }
                    for (int i = 0; i < dateInsertID.Rows.Count; i++)
                    {
                        for (int j = 0; j < objectNameCount; j++)
                        {
                            DataTable objectInsertID = DBManage.getTable(conn, "SELECT substr(insertTime,1,7),sum(number) as number FROM logObject,insertLog WHERE logObject.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + dateInsertID.Rows[i]["time"] + "' GROUP BY substr(insertTime,1,7),objectID order by  substr(insertTime,1,7)");
                            try
                            {
                                rawView.Rows[i].Cells[j + 1].Value = objectInsertID.Rows[j]["number"];
                            }
                            catch (System.IndexOutOfRangeException)
                            {
                                rawView.Rows[i].Cells[j + 1].Value = 0;
                            }

                        }
                    }
                }
                else if (selectedItem == "處置")
                {
                    DataTable handle = DBManage.getTable(conn, "SELECT * FROM injuryHandle ORDER BY handleID");
                    int handleCount = handle.Rows.Count;

                    rawView.ColumnCount = 1 + handleCount;
                    rawView.Columns[0].HeaderText = "月份";

                    for (int i = 1; i < 1 + handleCount; i++)
                    {
                        rawView.Columns[i].HeaderText = handle.Rows[i - 1]["handle"].ToString();
                    }

                    for (int i = 0; i < rawView.ColumnCount; i++)
                    {
                        rawView.Columns[i].Width = 130;
                    }

                    DataTable insertTime = DBManage.getTable(conn, "SELECT substr(insertTime,1,7) AS time FROM logHandle,insertLog WHERE logHandle.insertId=insertLog.insertId GROUP BY substr(insertTime,1,7)");
                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        rawView.Rows.Add(insertTime.Rows[i]["time"]);
                    }

                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        DataTable handleID = DBManage.getTable(conn, "SELECT sum(handleCount) AS number FROM logHandle,insertLog WHERE logHandle.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + insertTime.Rows[i]["time"] + "' GROUP BY substr(insertTime,1,7),handleID order by  substr(insertTime,1,7)");
                        for (int j = 0; j < handleCount; j++)
                        {
                            rawView.Rows[i].Cells[j + 1].Value = handleID.Rows[j]["number"];
                        }
                    }
                }
                else if (selectedItem == "部位")
                {
                    DataTable pos = DBManage.getTable(conn, "SELECT * FROM injuryPos ORDER BY posID");
                    int posCount = pos.Rows.Count;

                    rawView.ColumnCount = 1 + posCount;
                    rawView.Columns[0].HeaderText = "月份";

                    for (int i = 1; i < 1 + posCount; i++)
                    {
                        rawView.Columns[i].HeaderText = pos.Rows[i - 1]["pos"].ToString();
                    }

                    for (int i = 0; i < rawView.ColumnCount; i++)
                    {
                        rawView.Columns[i].Width = 130;
                    }

                    DataTable insertTime = DBManage.getTable(conn, "SELECT substr(insertTime,1,7) AS time FROM logPos,insertLog WHERE logPos.insertId=insertLog.insertId GROUP BY substr(insertTime,1,7)");
                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        rawView.Rows.Add(insertTime.Rows[i]["time"]);
                    }

                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        DataTable posID = DBManage.getTable(conn, "SELECT posID,count(posID) AS number FROM logPos,insertLog WHERE logPos.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + insertTime.Rows[i]["time"] + "' AND (side='左' OR side='右') GROUP BY substr(insertTime,1,7),posID order by substr(insertTime,1,7)");
                        for (int j = 0; j < posCount; j++)
                        {
                            rawView.Rows[i].Cells[j + 1].Value = 0;
                        }

                        for (int j = 0; j < posID.Rows.Count; j++)
                        {
                            try
                            {
                                rawView.Rows[i].Cells[Int32.Parse(posID.Rows[j]["posID"].ToString())].Value = posID.Rows[j]["number"];
                            }
                            catch (Exception ex)
                            {

                            }
                        }

                        DataTable posID2 = DBManage.getTable(conn, "SELECT posID,count(posID)*2 AS number FROM logPos,insertLog WHERE logPos.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + insertTime.Rows[i]["time"] + "' AND side='左右' GROUP BY substr(insertTime,1,7),posID order by substr(insertTime,1,7)");
                        for (int j = 0; j < posID2.Rows.Count; j++)
                        {
                            try
                            {
                                rawView.Rows[i].Cells[Int32.Parse(posID2.Rows[j]["posID"].ToString())].Value = Int32.Parse(rawView.Rows[i].Cells[Int32.Parse(posID2.Rows[j]["posID"].ToString())].Value.ToString()) + Int32.Parse(posID2.Rows[j]["number"].ToString());
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
            }
            else
            {
                if (selectedItem == "耗材")
                {
                    DataTable objectName = DBManage.getTable(conn, "SELECT objectName FROM object ORDER BY objectID");
                    int objectNameCount = objectName.Rows.Count;

                    rawView.ColumnCount = 1 + objectNameCount;
                    rawView.Columns[0].HeaderText = "月份";

                    for (int i = 1; i < 1 + objectNameCount; i++)
                    {
                        rawView.Columns[i].HeaderText = objectName.Rows[i - 1]["objectName"].ToString();
                    }

                    for (int i = 0; i < rawView.ColumnCount; i++)
                    {
                        rawView.Columns[i].Width = 130;
                    }

                    DataTable dateInsertID = DBManage.getTable(conn, "SELECT substr(insertTime,1,7) AS time FROM logObject,insertLog WHERE logObject.insertId=insertLog.insertId AND teamID='" + comboBoxTeamTab3.Text + "' GROUP BY substr(insertTime,1,7)");
                    for (int i = 0; i < dateInsertID.Rows.Count; i++)
                    {
                        rawView.Rows.Add(dateInsertID.Rows[i]["time"]);
                    }

                    for (int i = 0; i < dateInsertID.Rows.Count; i++)
                    {
                        for (int j = 0; j < objectNameCount; j++)
                        {
                            DataTable objectInsertID = DBManage.getTable(conn, "SELECT substr(insertTime,1,7),sum(number) as number FROM logObject,insertLog WHERE logObject.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + dateInsertID.Rows[i]["time"] + "' AND teamID='" + comboBoxTeamTab3.Text + "' GROUP BY substr(insertTime,1,7),objectID order by  substr(insertTime,1,7)");
                            try
                            {
                                rawView.Rows[i].Cells[j + 1].Value = objectInsertID.Rows[j]["number"];
                            }
                            catch (System.IndexOutOfRangeException)
                            {
                                rawView.Rows[i].Cells[j + 1].Value = 0;
                            }

                        }
                    }
                }
                else if (selectedItem == "處置")
                {
                    DataTable handle = DBManage.getTable(conn, "SELECT * FROM injuryHandle ORDER BY handleID");
                    int handleCount = handle.Rows.Count;

                    rawView.ColumnCount = 1 + handleCount;
                    rawView.Columns[0].HeaderText = "月份";

                    for (int i = 1; i < 1 + handleCount; i++)
                    {
                        rawView.Columns[i].HeaderText = handle.Rows[i - 1]["handle"].ToString();
                    }

                    for (int i = 0; i < rawView.ColumnCount; i++)
                    {
                        rawView.Columns[i].Width = 130;
                    }

                    string sql = "SELECT substr(insertTime,1,7) AS time FROM logHandle,insertLog WHERE logHandle.insertId=insertLog.insertId AND teamID='" + comboBoxTeamTab3.Text + "' GROUP BY substr(insertTime,1,7)";
                    Console.WriteLine(sql);
                    DataTable insertTime = DBManage.getTable(conn, sql);
                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        rawView.Rows.Add(insertTime.Rows[i]["time"]);
                    }

                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        DataTable handleID = DBManage.getTable(conn, "SELECT sum(handleCount) AS number FROM logHandle,insertLog WHERE logHandle.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + insertTime.Rows[i]["time"] + "' AND teamID='" + comboBoxTeamTab3.Text + "' GROUP BY substr(insertTime,1,7),handleID order by  substr(insertTime,1,7)");
                        for (int j = 0; j < handleCount; j++)
                        {
                            rawView.Rows[i].Cells[j + 1].Value = handleID.Rows[j]["number"];
                        }
                    }
                }
                else if (selectedItem == "部位")
                {
                    DataTable pos = DBManage.getTable(conn, "SELECT * FROM injuryPos ORDER BY posID");
                    int posCount = pos.Rows.Count;

                    rawView.ColumnCount = 1 + posCount;
                    rawView.Columns[0].HeaderText = "月份";

                    for (int i = 1; i < 1 + posCount; i++)
                    {
                        rawView.Columns[i].HeaderText = pos.Rows[i - 1]["pos"].ToString();
                    }

                    for (int i = 0; i < rawView.ColumnCount; i++)
                    {
                        rawView.Columns[i].Width = 130;
                    }

                    DataTable insertTime = DBManage.getTable(conn, "SELECT substr(insertTime,1,7) AS time FROM logPos,insertLog WHERE logPos.insertId=insertLog.insertId AND teamID='" + comboBoxTeamTab3.Text + "' GROUP BY substr(insertTime,1,7)");
                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        rawView.Rows.Add(insertTime.Rows[i]["time"]);
                    }

                    for (int i = 0; i < insertTime.Rows.Count; i++)
                    {
                        DataTable posID = DBManage.getTable(conn, "SELECT posID,count(posID) AS number FROM logPos,insertLog WHERE logPos.insertID=insertLog.insertID AND substr(insertTime,1,7)='" + insertTime.Rows[i]["time"] + "' AND teamID='" + comboBoxTeamTab3.Text + "' GROUP BY substr(insertTime,1,7),posID order by substr(insertTime,1,7)");
                        for (int j = 0; j < posCount; j++)
                        {
                            rawView.Rows[i].Cells[j + 1].Value = 0;
                        }

                        for (int j = 0; j < posID.Rows.Count; j++)
                        {
                            try
                            {
                                rawView.Rows[i].Cells[Int32.Parse(posID.Rows[j]["posID"].ToString())].Value = posID.Rows[j]["number"];
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
            }

            return rawView;
        }

        private void setComboBoxTab3(ComboBox rawComboBox, string type)
        {
            rawComboBox.Items.Clear();
            rawComboBox.Items.Add("全部");
            DataTable tmp = new DataTable();
            if (type == "team")
            {
                tmp = DBManage.getTable(conn, "SELECT teamName FROM team");
            }

            for (int i = 0; i < tmp.Rows.Count; i++)
            {
                rawComboBox.Items.Add(tmp.Rows[i]["teamName"]);
            }
            rawComboBox.Text = "全部";
        }

        private string getSelectString(DataTable tmpTable, string type, string tableType, int id)
        {
            string tmpString = "";
            DataTable tmpPos = DBManage.getTable(conn, "SELECT " + type + "ID FROM log" + tableType + " WHERE log" + tableType + ".insertID=" + tmpTable.Rows[id]["insertID"]);
            for (int j = 0; j < tmpPos.Rows.Count; j++)
            {
                DataTable tmp = DBManage.getTable(conn, "SELECT " + type + " AS tmp FROM injury" + tableType + " WHERE " + type + "ID=" + tmpPos.Rows[j][type + "ID"]);
                tmpString += tmp.Rows[0]["tmp"];
            }
            return tmpString;
        }

        private void contestButton_Click(object sender, EventArgs e)
        {
            dataGridViewContest.Columns.Clear();
            dataGridViewContest.Rows.Clear();
            dataGridViewContest.ColumnHeadersDefaultCellStyle.Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            dataGridViewContest.DefaultCellStyle.Font = new Font("微軟正黑體", 12);
            dataGridViewContest.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            DataTable objectName = DBManage.getTable(conn, "SELECT objectName FROM object ORDER BY objectID");
            objectNameCount = objectName.Rows.Count;

            DataTable handle = DBManage.getTable(conn, "SELECT handle FROM injuryHandle ORDER BY handleID");
            handleCount = handle.Rows.Count;

            DataTable kind = DBManage.getTable(conn, "SELECT kind FROM injuryKind ORDER BY kindID");
            kindCount = kind.Rows.Count;

            DataTable category = DBManage.getTable(conn, "SELECT category FROM injuryCategory ORDER BY categoryID");
            categoryCount = category.Rows.Count;

            DataTable side = DBManage.getTable(conn, "SELECT side FROM injurySide ORDER BY sideID");
            sideCount = side.Rows.Count;

            DataTable pos = DBManage.getTable(conn, "SELECT pos FROM injuryPos ORDER BY posID");
            posCount = pos.Rows.Count;

            dataGridViewContest.ColumnCount = 5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount + posCount;
            for (int i = 0; i < dataGridViewContest.ColumnCount; i++)
            {
                dataGridViewContest.Columns[i].Width = 130;
            }

            dataGridViewContest.Columns[0].HeaderText = "人員";
            dataGridViewContest.Columns[1].HeaderText = "隊伍";
            dataGridViewContest.Columns[2].HeaderText = "比賽";
            dataGridViewContest.Columns[3].HeaderText = "日期";

            for (int i = 4; i < 4 + objectNameCount; i++)
            {
                dataGridViewContest.Columns[i].HeaderText = objectName.Rows[i - 4]["objectName"].ToString();
            }
            dataGridViewContest.Columns[4 + objectNameCount].HeaderText = "insertID";
            dataGridViewContest.Columns[4 + objectNameCount].Visible = false;

            for (int i = (5 + objectNameCount); i < (5 + objectNameCount + handleCount); i++)
            {
                dataGridViewContest.Columns[i].HeaderText = handle.Rows[i - (5 + objectNameCount)]["handle"].ToString();
            }

            for (int i = (5 + objectNameCount + handleCount); i < (5 + objectNameCount + handleCount + kindCount); i++)
            {
                dataGridViewContest.Columns[i].HeaderText = kind.Rows[i - (5 + objectNameCount + handleCount)]["kind"].ToString();
            }

            for (int i = (5 + objectNameCount + handleCount + kindCount); i < (5 + objectNameCount + handleCount + kindCount + categoryCount); i++)
            {
                dataGridViewContest.Columns[i].HeaderText = category.Rows[i - (5 + objectNameCount + handleCount + kindCount)]["category"].ToString();
            }

            for (int i = (5 + objectNameCount + handleCount + kindCount + categoryCount); i < (5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount); i++)
            {
                dataGridViewContest.Columns[i].HeaderText = side.Rows[i - (5 + objectNameCount + handleCount + kindCount + categoryCount)]["side"].ToString();
                dataGridViewContest.Columns[i].Visible = false;
            }

            for (int i = (5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount); i < (5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount + posCount); i++)
            {
                dataGridViewContest.Columns[i].HeaderText = pos.Rows[i - (5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount)]["pos"].ToString();
            }

            if (comboBoxTab2.Text == "使用耗材")
            {
                for (int i = 5 + objectNameCount; i < dataGridViewContest.ColumnCount; i++)
                {
                    dataGridViewContest.Columns[i].Visible = false;
                }
            }
            else if (comboBoxTab2.Text == "受傷種類/處置")
            {
                for (int i = 4; i < objectNameCount + 4; i++)
                {
                    dataGridViewContest.Columns[i].Visible = false;
                }

                for (int i = 5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount; i < 5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount + posCount; i++)
                {
                    dataGridViewContest.Columns[i].Visible = false;
                }
            }
            else if (comboBoxTab2.Text == "受傷部位")
            {
                for (int i = 4; i < 5 + objectNameCount + handleCount + kindCount + categoryCount + sideCount; i++)
                {
                    dataGridViewContest.Columns[i].Visible = false;
                }
            }

            DataTable contest = DBManage.getTable(conn, "SELECT * FROM insertLog WHERE insertTime BETWEEN '" + startTimePicker.Text + "' and '" + endTimePicker.Text + "' ORDER BY insertID");
            for (int i = 0; i < contest.Rows.Count; i++)
            {
                dataGridViewContest.Rows.Add();
                DataTable number = DBManage.getTable(conn, "SELECT number FROM logObject WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY objectID");
                dataGridViewContest.Rows[i].Cells[2].Value = contest.Rows[i]["contestName"];
                dataGridViewContest.Rows[i].Cells[3].Value = contest.Rows[i]["insertTime"];

                var teamID = contest.Rows[i]["teamID"];
                dataGridViewContest.Rows[i].Cells[1].Value = teamID;

                var memberID = contest.Rows[i]["memberID"];
                dataGridViewContest.Rows[i].Cells[0].Value = memberID;

                DataTable objectTimes = DBManage.getTable(conn, "SELECT * FROM object");
                for (int j = 0; j < objectTimes.Rows.Count; j++)
                {
                    try
                    {
                        DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                        tCell.Items.Add("X");
                        tCell.Items.Add("V");

                        dataGridViewContest[j + 4, i] = tCell;
                        if (number.Rows[j]["number"].ToString() == "0")
                        {
                            dataGridViewContest.Rows[i].Cells[j + 4].Value = "X";
                        }
                        else
                        {
                            dataGridViewContest.Rows[i].Cells[j + 4].Value = "V";
                        }
                    }
                    catch (System.IndexOutOfRangeException)
                    {
                        dataGridViewContest.Rows[i].Cells[j + 4].Value = 0;
                    }
                }

                dataGridViewContest.Rows[i].Cells[4 + objectNameCount].Value = contest.Rows[i]["insertID"]; //insertID
                DataTable handleNumber = DBManage.getTable(conn, "SELECT handleID FROM logHandle WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY handleID");
                DataTable handleCountText = DBManage.getTable(conn, "SELECT handleID,sum(handleCount) FROM logHandle WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY handleID");

                for (int j = 0; j < handleCount; j++)
                {
                    try
                    {
                        dataGridViewContest[j + 5 + objectNameCount, i].Value = "0";
                    }
                    catch (Exception ex)
                    {

                    }
                }

                for (int j = 0; j < handleCount; j++)
                {
                    try
                    {
                        string sql = "SELECT handleID,sum(handleCount) AS handleCount FROM logHandle WHERE insertID=" + contest.Rows[i]["insertID"] + " GROUP BY handleID ORDER BY handleID";
                        DataTable handleSum = DBManage.getTable(conn, sql);

                        string handleLast = handleSum.Rows[j]["handleCount"].ToString();
                        dataGridViewContest[j + 5 + objectNameCount, i].Value = handleLast;
                    }
                    catch (Exception ex)
                    {

                    }

                }

                DataTable kindNumber = DBManage.getTable(conn, "SELECT kindID FROM logKind WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY kindID");

                for (int j = 0; j < kindCount; j++)
                {
                    try
                    {
                        DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                        tCell.Items.Add("X");
                        tCell.Items.Add("V");

                        dataGridViewContest[j + 5 + objectNameCount + handleCount, i] = tCell;
                        dataGridViewContest[j + 5 + objectNameCount + handleCount, i].Value = "X";
                    }
                    catch (Exception ex)
                    {

                    }
                }

                for (int j = 0; j < kindNumber.Rows.Count; j++)
                {
                    try
                    {
                        int kindNumberInt = Int32.Parse(kindNumber.Rows[j]["kindID"].ToString());
                        dataGridViewContest[kindNumberInt + handleCount + 4 + objectNameCount, i].Value = "V";
                    }
                    catch (Exception ex)
                    {

                    }
                }

                DataTable categoryNumber = DBManage.getTable(conn, "SELECT categoryID FROM logCategory WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY categoryID");

                for (int j = 0; j < categoryCount; j++)
                {
                    try
                    {
                        DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                        tCell.Items.Add("X");
                        tCell.Items.Add("V");

                        dataGridViewContest[j + 5 + kindCount + objectNameCount + handleCount, i] = tCell;
                        dataGridViewContest[j + 5 + kindCount + objectNameCount + handleCount, i].Value = "X";
                    }
                    catch (Exception ex)
                    {

                    }
                }

                for (int j = 0; j < categoryNumber.Rows.Count; j++)
                {
                    try
                    {
                        int categoryNumberInt = Int32.Parse(categoryNumber.Rows[j]["categoryID"].ToString());
                        dataGridViewContest[categoryNumberInt + kindCount + handleCount + 4 + objectNameCount, i].Value = "V";
                    }
                    catch (Exception ex)
                    {

                    }

                }

                DataTable posNumber = DBManage.getTable(conn, "SELECT posID FROM logPos WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY posID");
                DataTable posSide = DBManage.getTable(conn, "SELECT * FROM logPos WHERE insertID=" + contest.Rows[i]["insertID"] + " ORDER BY posID");

                for (int j = 0; j < posCount; j++)
                {
                    try
                    {
                        DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                        tCell.Items.Add("-");
                        tCell.Items.Add("左");
                        tCell.Items.Add("右");
                        tCell.Items.Add("左右");
                        dataGridViewContest[j + 5 + kindCount + objectNameCount + handleCount + categoryCount + sideCount, i] = tCell;
                        dataGridViewContest[j + 5 + kindCount + objectNameCount + handleCount + categoryCount + sideCount, i].Value = "-";
                    }
                    catch (Exception ex)
                    {

                    }
                }

                for (int j = 0; j < posNumber.Rows.Count; j++)
                {
                    try
                    {
                        int posNumberInt = Int32.Parse(posNumber.Rows[j]["posID"].ToString());
                        string posSideText = posSide.Rows[j]["side"].ToString();

                        DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                        tCell.Items.Add("-");
                        tCell.Items.Add("左");
                        tCell.Items.Add("右");
                        tCell.Items.Add("左右");
                        dataGridViewContest[posNumberInt + sideCount + categoryCount + kindCount + handleCount + 4 + objectNameCount, i] = tCell;
                        dataGridViewContest[posNumberInt + sideCount + categoryCount + kindCount + handleCount + 4 + objectNameCount, i].Value = posSideText;
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
        }

        private void changePassButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("您確定要修改密碼嗎?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (oldPassword.Text != "" && newPassword.Text != "" && (newRePassword.Text != ""))
                {
                    if (oldPassword.Text == LoginLib.checkPassword(connLogin))
                    {
                        if (newPassword.Text == newRePassword.Text)
                        {
                            try
                            {
                                connLogin.Open();
                                SQLiteCommand cmd = connLogin.CreateCommand();
                                cmd.CommandText = "UPDATE passwordTable SET loginpassword='" + newPassword.Text + "'";
                                cmd.ExecuteNonQuery();
                                connLogin.Close();
                                MessageBox.Show("密碼已變更");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("修改密碼錯誤，原因：" + ex);
                            }
                            oldPassword.Text = "";
                            newPassword.Text = "";
                            newRePassword.Text = "";
                        }
                        else
                        {
                            MessageBox.Show("新密碼不一致");
                            newPassword.Text = "";
                            newRePassword.Text = "";
                            newRePassword.Select();
                        }
                    }
                    else
                    {
                        MessageBox.Show("舊密碼輸入錯誤");
                        oldPassword.Text = "";
                        oldPassword.Select();
                    }
                }
                else
                {
                    MessageBox.Show("請勿輸入空值");
                }
            }

        }

        private Panel myGood(Panel myPanel, DataTable good)
        {
            labelArray2 = new Label[good.Rows.Count];
            textBoxArray2 = new TextBox[good.Rows.Count];
            negativeButtonArray2 = new Button[good.Rows.Count];
            positiveButtonArray2 = new Button[good.Rows.Count];
            deleteButtonArray2 = new Button[good.Rows.Count];
            panelInsertGood.Controls.Clear();

            for (int signal = 0; signal < labelArray2.Length; signal++)
            {
                labelArray2[signal] = new Label();
                labelArray2[signal].Text = good.Rows[signal]["name"].ToString();
                labelArray2[signal].Font = new Font("微軟正黑體", 12);
                labelArray2[signal].Top = signal * 30;
                myPanel.Controls.Add(labelArray2[signal]);

                textBoxArray2[signal] = new TextBox();
                textBoxArray2[signal].Font = new Font("微軟正黑體", 12);
                textBoxArray2[signal].Top = signal * 30;
                textBoxArray2[signal].Left = labelArray2[signal].Width + 2;
                textBoxArray2[signal].Width = 40;
                textBoxArray2[signal].Text = "0";
                myPanel.Controls.Add(textBoxArray2[signal]);
                textBoxArray2[signal].ImeMode = System.Windows.Forms.ImeMode.Hangul;
                textBoxArray2[signal].KeyPress += new KeyPressEventHandler(numberOnly);

                negativeButtonArray2[signal] = new Button();
                negativeButtonArray2[signal].Font = new Font("微軟正黑體", 12);
                negativeButtonArray2[signal].Top = signal * 30;
                negativeButtonArray2[signal].Visible = true;
                negativeButtonArray2[signal].Height = 28;
                negativeButtonArray2[signal].Width = 52;
                negativeButtonArray2[signal].Show();
                negativeButtonArray2[signal].Left = labelArray2[signal].Width + textBoxArray2[signal].Width + 10;
                negativeButtonArray2[signal].Text = "-";
                myPanel.Controls.Add(negativeButtonArray2[signal]);
                negativeButtonArray2[signal].Click += new System.EventHandler(delegate(object sender, EventArgs args)
                {
                    Button temp = (Button)sender;
                    if (textBoxArray2[Array.IndexOf(negativeButtonArray2, temp)].Text == "") { textBoxArray2[Array.IndexOf(negativeButtonArray2, temp)].Text = "0"; }
                    if (Int32.Parse(textBoxArray2[Array.IndexOf(negativeButtonArray2, temp)].Text) > 0) { textBoxArray2[Array.IndexOf(negativeButtonArray2, temp)].Text = (Int32.Parse(textBoxArray2[Array.IndexOf(negativeButtonArray2, temp)].Text) - 1).ToString(); }
                });

                positiveButtonArray2[signal] = new Button();
                positiveButtonArray2[signal].Font = new Font("微軟正黑體", 12);
                positiveButtonArray2[signal].Top = signal * 30;
                positiveButtonArray2[signal].Visible = true;
                positiveButtonArray2[signal].Height = 28;
                positiveButtonArray2[signal].Width = 52;
                positiveButtonArray2[signal].Show();
                positiveButtonArray2[signal].Left = labelArray2[signal].Width + textBoxArray2[signal].Width + negativeButtonArray2[signal].Width + 10;
                positiveButtonArray2[signal].Text = "+";
                myPanel.Controls.Add(positiveButtonArray2[signal]);
                positiveButtonArray2[signal].Click += new System.EventHandler(delegate(object sender, EventArgs args)
                {
                    Button temp = (Button)sender;
                    if (textBoxArray2[Array.IndexOf(positiveButtonArray2, temp)].Text == "") { textBoxArray2[Array.IndexOf(positiveButtonArray2, temp)].Text = "0"; }
                    textBoxArray2[Array.IndexOf(positiveButtonArray2, temp)].Text = (Int32.Parse(textBoxArray2[Array.IndexOf(positiveButtonArray2, temp)].Text) + 1).ToString();
                });

                deleteButtonArray2[signal] = new Button();
                deleteButtonArray2[signal].Font = new Font("微軟正黑體", 12);
                deleteButtonArray2[signal].Top = signal * 30;
                deleteButtonArray2[signal].Visible = true;
                deleteButtonArray2[signal].Height = 28;
                deleteButtonArray2[signal].Width = 100;
                deleteButtonArray2[signal].Show();
                deleteButtonArray2[signal].Left = labelArray2[signal].Width + textBoxArray2[signal].Width + negativeButtonArray2[signal].Width + positiveButtonArray2[signal].Width + 10;
                deleteButtonArray2[signal].Text = "刪除";
                myPanel.Controls.Add(deleteButtonArray2[signal]);
                deleteButtonArray2[signal].Click += new System.EventHandler(delegate(object sender, EventArgs args) { Button temp = (Button)sender; if (MessageBox.Show("確定要刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) { deleteButtonAction(temp); getDataGridView1(); } });

            }

            Button insertNew = new Button();
            insertNew.Font = new Font("微軟正黑體", 16);
            insertNew.Top = labelArray2.Length * 30;
            insertNew.Visible = true;
            insertNew.Height = 40;
            insertNew.Width = 212;
            insertNew.Show();
            insertNew.Left = 2;
            insertNew.Top = 401;
            insertNew.Text = "新增類別";
            insertNew.Click += new System.EventHandler(delegate(object sender, EventArgs args) { insertLabelArray2(); getDataGridView1(); });

            groupBox6.Controls.Add(insertNew);
            return myPanel;
        }

        private void deleteButtonAction(Button temp)
        {
            string goodID = DBManage.getTable(conn, "SELECT goodID FROM good WHERE name='" + labelArray2[Array.IndexOf(deleteButtonArray2, temp)].Text + "'").Rows[0]["goodID"].ToString();
            DBManage.createOrInsertCmd(conn, "DELETE FROM good WHERE goodID=" + goodID);
            DBManage.createOrInsertCmd(conn, "DELETE FROM goodLog WHERE goodID=" + goodID);
            DataTable goodName = DBManage.getTable(conn, "SELECT name FROM good");
            myGood(panelInsertGood, goodName);
        }

        private void insertObjectButton_Click(object sender, EventArgs e)
        {
            insertLabelArray();
        }

        private void insertLabelArray()
        {
            string value = "";
            if (Lib.InputBox("請輸入新耗材", "新增耗材:", ref value) == DialogResult.OK)
            {
                DBManage.createOrInsertCmd(conn, "INSERT INTO object(objectName) VALUES ('" + value + "')");
                DataTable insertID = DBManage.getTable(conn, "SELECT insertID FROM insertLog ORDER BY insertID DESC");
                DataTable objectID = DBManage.getTable(conn, "SELECT objectID FROM object ORDER BY objectID DESC");
                for (int i = 0; i < insertID.Rows.Count; i++)
                {
                    string sql = "INSERT INTO logObject(objectID,insertID,number) VALUES (" + (Int32.Parse(objectID.Rows[0]["objectID"].ToString())) + "," + Int32.Parse(insertID.Rows[i]["insertID"].ToString()) + ",0)";
                    Console.WriteLine(sql);
                    DBManage.createOrInsertCmd(conn, sql);
                }


                Queue tmp = DBManage.selectCmd(conn, "SELECT objectName FROM object", "objectName");
                myObject(objectPanel, tmp);
                deleteUsedSupplies.Text = "刪除使用耗材類別";
            }
        }

        private void insertLabelArray2()
        {
            string value = "";
            if (Lib.InputBox("請輸入新貼布或新外傷用品", "新增貼布或外傷用品:", ref value) == DialogResult.OK)
            {
                DBManage.createOrInsertCmd(conn, "INSERT INTO good(name) VALUES ('" + value + "')");
                DataTable goodName = DBManage.getTable(conn, "SELECT name FROM good");
                panelInsertGood = myGood(panelInsertGood, goodName);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DBManage.createOrInsertCmd(conn, "INSERT INTO insertGoodLog(insertMonth) VALUES ('" + textBoxSchoolYear.Text + comboBoxSemester.Text + "')");
            DataTable IGID = DBManage.getTable(conn, "SELECT IGID FROM insertGoodLog ORDER BY IGID DESC");
            for (int i = 0; i < textBoxArray2.Length; i++)
            {
                DataTable goodID = DBManage.getTable(conn, "SELECT goodID FROM good WHERE name='" + labelArray2[i].Text + "'");
                DBManage.createOrInsertCmd(conn, "INSERT INTO goodLog(goodID, number, IGID) VALUES ('" + goodID.Rows[0]["goodID"] + "','" + textBoxArray2[i].Text + "'," + IGID.Rows[0]["IGID"] + ")");
            }

            getDataGridView1();

            for (int i = 0; i < textBoxArray2.Length; i++)
            {
                textBoxArray2[i].Text = "0";
            }
        }

        private void getDataGridView1()
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.Font = new Font("微軟正黑體", 12);
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            DataTable goodName = DBManage.getTable(conn, "SELECT name FROM good ORDER BY goodID");
            DataTable insertGoodLog = DBManage.getTable(conn, "SELECT * FROM insertGoodLog ORDER BY IGID");
            dataGridView1.ColumnCount = 2 + goodName.Rows.Count;

            dataGridView1.Columns[0].HeaderText = "學年";
            for (int i = 1; i < dataGridView1.ColumnCount - 1; i++)
            {
                dataGridView1.Columns[i].HeaderText = goodName.Rows[i - 1]["name"].ToString();
            }

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].Width = 140;
            }

            for (int i = 0; i < insertGoodLog.Rows.Count; i++)
            {
                dataGridView1.Rows.Add(insertGoodLog.Rows[i]["insertMonth"]);
                for (int j = 0; j < dataGridView1.ColumnCount - 2; j++)
                {
                    var debug = dataGridView1.Columns[j + 1].HeaderText;
                    DataTable goodID = DBManage.getTable(conn, "SELECT goodID FROM good WHERE name='" + dataGridView1.Columns[j + 1].HeaderText + "'");
                    string sql = "SELECT number FROM goodLog WHERE IGID=" + insertGoodLog.Rows[i]["IGID"] + " and goodID=" + goodID.Rows[0]["goodID"];
                    DataTable number = DBManage.getTable(conn, sql);
                    try
                    {
                        dataGridView1.Rows[i].Cells[j + 1].Value = number.Rows[0]["number"];
                    }
                    catch (IndexOutOfRangeException ex)
                    {
                        dataGridView1.Rows[i].Cells[j + 1].Value = 0;
                    }
                }
            }

            dataGridView1.Columns[1 + goodName.Rows.Count].Visible = false;
            dataGridView1.Columns[1 + goodName.Rows.Count].HeaderText = "IGID";
            for (int i = 0; i < insertGoodLog.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells[1 + goodName.Rows.Count].Value = insertGoodLog.Rows[i]["IGID"];
            }
        }

        private void teamComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable memberName = DBManage.getTable(conn, "SELECT name FROM member WHERE teamID=(SELECT teamID FROM team WHERE teamName='" + teamComboBox.Text + "')");
            memberComboBox.Items.Clear();
            for (int i = 0; i < memberName.Rows.Count; i++)
            {
                memberComboBox.Items.Add(memberName.Rows[i]["name"]);
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string number = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            number = Strings.StrConv(number, VbStrConv.Narrow, 0);
            DataTable dt1 = DataTypeTrans.dvTodt(dataGridView1);
            string IGID = dt1.Rows[e.RowIndex]["IGID"].ToString();
            string sql = "";

            if (e.ColumnIndex == 0)
            {
                string insertMonth = dt1.Rows[e.RowIndex]["月份"].ToString();
                sql = "UPDATE insertGoodLog SET insertMonth='" + insertMonth + "' WHERE IGID='" + IGID + "'";
            }
            else
            {
                string goodID = e.ColumnIndex.ToString();
                sql = "UPDATE goodLog SET number=" + number + " WHERE IGID=" + IGID + " AND goodID=" + goodID;
            }

            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("您確定要刪除嗎?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DataGridViewRow r1 = dataGridView1.CurrentRow;
                if (r1 != null)
                {
                    deleteDataGridView1Event(dataGridView1.CurrentCell.RowIndex.ToString());
                    dataGridView1.Rows.Remove(r1);

                }
            }
        }

        private void deleteDataGridView1Event(string selectIndex) //刪除進貨紀錄的事件
        {
            DataTable dt = DataTypeTrans.dvTodt(dataGridView1);
            conn.Open();
            string sql = "DELETE FROM insertGoodLog WHERE IGID='" + dt.Rows[Int32.Parse(selectIndex)]["IGID"] + "'";
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();

            sql = "DELETE FROM goodLog WHERE IGID='" + dt.Rows[Int32.Parse(selectIndex)]["IGID"] + "'";
            cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            button3.Visible = true;
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            deleteDataGridView1Event(e.Row.Index.ToString());
        }

        private void deleteContestButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("您確定要刪除嗎?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DataGridViewRow r1 = dataGridViewContest.CurrentRow;
                if (r1 != null)
                {
                    deleteDataGridViewEvent(dataGridViewContest.CurrentCell.RowIndex.ToString());
                    dataGridViewContest.Rows.Remove(r1);
                }
            }
        }

        private void deleteDataGridViewEvent(string selectIndex)
        {
            DataTable dt = DataTypeTrans.dvTodt(dataGridViewContest);
            switch(comboBoxTab2.Text)
            {
                case "全部":
                    string sql = "DELETE FROM insertLog WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];
                    string sql2 = "DELETE FROM logObject WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];

                    conn.Open();
                    SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    cmd = new SQLiteCommand(sql2, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    break;
                case "使用耗材":
                    sql = "UPDATE logObject SET number=0 WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];
                    conn.Open();
                    cmd = new SQLiteCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    break;
                case "受傷種類/處置":
                    sql = "UPDATE logHandle SET handleCount=0 WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];
                    sql2 = "DELETE FROM logCategory WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];
                    string sql3 = "DELETE FROM logKind WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];
                    conn.Open();
                    cmd = new SQLiteCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    cmd = new SQLiteCommand(sql2, conn);
                    cmd.ExecuteNonQuery();
                    cmd = new SQLiteCommand(sql3, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    break;
                case "受傷部位":
                    sql = "DELETE FROM logPos WHERE insertID=" + dt.Rows[Int32.Parse(selectIndex)]["insertID"];
                    conn.Open();
                    cmd = new SQLiteCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    break;
            }
                
            
        }

        private void dataGridViewContest_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataTable dt = DataTypeTrans.dvTodt(dataGridViewContest);
                string number = dataGridViewContest.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                number = Strings.StrConv(number, VbStrConv.Narrow, 0);
                string insertID = dt.Rows[e.RowIndex]["insertID"].ToString();
                string sql = "";
                if (number != "")
                {
                    if (e.ColumnIndex >= 4 && e.ColumnIndex < 4 + objectNameCount)
                    {
                        string objectID = (e.ColumnIndex - 3).ToString();
                        if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "V" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "UPDATE logObject SET number=1 WHERE insertID=" + insertID + " AND objectID=" + objectID;
                        }
                        else if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "X" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "UPDATE logObject SET number=0 WHERE insertID=" + insertID + " AND objectID=" + objectID;
                        }
                    }
                    else if (e.ColumnIndex == 2)
                    {
                        sql = "UPDATE insertLog SET contestName='" + number + "' WHERE insertID=" + insertID;
                    }
                    else if (e.ColumnIndex == 3)
                    {
                        sql = "UPDATE insertLog SET insertTime='" + number + "' WHERE insertID=" + insertID;
                    }
                    else if (e.ColumnIndex == 1)
                    {
                        sql = "UPDATE insertLog SET teamID='" + number + "' WHERE insertID=" + insertID;
                    }
                    else if (e.ColumnIndex == 0)
                    {
                        sql = "UPDATE insertLog SET memberID='" + number + "' WHERE insertID=" + insertID;
                    }
                    else if (e.ColumnIndex > (4 + objectNameCount) && e.ColumnIndex <= (4 + objectNameCount + handleCount))
                    {
                        string handleID = (e.ColumnIndex - (4 + objectNameCount)).ToString();
                        sql = "UPDATE logHandle SET handleCount='" + number + "' WHERE insertID=" + insertID + " AND handleID=" + handleID;
                    }
                    else if (e.ColumnIndex > (4 + objectNameCount + handleCount) && e.ColumnIndex <= (4 + objectNameCount + handleCount + kindCount))
                    {
                        if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "V" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "INSERT INTO logKind(kindID,insertID) VALUES (" + (e.ColumnIndex - 4 - objectNameCount - handleCount) + "," + insertID + ")";
                        }
                        else if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "X" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "DELETE FROM logKind WHERE kindID=" + (e.ColumnIndex - 4 - objectNameCount - handleCount) + " AND insertID=" + insertID;
                        }
                    }
                    else if (e.ColumnIndex > (4 + kindCount + objectNameCount + handleCount) && e.ColumnIndex <= (4 + objectNameCount + handleCount + kindCount + categoryCount))
                    {
                        if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "V" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "INSERT INTO logCategory(categoryID,insertID) VALUES (" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount) + "," + insertID + ")";
                        }
                        else if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "X" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "DELETE FROM logCategory WHERE categoryID=" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount) + " AND insertID=" + insertID;
                        }
                    }
                    else if (e.ColumnIndex > (4 + kindCount + objectNameCount + handleCount + categoryCount) && e.ColumnIndex <= (4 + objectNameCount + handleCount + kindCount + categoryCount + sideCount))
                    {
                        if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "V" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "INSERT INTO logSide(sideID,insertID) VALUES (" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount - categoryCount) + "," + insertID + ")";
                        }
                        else if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "X" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "DELETE FROM logSide WHERE sideID=" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount - categoryCount) + " AND insertID=" + insertID;
                        }
                    }
                    else if (e.ColumnIndex > (4 + kindCount + objectNameCount + handleCount + categoryCount + sideCount) && e.ColumnIndex <= (4 + objectNameCount + handleCount + kindCount + categoryCount + sideCount + posCount))
                    {
                        if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != "-" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            SQLiteConnection conn1 = new SQLiteConnection(@"Data Source=" + Constants.FHSDatabaseFile);
                            conn1.Open();
                            SQLiteCommand cmd1 = new SQLiteCommand("DELETE FROM logPos WHERE posID=" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount - categoryCount - sideCount) + " AND insertID=" + insertID, conn1);
                            cmd1.ExecuteNonQuery();
                            conn1.Close();

                            sql = "INSERT INTO logPos(posID,insertID,side) VALUES (" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount - categoryCount - sideCount) + "," + insertID + ",'" + number + "')";
                        }
                        else if (dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() == "-" && dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString() != temp)
                        {
                            sql = "DELETE FROM logPos WHERE posID=" + (e.ColumnIndex - 4 - objectNameCount - handleCount - kindCount - categoryCount - sideCount) + " AND insertID=" + insertID;
                        }
                    }

                    conn.Open();
                    SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }

            }
            catch (Exception ex)
            {

            }
        }

        private void dataGridViewContest_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            deleteDataGridViewEvent(e.Row.Index.ToString());
        }

        private void dataGridViewContest_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridViewContest.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.ImeMode = System.Windows.Forms.ImeMode.OnHalf;
        }

        private void tabC2Export_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "匯出成Excel檔";
            saveFileDialog1.Filter = "Excel活頁簿(*.xlsx)|*.xlsx|Excel 97-2003活頁簿(*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataTable temp = DataTypeTrans.dvTodt(dataGridViewContest);
                    DataSet ds = new DataSet();
                    ds.Tables.Add(temp);
                    ExcelHandle.ExportDataSetToExcel(ds, saveFileDialog1.FileName.ToString());
                    MessageBox.Show("匯出完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("匯出失敗，原因：" + ex);
                }
            }
        }

        private void tabC3Export_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "匯出成Excel檔";
            saveFileDialog1.Filter = "Excel活頁簿(*.xlsx)|*.xlsx|Excel 97-2003活頁簿(*.xls)|*.xls";
            var debug = saveFileDialog1.ShowDialog();
            if (debug == DialogResult.OK)
            {
                try
                {
                    DataTable temp = DataTypeTrans.dvTodt(dataGridViewStatistic);
                    DataSet ds = new DataSet();
                    ds.Tables.Add(temp);
                    ExcelHandle.ExportDataSetToExcel(ds, saveFileDialog1.FileName.ToString());
                    MessageBox.Show("匯出完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("匯出失敗，原因：" + ex);
                }
            }
        }

        private void tabC4Export_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "匯出成Excel檔";
            saveFileDialog1.Filter = "Excel活頁簿(*.xlsx)|*.xlsx|Excel 97-2003活頁簿(*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataTable temp = DataTypeTrans.dvTodt(dataGridView1);
                    DataSet ds = new DataSet();
                    ds.Tables.Add(temp);
                    ExcelHandle.ExportDataSetToExcel(ds, saveFileDialog1.FileName.ToString());
                    MessageBox.Show("匯出完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("匯出失敗，原因：" + ex);
                }
            }
        }

        private void backup_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "備份資料";
            saveFileDialog1.Filter = "檔案(*.dat)|*.dat";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    File.Copy(Constants.FHSDatabaseFile, saveFileDialog1.FileName.ToString(), true);
                    MessageBox.Show("備份完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("備份失敗，原因：" + ex);
                }
            }
        }

        private void dataGridViewContest_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DataTable dt = DataTypeTrans.dvTodt(dataGridViewContest);
            string insertID = dt.Rows[e.RowIndex]["insertID"].ToString();
            if (e.ColumnIndex > (4 + objectNameCount))
            {
                temp = dataGridViewContest[e.ColumnIndex, e.RowIndex].Value.ToString();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string value = "";
            if (Lib.InputBox("請輸入新部位", "新增部位:", ref value) == DialogResult.OK)
            {
                DBManage.createOrInsertCmd(conn, "INSERT INTO injuryPos(pos) VALUES ('" + value + "')");
                DataTable tmp;
                tmp = DBManage.getTable(conn, "SELECT pos FROM injuryPos ORDER BY posID");
                myPos(panelPos, tmp);
                button8.Text = "刪除部位";
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            if (button8.Text == "刪除部位")
            {
                for (int i = 0; i < deletePosArray.Length; i++)
                {
                    deletePosArray[i].Visible = true;
                    deletePosArray[i].Top = comboBoxPosArray[i].Bottom + 2;
                    button8.Text = "隱藏刪除部位";
                }
            }
            else if (button8.Text == "隱藏刪除部位")
            {
                for (int i = 0; i < deletePosArray.Length; i++)
                {
                    deletePosArray[i].Visible = false;
                    button8.Text = "刪除部位";
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (button9.Text == "刪除處置")
            {
                for (int i = 0; i < textBoxPosArray.Length; i++)
                {
                    negativeButtonArray[i].Width = 30;
                    positiveButtonArray[i].Width = 30;
                    positiveButtonArray[i].Left = negativeButtonArray[i].Right + 2;
                    deleteButtonArrayHandle[i].Left = positiveButtonArray[i].Right + 2;
                    deleteButtonArrayHandle[i].Visible = true;
                    button9.Text = "隱藏刪除處置";
                }
            }
            else if (button9.Text == "隱藏刪除處置")
            {
                for (int i = 0; i < textBoxPosArray.Length; i++)
                {
                    negativeButtonArray[i].Width = 50;
                    positiveButtonArray[i].Width = 50;
                    positiveButtonArray[i].Left = negativeButtonArray[i].Right + 2;
                    deleteButtonArrayHandle[i].Left = positiveButtonArray[i].Right + 2;
                    deleteButtonArrayHandle[i].Visible = false;
                    button9.Text = "刪除處置";
                }
            }
        }

        private void deleteUsedSupplies_Click(object sender, EventArgs e)
        {
            if (deleteUsedSupplies.Text == "刪除耗材")
            {
                for (int i = 0; i < deleteButtonArray.Length; i++)
                {
                    deleteButtonArray[i].Visible = true;
                    deleteUsedSupplies.Text = "隱藏刪除";
                }
            }
            else if (deleteUsedSupplies.Text == "隱藏刪除")
            {
                for (int i = 0; i < deleteButtonArray.Length; i++)
                {
                    deleteButtonArray[i].Visible = false;
                    deleteUsedSupplies.Text = "刪除耗材";
                }
            }

        }
        // TabC1 新增當日耗材
        private void addTodayUsed_Click(object sender, EventArgs e)
        {
            DBManage.createOrInsertCmd(connToday, "INSERT INTO todayTable(Today,t1,t2,t3,t4,t5,t6,t7) VALUES ('" + DateTime.Now.ToString("yyyy/MM/dd") + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + textBox9.Text + "','" + textBox10.Text + "')");
            MessageBox.Show("新增完成");
            textBox4.Text = "0";
            textBox5.Text = "0";
            textBox6.Text = "0";
            textBox7.Text = "0";
            textBox8.Text = "0";
            textBox9.Text = "0";
            textBox10.Text = "0";
        }
        // TabC1 匯出當日耗材至Excel
        private void exportTodayUsed_Click(object sender, EventArgs e)
        {
            DataTable datatableFHSToday = new DataTable();
            SQLiteCommand cmd = new SQLiteCommand("SELECT Today AS '日期',t1 AS '內膜',t2 AS '白貼0.5吋',t3 AS '白貼1吋',t4 AS '白貼1.5吋',t5 AS '輕彈2吋',t6 AS '強彈2吋',t7 AS '強彈3吋' FROM todayTable", connToday);
            SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);
            try
            {
                connToday.Open();
                da.Fill(datatableFHSToday);
                connToday.Close();
                da.Dispose();
                if (datatableFHSToday.Rows.Count > 0)
                {
                    saveFileDialog1.Title = "匯出成Excel檔";
                    saveFileDialog1.Filter = "Excel活頁簿(*.xlsx)|*.xlsx|Excel 97-2003活頁簿(*.xls)|*.xls";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string strFilePath = saveFileDialog1.FileName.ToString();
                        try
                        {
                            DataSet ds = new DataSet();
                            ds.Tables.Add(datatableFHSToday);
                            ExcelHandle.ExportDataSetToExcel(ds, strFilePath);
                            exportTodayUsed.Text = "匯出當日耗材用量Excel";
                            MessageBox.Show("匯出完成");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("匯出失敗，原因：" + ex);
                        }
                    }
                    else
                    {
                        MessageBox.Show("匯出已取消");
                    }
                }
                else
                {
                    MessageBox.Show("無資料，請先新增");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取本日耗材用量失敗，原因：" + ex);
            }
        }

    }
}
