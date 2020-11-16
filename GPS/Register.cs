using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Register : Form
    {
        public Register()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string auth_level = null;
            switch (textBox4.Text)
            {
                case "GFS1":
                    //DO something
                    auth_level = "L1";
                    connectDB(auth_level);
                    break;
                case "GFS2":
                    //DO something
                    auth_level = "L2";
                    connectDB(auth_level);
                    break;
                case "GFS3":
                    //DO something
                    auth_level = "L3";
                    connectDB(auth_level);
                    break;
                case "":
                    MessageBox.Show("權限密碼欄位，不可以為空 \r\n 相關資訊請通知俊億");
                    break;
            }
        }

        public void connectDB(string auth_level)
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = "insert into Employee (employee_name, employee_account, employee_password, employee_authlevel) values('" + textBox3.Text.Trim() + "','" + textBox1.Text.Trim() + "','" + textBox2.Text.Trim() + "','" + auth_level + "')";
                    //textBox5.Text = sqlcmd;
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("資料建立成功！ \r\n 請關閉註冊視窗，嘗試登入");
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("建立資料時，出現資料庫連線錯誤，請通知俊億 \r\n" +ex.Message,ToString());
            }
        }


    }
}
