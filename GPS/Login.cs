using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace GPS
{
    public partial class Login : Form
    {
        MainPanel mainpanel1 = null;
        public static string username = null; 
        public Login(MainPanel fr1)
        {
            InitializeComponent();
            mainpanel1 = fr1;
            
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Register frm1 = new Register();
            frm1.Show();
        }
        private void ConnectAccess()
        {
            //OleDbConnection cn = new OleDbConnection(@"Provider=Microsoft.Ace.OLEDB.12.0;Data Source=C:\\Users\\YGE\\Desktop\\MIS\\MIS\\bin\\Debug\\GFS.accdb");
            var DBPath = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source="+ Application.StartupPath + "\\GFS.accdb";
            using (OleDbConnection conn = new OleDbConnection(DBPath))
            {
                try
                {
                    conn.Open();
                    string sqlcmd = "select * from Employee where employee_account = '" + textBox1.Text + "'";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    OleDbDataReader dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        string confirm = dr.GetString(3);
                        username = dr.GetString(1);
                        if (confirm == textBox2.Text)
                        {
                            MessageBox.Show("Hi !  " + username);
                            mainpanel1.Controls["label2"].Text = username;
                            mainpanel1.Show();
                            this.Hide();
                        }
                        else
                        {
                            MessageBox.Show("密碼輸入錯誤");
                        }
                    }
                    else
                    {
                        MessageBox.Show("帳號密碼不存在，請重新輸入");
                        textBox1.Text = "";
                        textBox2.Text = "";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("資料庫連線錯誤 \r\n" + ex.Message.ToString());
                }
                finally
                {

                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ConnectAccess();
        }
        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            mainpanel1.Close();
        }
    }
}
 