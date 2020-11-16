using System;
/**/
using System.Data.OleDb;
using System.Windows.Forms;

namespace GPS
{
    public partial class Create_Parents_datas : Form
    {
        public Create_Parents_datas()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!radioButton1.Checked && !radioButton2.Checked)
                {
                    MessageBox.Show("請點選性別");
                }

                if (radioButton1.Checked)
                {
                    int pgender = 1;
                    create_parentsData(pgender);
                }

                if (radioButton2.Checked)
                {
                    int pgender = 0;
                    create_parentsData(pgender);
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立資料時，出現資料庫連線錯誤，請通知俊億 \r\n" + ex.Message, ToString());
            }
        }

        private void create_parentsData(int pgender)
        {
            try
            {
                DateTime date1 = DateTime.Now;
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = "insert into Parent(parent_name,parent_gender,parent_telephone,parent_cellphone,parent_id,parent_address,parent_ec,parent_ectelephone,parent_eccellphone,parent_createdate) values('" + textBox1.Text + "'," + pgender + ",'" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox7.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox8.Text + "','" + date1 + "')";
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.ExecuteNonQuery();
                    if(MessageBox.Show("家長資料建立成功！ \r\n繼續建立寵物請按「是」，要離開請按「否」", "訊息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes){
                        Create_Pets_datas createdatas2 = new Create_Pets_datas();
                        createdatas2.Show();
                        this.Close();
                    }
                    else
                    {
                        this.Close();
                    }
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
    }
}
