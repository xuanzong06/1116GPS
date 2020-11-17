/*update*/
//2020.11.16 Login
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
                    //string sqlcmd = "select * from Employee where employee_account = @account"

                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.Parameters.Add("@account", OleDbType.VarChar);
                    cmd.Parameters["@account"].Value = textBox1.Text;


                    OleDbDataReader dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        string confirm = dr.GetString(3);
                        username = dr.GetString(1);
                        if (confirm == textBox2.Text.Trim())
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

//2020.11.16 Create_Parents_datas
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
                    //string sqlcmd = "insert into Parent(parent_name,parent_gender,parent_telephone,parent_cellphone,parent_id,parent_address,parent_ec,parent_ectelephone,parent_eccellphone,parent_createdate) values (@parent_name,@parent_gender,@parent_telephone,@parent_cellphone,@parent_id,@parent_address,@parent_ec,@parent_ectelephone,@parent_eccellphone,@parent_createdate)";
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.Parameters.Add("@parent_name", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_gender", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_telephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_cellphone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_id", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_address", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_ec", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_ectelephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_eccellphone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_createdate", OleDbType.VarChar);

                    cmd.Parameters["@parent_name"].Value = "";
                    cmd.Parameters["@parent_gender"].Value = "";
                    cmd.Parameters["@parent_telephone"].Value = "";
                    cmd.Parameters["@parent_cellphone"].Value = "";
                    cmd.Parameters["@parent_id"].Value = "";
                    cmd.Parameters["@parent_address"].Value = "";
                    cmd.Parameters["@parent_ec"].Value = "";
                    cmd.Parameters["@parent_ectelephone"].Value = "";
                    cmd.Parameters["@parent_eccellphone"].Value = "";
                    cmd.Parameters["@parent_createdate"].Value = "";
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

//2020.11.16 Create_Pets_datas
using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Create_Pets_datas : Form
    {
        //儲存寵物編號
        string pets_index = null;

        public Create_Pets_datas()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        private void create_pets_data(int pets_sex, string pets_sstatus, string precaution, string allergy_txt, string notouch_txt, string exp_txt)
        {            
            try
            {
                // insert Pets
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_colors, pets_lotion, pets_skinCoditions, pets_style, pets_comments, pets_parent_index, pets_createdate, pets_updatedate) values ('" + textBox8.Text + "','" + comboBox1.SelectedItem.ToString() + "'," + pets_sex + ",'" + textBox1.Text + "','" + textBox2.Text + "',"+ pets_sex + ",'"+ pets_sstatus + "','"+ precaution + "','"+ allergy_txt + "','"+ notouch_txt + "','"+ exp_txt + "')";
                    //改寫 sql injection prevent
                    //insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_colors, pets_lotion, pets_skinCoditions, pets_style, pets_comments, pets_parent_index, pets_createdate, pets_updatedate) values (@pets_names, @pets_breed, @pets_sex, @pets_birthDay, @pets_weight, @pets_ligation, @pets_sstatus, @pets_precautions, @pets_allergy, @pets_notouch, @pets_experiences, @pets_colors, @pets_lotion, @pets_skinCoditions, @pets_style, @pets_comments, @pets_parent_index, @pets_createdate, @pets_updatedate) 
                     
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    //改寫 sql injection prevent
                    cmd.Parameters.Add("@pets_names", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_breed", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sex", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_birthDay", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_weight", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_ligation", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sstatus", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_precautions", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_allergy", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_notouch", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_experiences", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_colors", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_lotion", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_skinCoditions", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_style", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_comments", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_parent_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_createdate", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_updatedate", OleDbType.VarChar);

                    cmd.Parameters["@pets_names"].Value = "";
                    cmd.Parameters["@pets_breed"].Value = "";
                    cmd.Parameters["@pets_sex"].Value = "";
                    cmd.Parameters["@pets_birthDay"].Value = "";
                    cmd.Parameters["@pets_weight"].Value = "";
                    cmd.Parameters["@pets_ligation"].Value = "";
                    cmd.Parameters["@pets_sstatus"].Value = "";
                    cmd.Parameters["@pets_precautions"].Value = "";
                    cmd.Parameters["@pets_allergy"].Value = "";
                    cmd.Parameters["@pets_notouch"].Value = "";
                    cmd.Parameters["@pets_experiences"].Value = "";
                    cmd.Parameters["@pets_colors"].Value = "";
                    cmd.Parameters["@pets_lotion"].Value = "";
                    cmd.Parameters["@pets_skinCoditions"].Value = "";
                    cmd.Parameters["@pets_style"].Value = "";
                    cmd.Parameters["@pets_comments"].Value = "";
                    cmd.Parameters["@pets_parent_index"].Value = "";
                    cmd.Parameters["@pets_createdate"].Value = "";
                    cmd.Parameters["@pets_updatedate"].Value = "";


                    //string sqlcmd = "insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences) values ('" + textBox8.Text + "','" + comboBox1.SelectedItem.ToString() + "'," + pets_sex + ",'" + textBox1.Text + "','" + textBox2.Text + "'," + pets_sex + ",'" + pets_sstatus + "','" + precaution + "','" + allergy_txt + "','" + notouch_txt + "','" + exp_txt + "')";
                    
                    //cmd.Parameters.Add("@pets_names", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_breed", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_sex", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_birthDay", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_weight", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_ligation", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_sstatus", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_precautions", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_allergy", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_notouch", OleDbType.VarChar);
                    //cmd.Parameters.Add("@pets_experiences", OleDbType.VarChar);
                    
                    //cmd.Parameters["@pets_names"].Value = textBox8.Text;
                    //cmd.Parameters["@pets_breed"].Value = comboBox1.SelectedItem.ToString();
                    //cmd.Parameters["@pets_sex"].Value = pets_sex;
                    //cmd.Parameters["@pets_birthDay"].Value = textBox1.Text;
                    //cmd.Parameters["@pets_weight"].Value = textBox2.Text;
                    //cmd.Parameters["@pets_ligation"].Value = pets_sex;
                    //cmd.Parameters["@pets_sstatus"].Value = pets_sstatus;
                    //cmd.Parameters["@pets_precautions"].Value = precaution;
                    //cmd.Parameters["@pets_allergy"].Value = allergy_txt;
                    //cmd.Parameters["@pets_notouch"].Value = notouch_txt;
                    //cmd.Parameters["@pets_experiences"].Value = exp_txt;

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增資料到寵物基本資料時發生錯誤："+ex.Message.ToString());
            }

            //Select 編號
            try
            {
                // insert Pets
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "select ____ from Pets where ";　// SQL 還沒撰寫完畢
                    //我在這邊想到的是，我要怎麼確認我抓到的編號就是最新的？
                    //1.因為這個APP只會同一時間只有一人操作，所以我應該抓最新的寵物資料？
                    //2.我建立一個時間欄位，然後宣告一個datetime /string去儲存insert時的資料時間，用這個時間當作key去找？
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    OleDbDataReader dr = cmd.ExecuteReader();
                    while(dr.Read()){
                        pets_index = dr.GetInt32(0).ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }

            try
            {
                //insert Pets_info
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增寵物健康狀況時發生錯誤：" + ex.Message.ToString());
            }
        }

        private void status_check() 
        {
            int pets_sex = 2;
            string pets_sstatus = null;
            string precaution = null;
            string allergy_txt = null;
            string notouch_txt = null;
            string exp_txt = null;
            
            try
            {
                if (!radioButton3.Checked && !radioButton4.Checked)
                {
                    MessageBox.Show("請點選性別");
                }

                if (radioButton4.Checked) //radioButton4 = 公
                {
                    pets_sex = 1;
                }

                if (radioButton3.Checked)//radioButton3 = 母
                {
                    pets_sex = 0;
                }
                
                if (checkBox2.Checked)
                {
                    pets_sstatus = "良好";
                }
                else
                {
                    pets_sstatus = textBox3.Text;
                }

                if (checkBox3.Checked)
                {
                    precaution = "無";
                }
                else
                {
                    precaution = textBox4.Text;
                }

                if (checkBox4.Checked)
                {
                    allergy_txt = "無";
                }
                else if (checkBox5.Checked)
                {
                    allergy_txt += "電剪";
                }
                else if (checkBox6.Checked)
                {
                    allergy_txt += "洗毛精";
                }
                else if (checkBox7.Checked)
                {
                    allergy_txt += "其他";
                    allergy_txt += textBox5.Text;
                }

                if (checkBox11.Checked)
                {
                    notouch_txt = "無";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "腳";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "耳朵";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "嘴";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "其他";
                    notouch_txt += textBox6.Text;
                }

                if (checkBox17.Checked)
                {
                    exp_txt = "無";
                }
                else
                {
                    exp_txt = textBox7.Text;
                }

                create_pets_data(pets_sex, pets_sstatus, precaution, allergy_txt, notouch_txt, exp_txt);
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立資料時，出現錯誤，請通知俊億 \r\n" + ex.Message, ToString());
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            /*使用者輸入 20201010，離開欄位後自動轉換 2020/10/10 */
            DateTime dt = DateTime.ParseExact(textBox1.Text, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            var result = dt.ToString("yyyy/MM/dd");
            textBox1.Text = result;
        }
    }
}

//2020.11.16 Create_PetsSalon_datas
//介面

//增加一個Button


2020.11.17

//Create_Parents_datas
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
                DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = "insert into Parent(parent_name,parent_gender,parent_telephone,parent_cellphone,parent_id,parent_address,parent_ec,parent_ectelephone,parent_eccellphone,parent_createdate) values (@parent_name,@parent_gender,@parent_telephone,@parent_cellphone,@parent_id,@parent_address,@parent_ec,@parent_ectelephone,@parent_eccellphone,@parent_createdate)";
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.Parameters.Add("@parent_name", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_gender", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_telephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_cellphone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_id", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_address", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_ec", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_ectelephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_eccellphone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_createdate", OleDbType.VarChar);
                    
                    cmd.Parameters["@parent_name"].Value = textBox1.Text;
                    cmd.Parameters["@parent_gender"].Value = pgender;
                    cmd.Parameters["@parent_telephone"].Value = textBox2.Text;
                    cmd.Parameters["@parent_cellphone"].Value = textBox3.Text;
                    cmd.Parameters["@parent_id"].Value = textBox4.Text;
                    cmd.Parameters["@parent_address"].Value = textBox7.Text;
                    cmd.Parameters["@parent_ec"].Value = textBox5.Text;
                    cmd.Parameters["@parent_ectelephone"].Value = textBox6.Text;
                    cmd.Parameters["@parent_eccellphone"].Value = textBox8.Text;
                    cmd.Parameters["@parent_createdate"].Value = date1;

                    cmd.ExecuteNonQuery();

                    if(MessageBox.Show("家長資料建立成功！ \r\n繼續建立寵物請按「是」，要離開請按「否」", "訊息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes){

                        //select for Index
                        int parent_index = Get_Parent_Index();

                        Create_Pets_datas createdatas2 = new Create_Pets_datas(parent_index);
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

        private int Get_Parent_Index()
        {
            int parent_index = 0;
            using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
            {
                string sqlcmd = "select parent_index from Parent where parent_name = '" + textBox1.Text + "' and parent_cellphone = '" + textBox3.Text + "'";
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                OleDbDataReader dr = cmd.ExecuteReader();
                while(dr.Read()){
                    parent_index = dr.GetInt32(0);
                }
                label9.Text = "編號(成功建檔後產生)：" + parent_index.ToString(); //顯示在畫面上，可能要先取消 this.close()用來看有沒有正常顯示
            }
            return parent_index;
        }
    }
}

//Create_Pets_datas
using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Create_Pets_datas : Form
    {
        //儲存寵物編號
        string pets_index = null;
        //Get家長編號
        int parent_index = 0;

        public Create_Pets_datas(int p_index)
        {
            InitializeComponent();
            parent_index = p_index;
        }

        private void button1_Click(object sender, EventArgs e)
        {            

        }

        private void create_pets_data(int pets_sex, string pets_sstatus, string precaution, string allergy_txt, string notouch_txt, string exp_txt)
        {            
            try
            {
                // insert Pets
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    //改寫 sql injection prevent
                    string sqlcmd = "insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_parent_index, pets_createdate ) values ( @pets_names, @pets_breed, @pets_sex, @pets_birthDay, @pets_weight, @pets_ligation, @pets_sstatus, @pets_precautions, @pets_allergy, @pets_notouch, @pets_experiences, @pets_parent_index, @pets_createdate )";
                     
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    //改寫 sql injection prevent

                    cmd.Parameters.Add("@pets_names", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_breed", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sex", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_birthDay", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_weight", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_ligation", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sstatus", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_precautions", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_allergy", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_notouch", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_experiences", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_parent_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_createdate", OleDbType.Date);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@pets_names"].Value = textBox8.Text;
                    cmd.Parameters["@pets_breed"].Value = comboBox1.SelectedItem.ToString();
                    cmd.Parameters["@pets_sex"].Value = pets_sex;
                    cmd.Parameters["@pets_birthDay"].Value = textBox1.Text;
                    cmd.Parameters["@pets_weight"].Value = textBox2.Text;
                    cmd.Parameters["@pets_ligation"].Value = pets_sex;
                    cmd.Parameters["@pets_sstatus"].Value = pets_sstatus;
                    cmd.Parameters["@pets_precautions"].Value = precaution;
                    cmd.Parameters["@pets_allergy"].Value = allergy_txt;
                    cmd.Parameters["@pets_notouch"].Value = notouch_txt;
                    cmd.Parameters["@pets_experiences"].Value = exp_txt;
                    cmd.Parameters["@pets_parent_index"].Value = "";
                    cmd.Parameters["@pets_createdate"].Value = date1;

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增資料到寵物基本資料時發生錯誤："+ex.Message.ToString());
            }

            //Select 編號
            try
            {
                int p_i = Get_pet_index();
                pets_index = Get_pet_index().ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }

            try
            {
                //insert Pets_info
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd =" insert into Pets_info (petsinfo_hospital, petsinfo_doctors, petsinfo_telephone, petsinfo_address, petsinfo_medicalHistory, petsinfo_medicalTreatement, petsinfo_vacciendate, petsinfo_rabiesvdate, petsinfo_pets_index, petsinfo_createdate) values ('" + textBox10.Text + "','" + textBox11.Text + "','" + textBox12.Text + "','" + textBox13.Text + "','" + textBox14.Text + "',true or false'" + textBox15.Text + "','" + textBox16.Text + "')";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增寵物健康狀況時發生錯誤：" + ex.Message.ToString());
            }
        }

        private void status_check() 
        {
            int pets_sex = 2;
            string pets_sstatus = null;
            string precaution = null;
            string allergy_txt = null;
            string notouch_txt = null;
            string exp_txt = null;
            
            try
            {
                if (!radioButton3.Checked && !radioButton4.Checked)
                {
                    MessageBox.Show("請點選性別");
                }

                if (radioButton4.Checked) //radioButton4 = 公
                {
                    pets_sex = 1;
                }

                if (radioButton3.Checked)//radioButton3 = 母
                {
                    pets_sex = 0;
                }
                
                if (checkBox2.Checked)
                {
                    pets_sstatus = "良好";
                }
                else
                {
                    pets_sstatus = textBox3.Text;
                }

                if (checkBox3.Checked)
                {
                    precaution = "無";
                }
                else
                {
                    precaution = textBox4.Text;
                }

                if (checkBox4.Checked)
                {
                    allergy_txt = "無";
                }
                else if (checkBox5.Checked)
                {
                    allergy_txt += "電剪";
                }
                else if (checkBox6.Checked)
                {
                    allergy_txt += "洗毛精";
                }
                else if (checkBox7.Checked)
                {
                    allergy_txt += "其他";
                    allergy_txt += textBox5.Text;
                }

                if (checkBox11.Checked)
                {
                    notouch_txt = "無";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "腳";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "耳朵";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "嘴";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "其他";
                    notouch_txt += textBox6.Text;
                }

                if (checkBox17.Checked)
                {
                    exp_txt = "無";
                }
                else
                {
                    exp_txt = textBox7.Text;
                }

                create_pets_data(pets_sex, pets_sstatus, precaution, allergy_txt, notouch_txt, exp_txt);
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立資料時，出現錯誤，請通知俊億 \r\n" + ex.Message, ToString());
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            /*使用者輸入 20201010，離開欄位後自動轉換 2020/10/10 */
            DateTime dt = DateTime.ParseExact(textBox1.Text, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            var result = dt.ToString("yyyy/MM/dd");
            textBox1.Text = result;
        }

        private int Get_pet_index()
        {
            int _pets_index = 0;

            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "select pets_index from Pets where pets_names ='" + textBox8 + "' and pets_parent_index = '" + parent_index + "'";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    OleDbDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        _pets_index = dr.GetInt32(0);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }

            return _pets_index;
        }
    }
}
