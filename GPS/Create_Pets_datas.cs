using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Create_Pets_datas : Form
    {
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
                    string sqlcmd = "insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_colors, pets_lotion, pets_skinCoditions, pets_style, pets_comments, pets_parent_index, pets_createdate, pets_updatedate) values ('" + textBox8.Text + "','" + comboBox1.SelectedItem.ToString() + "'," + pets_sex + ",'" + textBox1.Text + "','" + textBox2.Text + "',"+ pets_sex + ",'"+ pets_sstatus + "','"+ precaution + "','"+ allergy_txt + "','"+ notouch_txt + "','"+ exp_txt + "')";
                }
                // 這邊先insert後取得寵物的編號，我們select編號出來進行下面的

                //insert Pets_info
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = "insert into Pets_info (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_colors, pets_lotion, pets_skinCoditions, pets_style, pets_comments, pets_parent_index, pets_createdate, pets_updatedate) values ('" + textBox8.Text + "','" + comboBox1.SelectedItem.ToString() + "'," + pets_sex + ",'" + textBox1.Text + "','" + textBox2.Text + "'," + pets_sex + ",'" + pets_sstatus + "','" + precaution + "','" + allergy_txt + "','" + notouch_txt + "','" + exp_txt + "')";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
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
