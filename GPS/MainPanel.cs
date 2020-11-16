using System;
using System.Windows.Forms;

namespace GPS
{
    public partial class MainPanel : Form
    {
        public string username = null;
        
        public MainPanel()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Login logfrm = new Login(this);
            logfrm.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Create_Parents_datas createdatasfrm = new Create_Parents_datas();
            createdatasfrm.Show();
            //this.Hide();
        }
    }
}
