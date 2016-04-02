using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutodeskAccess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var link = new deskTool();
            link.selectDB();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var link = new deskTool();
            //link.selectScript();
            if(link.selectScript() !=null)
            {
                link.writeContents();
            }
        }
    }
}
