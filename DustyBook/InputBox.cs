using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace xml_treeview
{
    public partial class InputBox : Form
    {
        public string inputValue;

        public InputBox()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            inputValue = textBox1.Text;
            this.Hide();
        }
    }
}
