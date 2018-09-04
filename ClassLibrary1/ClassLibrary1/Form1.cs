using System;
using System.Windows.Forms;

namespace ClassLibrary1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public Form1(bool param1, int param2, double param3, string param4)
        {
            InitializeComponent();
            button1.Visible = false; //If you open a form with ShowDialog the communication to VB6 isn't working
            MessageBox.Show("Parameters:" + Environment.NewLine + param1 + Environment.NewLine + param2 + Environment.NewLine + param3 + Environment.NewLine + param4, "C# Form", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Class1.OnMessageEvent("OpenForm1");
        }
    }
}
