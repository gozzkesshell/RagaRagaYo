using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.label1.Text = "Hello, Visitor! Nice to meet you :)";
            this.label2.Text = "What's your name?";
            this.label3.Text = "What's your e-mail?";
            this.label4.Text = "Choose any day. P.S.Choose all ;)";
            this.label5.Text = "Finish?";
            this.listBox1.Location = new System.Drawing.Point(408, 64);
            this.listBox1.Size = new System.Drawing.Size(128, 186);
            this.listBox1.TabIndex = 3;
            this.button1.Enabled = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "")
            {
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }

           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {
           
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" || textBox2.Text != "")
            {
                if (checkedListBox1.CheckedItems.Contains(textBox1.Text) == false
                   || checkedListBox1.CheckedItems.Contains(textBox2.Text) == false)
                {
                    checkedListBox1.Items.Add(textBox1.Text, CheckState.Checked);
                    checkedListBox1.Items.Add(textBox2.Text, CheckState.Checked);
                }
                    textBox1.Text = "";
                    textBox2.Text = "";
            }

            listBox1.Items.Clear();

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                listBox1.Items.Add(checkedListBox1.CheckedItems[i]);
            }

            string filename = "data.txt";
            string listboxData = "";
            foreach (string str in listBox1.Items)
            {
                listboxData += str + " ";
            }
            listboxData += Environment.NewLine;
            File.AppendAllText(filename, listboxData);

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
