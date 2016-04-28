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
            this.label6.Text = "If you need a tent please check this item";
            this.button1.Enabled = false;
        }

        private void textBox_TextChanged(object sender, EventArgs e)
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
            listBox1.Items.Clear();

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                if (checkedListBox1.CheckedItems.Contains(textBox1.Text) == false
                  && checkedListBox1.CheckedItems.Contains(textBox2.Text) == false)
                
                {
                    
                    listBox1.Items.Add(textBox1.Text);
                    listBox1.Items.Add(textBox2.Text);
                }

                textBox1.Text = "";
                textBox2.Text = "";
                
            }


            foreach(object item in checkedListBox1.Items)
            {
                if (checkedListBox1.CheckedItems.Contains(item))
                    listBox1.Items.Add(item + "TRUE");
                if (!checkedListBox1.CheckedItems.Contains(item))
                 {
                    listBox1.Items.Add(item + "FALSE");
                 }
            }


            if (checkBox1.Checked == true)
                listBox1.Items.Add("TentTRUE");
            else
                listBox1.Items.Add("TentFALSE");

            string filename = "Visitors.txt";
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
