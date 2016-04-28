using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            this.label1.Text = "Hello Band! Nice to meet you :)";
            this.label2.Text = "What's your name?";
            this.label3.Text = "Please add link on your social network page";
            this.label4.Text = "How many songs do you want to play?";
            this.label5.Text = "Finish?";
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                
                listBox1.Items.Add(textBox1.Text);
                listBox1.Items.Add(textBox2.Text);
           
                textBox1.Text = "";
                textBox2.Text = "";

            }

            listBox1.Items.Add(numericUpDown1.Value.ToString());

            string filename = "Bands.txt";
            string listboxData = "";
            foreach (string str in listBox1.Items)
            {
                listboxData += str + " ";
            }
            listboxData += Environment.NewLine;
            File.AppendAllText(filename, listboxData);

            this.Close();   

        }
    }
}
