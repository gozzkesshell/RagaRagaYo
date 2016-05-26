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
        //Microsoft.Office.Interop.Excel.Application xlexcel;
        //Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        //object misValue = System.Reflection.Missing.Value;

    
        

        public Form2()
        {
            InitializeComponent();
            this.label1.Text = "Hello, Visitor! Nice to meet you :)";
            this.label2.Text = "What's your name?";
            this.label3.Text = "What's your e-mail?";
            this.label4.Text = "Choose any day. P.S.Choose all ;)";
            this.label5.Text = "Finish?";
            this.label6.Text = "If you need a tent please check this item";
            this.button1.Enabled = true;

            


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
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
           string fileExcel = "C:/Users/Евгения/Source/Repos/RagaRagaYo2/WindowsFormsApplication1/bin/Debug/Visitors.xlsx";
           Excel.Application excelApp = new Excel.Application();
           Excel.Workbook excelBook = excelApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.ActiveSheet;
            excelApp.Visible = true;

            excelSheet.Cells[1, 1] = "Name";
            excelSheet.Cells[1, 2] = "E-mail";
            excelSheet.Cells[1, 3] = "Count of days";
            excelSheet.Cells[1, 4] = "Tent";

            int _lastRow = excelSheet.Range["A" + excelSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;

           int i = 0;
            foreach (int indexChecked in checkedListBox1.CheckedIndices){
                i++;
            }
           excelSheet.Cells[_lastRow, 1] = textBox1.Text;
           excelSheet.Cells[_lastRow, 2] = textBox2.Text;
           excelSheet.Cells[_lastRow, 3] = i;
           excelSheet.Cells[_lastRow, 4] = checkBox1.Text;

            excelBook.Save();
            this.Close();

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
