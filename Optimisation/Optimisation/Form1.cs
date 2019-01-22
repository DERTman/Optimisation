using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Optimisation
{
    public partial class Form1 : Form
    {
        
        string fileName = System.IO.Path.GetFullPath(@"Дешевая химия.xlsx");

        public Form1()
        {
            
            InitializeComponent();
            label3.Text = "0";
            label10.Text = "0";
            label11.Text = "0";
            label12.Text = "0";
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            //Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле
            //label3.Text = xlSht.Cells[11, "C"].Value.ToString();

            //xlApp.Quit();

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            //xlApp = null;
            //xlWb = null;
            //xlSht = null;

            //System.GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле

            xlSht.Cells[9, "C"] = textBox1.Text;
            xlSht.Cells[9, "D"] = textBox2.Text;
            xlSht.Cells[9, "E"] = textBox3.Text;
            label3.Text = xlSht.Cells[11, "C"].Value.ToString();
            label10.Text = xlSht.Cells[14, "C"].Value.ToString();
            label11.Text = xlSht.Cells[15, "C"].Value.ToString();
            label12.Text = xlSht.Cells[16, "C"].Value.ToString();

            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            xlApp = null;
            xlWb = null;
            xlSht = null;


        }
    }
}
