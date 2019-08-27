using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Windows.Forms.DataVisualization.Charting;
using offis = Microsoft.Office.Interop.Excel;

namespace Distance_Sensor
{
    public partial class Form1 : Form
    {
        string mesafe = "0";
        string deneme = "A";
        int t = 1;
        int k = 0;
        DateTime yeni = DateTime.Now;
        int satir = 1;
        int sutun = 1;
        int zaman = 0;
        int satirNo = 1;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            serialPort1.Close();
            this.chart1.Titles.Add("Mesafe Ölçüm");
            DateTime yeni = DateTime.Now; 
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            serialPort1.PortName = "COM9";
            serialPort1.Open();
            timer1.Enabled = true;
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            mesafe = serialPort1.ReadLine();
            this.chart1.Series["Veri"].Points.AddXY(zaman, mesafe);
            zaman = (zaman + 1);

            satir = dataGridView1.Rows.Add();

            dataGridView1.Rows[satir].Cells[0].Value = satirNo;
            dataGridView1.Rows[satir].Cells[1].Value = mesafe;
            dataGridView1.Rows[satir].Cells[2].Value = yeni.ToLongTimeString();
            dataGridView1.Rows[satir].Cells[3].Value = yeni.ToShortDateString();

            satir++;
            satirNo++;

            label1.Text = mesafe;
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            serialPort1.Close();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            objExcel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook objbook = objExcel.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);

            for (int s = 0; s < dataGridView1.Columns.Count; s++)
            {
                Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, s + 1];
                myrange.Value2 = dataGridView1.Columns[s].HeaderText;
            }

            for (int s = 0; s < dataGridView1.Columns.Count; s++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[j + 2, s + 1];
                    myrange.Value2 = dataGridView1[s, j].Value;
                }


            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }
    }
}
