using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Form2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFile();
        }

        public void OpenFile()
        {
            ExcelClass excel = new ExcelClass(@"C:\Users\varas\OneDrive\Documents\2019-2020NBAPlayerStats.xlsx", 1);

            MessageBox.Show(excel.ReadCell(2, 2));
        }
    }
}
