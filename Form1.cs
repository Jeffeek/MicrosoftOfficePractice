using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;


namespace MicrosoftOfficePractice
{
    public partial class Form1 : Form
    {
        Word word;
        Excel excel;
        public Form1()
        {
            InitializeComponent();
            //excel = new Excel("HUI", 1);
            //word = new Word("1");
            //word.FindAndReplace("{раз}", "228");
            //word.FindAndReplace("{2раз}", "1337");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.GetTotalMemory(false);
            if (word != null || excel != null)
            {
                if (word != null)
                {
                    word.Close();
                    word = null;
                }
                if (excel != null)
                {
                    excel.CloseWorkBook();
                    excel = null;
                }
                GC.Collect();
                GC.GetTotalMemory(false);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var stream = File.Open($"{Directory.GetCurrentDirectory()}\\HUI.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    dataGridView1.DataSource = result.Tables[0];
                }
            }
        }
    }
}
