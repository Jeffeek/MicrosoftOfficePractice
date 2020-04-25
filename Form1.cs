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
using Word = Microsoft.Office.Interop.Word;

namespace MicrosoftOfficePractice
{
    public partial class Form1 : Form
    {
        Word word;
        Excel excel;
        public Form1()
        {
            InitializeComponent();
            excel = new Excel("HUI", 1);
            word = new Word("1");
            word.FindAndReplace("{раз}", "228");
            word.FindAndReplace("{2раз}", "1337");
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
            string[,] send = new string[dataGridView1.RowCount-1, dataGridView1.ColumnCount];
            for (int i = 0; i < dataGridView1.RowCount-1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    send[i, j] = Convert.ToString(dataGridView1[j, i].Value);
                }
            }
            excel.WriteRange(1, 1, dataGridView1.RowCount-1, dataGridView1.ColumnCount, send);
        }
    }
}
