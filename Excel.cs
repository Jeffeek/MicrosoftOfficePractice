using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace MicrosoftOfficePractice
{
    public class Excel
    {
        public string Path { get; private set; }
        _Application excel = new _Excel.Application();
        Workbook WorkBook;
        Worksheet WorkSheet;
        public Excel(string name, int Sheet)
        {
            Path = $"{Directory.GetCurrentDirectory()}\\{name}.xlsx";
            if (!File.Exists(Path))
            {
                var workbook = excel.Workbooks.Add();
                workbook.SaveAs(Path);
            }
            WorkBook = excel.Workbooks.Open(Path);
            WorkSheet = WorkBook.Worksheets[Sheet];
        }

        public string ReadCell(int i, int j)
        {
            i++; j++;

            if (WorkSheet.Cells[i, j] != null)
                return WorkSheet.Cells[i, j].Value2;
            else
                return "";
        }

        public void WriteToCell(int i, int j, string text)
        {
            i++; j++;
            WorkSheet.Cells[i, j].Value2 = text;
        }

        public void Save()
        {
            WorkBook.Save();
        }

        public void SaveAs(string name)
        {
            WorkBook.SaveAs($"{Directory.GetCurrentDirectory()}\\{name}.xlsx");
        }

        public void CloseWorkBook()
        {
            var ExcelProc = Process.GetProcessesByName("Excel");
            foreach (var process in ExcelProc)
            {
                process.Kill();
            }
        }

        public void CreateNewFile()
        {
            WorkBook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.WorkSheet = WorkBook.Worksheets[1];
        }

        public void CreateNewWorksheet()
        {
            Worksheet sheet = WorkBook.Worksheets.Add(After: WorkSheet);
        }

        public void SelectWorksheet(int numberOf)
        {
            this.WorkSheet = WorkBook.Worksheets[numberOf];
        }

        public void DeleteWorksheet(int numberOf)
        {
            WorkBook.Worksheets[numberOf].Delete();
        }

        public void ProtectSheet()
        {
            WorkSheet.Protect();
        }

        public void ProtectSheet(string Password)
        {
            WorkSheet.Protect(Password);
        }

        public void UnprotectSheet()
        {
            WorkSheet.Unprotect();
        }

        public void UnprotectSheet(string Password)
        {
            WorkSheet.Unprotect(Password);
        }

        public string[,] ReadRange(int startI, int startJ, int finishI, int finishJ)
        {
            Range range = (Range)WorkSheet.Range[WorkSheet.Cells[startI, startJ], WorkSheet.Cells[finishI, finishJ]];
            var temp = range.Value2;
            string[,] ret = new string[finishI - startI + 1, finishJ - startJ + 1];
<<<<<<< HEAD
            for (int i = 1; i <= finishI - startI; i++)
            {
                for (int j = 1; j <= finishJ - startJ; j++)
=======
            for (int i = 1; i <= finishI - startI + 1; i++)
            {
                for (int j = 1; j <= finishJ - startJ + 1; j++)
>>>>>>> 3c15c14f3cf824db8208d6004d2e0152db2a686f
                {
                    if (temp[i, j] == null)
                    {
                        ret[i - 1, j - 1] = "";
                    }
                    else
                    {
                        ret[i - 1, j - 1] = temp[i, j].ToString();
                    }
                }
            }
            return ret;
        }

        public void WriteRange(int startI, int startJ, int finishI, int finishJ, string[,] writeArr)
        {
            int rowsArr = 0;
            int ColumnArr = 0;
            for (int i = startI; i <= finishI; i++)
            {
                for (int j = startJ; j <= finishJ; j++)
                {
                    WorkSheet.Cells[i, j] = writeArr[rowsArr,ColumnArr];
                    ColumnArr++;
                }
                rowsArr++;
                ColumnArr = 0;
            }
            Save();
        }
    }
}
