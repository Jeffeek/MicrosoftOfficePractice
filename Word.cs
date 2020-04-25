using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace MicrosoftOfficePractice
{
    public class Word
    {
        Application WordApp = new Application();
        Document doc;
        public string Name { get; private set; }
        public Word(string name)
        {
            Name = name;
            WordApp.Visible = false;
            if (File.Exists($"{Directory.GetCurrentDirectory()}\\{Name}.docx"))
            {
                doc = WordApp.Documents.Open(FileName: $"{Directory.GetCurrentDirectory()}\\{Name}.docx");
            }
            else
            {
                var result = System.Windows.Forms.MessageBox.Show("File not Found! Create?", "sudfjsdf", System.Windows.Forms.MessageBoxButtons.OKCancel);
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    doc = WordApp.Documents.Add();
                    doc.SaveAs2(FileName: $"{Directory.GetCurrentDirectory()}\\{Name}.docx");
                    WordApp.Documents.Open(FileName: $"{Directory.GetCurrentDirectory()}\\{Name}.docx");
                }
            }
        }

        //TODO: второй фильтр
        public void FindAndReplace(string ToFindText, string replaceWithText)
        {
            var range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: ToFindText, ReplaceWith: replaceWithText);
            Save();
        }

        public void Save()
        {
            doc.Save();
        }

        public void SaveAs(string name)
        {
            doc.SaveAs2(FileName: name);           
        }

        public void Close()
        {
            doc.Close();
            WordApp.Quit();
        }

        public void WriteText(string text)
        {
            Range range = doc.Range();
            range.Text = text;
            Save();
        }
    }
}
