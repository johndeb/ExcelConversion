using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using Microsoft.CSharp.RuntimeBinder;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
namespace HospiceExcelTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filePath;
        private void button2_Click(object sender, EventArgs e)
        {
            ArrayList ValidMembers = new ArrayList();
            ArrayList InValidMembers = new ArrayList();
            int validMembers = 0;
            int invalidMembers = 0;
            bool openConn = false;
            Conncetion conn = new Conncetion(openConn);
            openConn = conn.openConncetion();
            if (openConn)
            {
                filePath = "";
                OpenFileDialog open = new OpenFileDialog();
                open.Filter = "Excel |*.xlsx; *.xls;";
                DialogResult result = open.ShowDialog();
                if (result == DialogResult.OK)
                {
                    filePath = open.FileName;
                }
                ArrayList members = firstColumn(filePath);
                for (int i = 0; i < members.Count; i++)
                {
                    Members mc = (Members)members[i];
                    if (mc.Name != String.Empty && mc.Surname != String.Empty && mc.Address != String.Empty && mc.Street != String.Empty && mc.Town != String.Empty && mc.PostCode != String.Empty && mc.Id != String.Empty && mc.Gender.ToString() != String.Empty && mc.LandLine.ToString() != String.Empty && mc.Mobile.ToString() != String.Empty && mc.InContact.ToString() != String.Empty && mc.MonthStarted.ToString() != String.Empty && mc.ReceiptNumber.ToString() != String.Empty)
                    {
                        conn.insert(mc.Name, mc.Surname, mc.Address, mc.Street, mc.Town, mc.PostCode, mc.Id, Convert.ToInt32(mc.Gender), mc.LandLine, mc.Mobile, mc.Email, Convert.ToInt32(mc.InContact), mc.MonthStarted, "Yes", mc.ReceiptNumber, "1");
                        validMembers++;
                        ValidMembers.Add(mc);
                    }
                    else
                    {
                        InValidMembers.Add(mc);
                        invalidMembers++;
                        
                    }
                }
                MessageBox.Show(invalidMembers.ToString());
                MessageBox.Show(validMembers.ToString());
                conn.closeConnection();
                button1.Visible = true;
                pictureBox1.Visible = true;
           }
 
        }

        public ArrayList firstColumn(string filename)
        {
            ArrayList members = new ArrayList();
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlsApp == null)
            {
                MessageBox.Show("Excel Could not be started");
                return null;
            }

            Workbook wb = xlsApp.Workbooks.Open(filename, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
            Sheets sheet = wb.Worksheets;
            Worksheet ws = (Worksheet)sheet.get_Item(1);
            Range range = ws.UsedRange;   
            int columnCount = ws.UsedRange.Columns.Count;
            int rowCount = ws.UsedRange.Rows.Count;

            for (int rows = 2; rows <= rowCount; rows++)
            {
                string member = "";
                for (int col = 1; col < columnCount; col++)
                {
                    try
                    {
                        if (col ==16)
                        {
                            string[] dateOnly = range.Cells[rows, col].Value.ToString().Split(' ');
                            member += dateOnly[0].ToString() + ":";
                        }
                        else
                        {
                            member += range.Cells[rows, col].Value2.ToString() + ":";
                        }
                    }
                    catch (RuntimeBinderException e)
                    {
                        member += ":";
                    }
                }
                string[] m = member.Split(':');

                Members newMember = new Members(Convert.ToInt32(m[0]), m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8], m[9], m[10], m[11], m[12], m[13], m[14],m[15], Convert.ToInt16(m[16]), m[17]);
                members.Add(newMember);
            }
            ws.Cells[19, 1] = "John";
            wb.SaveAs(@"C:\Users\johndebono10\Desktop\InvalidMembers.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            return members;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult res = saveFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                Document doc = new Document(iTextSharp.text.PageSize.A4, 10, 10, 42, 35);
                PdfWriter write = PdfWriter.GetInstance(doc, new FileStream(saveFileDialog1.FileName + ".pdf", FileMode.Create));
                doc.Open();
                Paragraph par = new Paragraph("Number  Name  Surname  ID  DateOfbirth Email Town Title Adress Street PostCode Gender");
                
                doc.Add(par);
                doc.Close();
            }
        }
    }
}
