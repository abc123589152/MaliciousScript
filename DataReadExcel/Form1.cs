
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Windows.Forms;
namespace DataReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)//Export FortigateAddressScript
        {
            String filePath = @"C:\temp\"+textBox2.Text+".xlsx";//Select ReadExcel filePath
            string CellText = "";//存放每個輸入值
            string ExportTxtFilePath = textBox5.Text;
            int ExcelSheetNum = int.Parse(textBox4.Text);
            int yist_num = int.Parse(textBox1.Text);
            int sheet_Row_num = int.Parse(textBox3.Text);//first sheet Row number
            int sheet_Row_num_1 = int.Parse(textBox7.Text);//last sheet Row number
            int j = 1;
            int firstCellNum;
            //Pass the filepath and filename to the StreamWriter Constructor
            StreamWriter sw = new StreamWriter("C:\\temp\\"+ExportTxtFilePath+".txt");//Export txt FilePath
            try
            {
                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(ExcelSheetNum-1);
                int rowCount;
                if (sheet != null)
                {
                    firstCellNum = sheet_Row_num - 1;//以十進位來看第一筆數字要減1
                    IRow Currowfirstone = sheet.GetRow(firstCellNum);
                    if (sheet_Row_num_1 != 0)
                    {
                        rowCount = sheet_Row_num_1;
                    }
                    else 
                    {             
                        rowCount = sheet.LastRowNum+1;
                    }
                    sw.WriteLine("config firewall address");//Write a first line of text
                    sw.WriteLine("   edit yicst_0" + yist_num.ToString());
                    sw.WriteLine("    set subnet " + Currowfirstone.GetCell(0).ToString()+ "/32");
                    sw.WriteLine("next");
                    for (int i=firstCellNum+1; i<rowCount; i++)//此處加1為第二個開始
                    {
                        IRow curROw = sheet.GetRow(i);//Read Sheet Row       
                        CellText = curROw.GetCell(0).ToString();
                        //Write a line of text
                        
                            sw.WriteLine("config firewall address");//Write a second line of text
                            sw.WriteLine("   edit yicst_0" + (yist_num + j).ToString());
                            sw.WriteLine("    set subnet " + CellText + "/32");
                            sw.WriteLine("next");
                            j++;
                    }
                    if (CellText != null)
                    {
                        sw.WriteLine("end");
                    }
                    sw.Close();//Close the file
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            finally 
            {
                MessageBox.Show("Export Sussess");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            String filePath = @"C:\temp\" + textBox2.Text+".xlsx";
            string ExportTxtFilePath = textBox5.Text;
            int ExcelSheetNum = int.Parse(textBox4.Text);
            int yist_num = int.Parse(textBox1.Text);
            int sheet_Row_num = int.Parse(textBox3.Text);
            int sheet_Row_num_1 = int.Parse(textBox7.Text);
            int j = 1;
            string addgrpName = textBox6.Text;
            //Pass the filepath and filename to the StreamWriter Constructor
            StreamWriter sw = new StreamWriter("C:\\temp\\" + ExportTxtFilePath+".txt");//Export txt FilePath
            try
            {

                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);             
                ISheet sheet = workbook.GetSheetAt(ExcelSheetNum-1);
                int rowCount;
                if (sheet != null)
                {
                    if (sheet_Row_num_1 != 0)
                    {
                        rowCount = sheet_Row_num_1;
                    }
                    else
                    {
                        rowCount = sheet.LastRowNum+1;
                    }
                    sw.WriteLine("config firewall addrgrp");//Write a first line of text
                    sw.WriteLine("   edit " + addgrpName);
                    sw.Write("    append member yicst_0"+(yist_num).ToString()+" ");
                    for (int i = sheet_Row_num; i < rowCount; i++)
                    {
                        sw.Write("yicst_0"+(yist_num + j).ToString() + " ");                      
                        j++;
                    }
                    sw.WriteLine();
                    sw.WriteLine("next");
                    sw.WriteLine("end");
                    sw.Close();//Close the file
                }

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            finally
            {
                MessageBox.Show("Export Sussess");
            }
        }
    }
}
