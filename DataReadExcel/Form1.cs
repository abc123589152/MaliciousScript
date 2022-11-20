
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.ComponentModel;
using System.Diagnostics.Eventing.Reader;

namespace MaliciousScriptCreate
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
            string addgrpName = textBox6.Text;
            //Pass the filepath and filename to the StreamWriter Constructor
            StreamWriter sw = new StreamWriter("C:\\temp\\"+ExportTxtFilePath+".ps1");//Export txt FilePath
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
                    sw.WriteLine("$username = read-host('Please Enter fortigae username')");
                    sw.Write("ssh 10.22.242.231 -l $username '");
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
                    sw.WriteLine("config firewall addrgrp");//Write a first line of text
                    sw.WriteLine("   edit " + addgrpName);
                    sw.Write("    append member yicst_0" + (yist_num).ToString() + " ");
                    for (int i = sheet_Row_num; i < rowCount; i++)
                    {
                        sw.Write("yicst_0" + ((yist_num-rowCount+1) + j).ToString() + " ");
                        j++;
                    }
                    sw.WriteLine();
                    sw.WriteLine("next");
                    sw.WriteLine("end'");
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
        private void button1_Click(object sender, EventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-TW");
            Thread.CurrentThread.CurrentCulture.DateTimeFormat.Calendar = new TaiwanCalendar();//use taiwan calendar
            String filePath = @"C:\temp\" + textBox2.Text + ".xlsx";
            string NowDate = DateTime.Now.ToString("yyyMMdd");//Get Taiwan DateTime
            string DNSmaliciousDirectoryPath = @"C:\temp\"+NowDate;
            System.IO.Directory.CreateDirectory(DNSmaliciousDirectoryPath);//Create Directory for DNS malicious script
            if (Directory.Exists(DNSmaliciousDirectoryPath))
            {
                StreamWriter DNSmalicious_main = new StreamWriter("C:\\temp\\"+ NowDate+"\\DNSmalicious_main.ps1");//Export txt FilePath
                StreamWriter Zone = new StreamWriter("C:\\temp\\" + NowDate + "\\Add_DNS_Zone.ps1");
                StreamWriter A = new StreamWriter("C:\\temp\\" + NowDate + "\\Add_DNS_A.ps1");
                StreamWriter CNAME = new StreamWriter("C:\\temp\\" + NowDate + "\\Add_DNS_CNAME.ps1");
                StreamWriter DN = new StreamWriter("C:\\temp\\" + NowDate + "\\DN.txt");
                try
                {
                    IWorkbook workbook = null;
                    FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);//Get Excel file to open and read content 
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);                                                         
                    int rowCount;
                    ISheet sheet = workbook.GetSheetAt(0);
                    rowCount = sheet.LastRowNum;
                    DNSmalicious_main.WriteLine("[int]$Time = 300");
                    DNSmalicious_main.WriteLine("$Lenght = $Time / 100");
                    DNSmalicious_main.WriteLine("echo 'Start to DNS Malicious addding...'");
                    DNSmalicious_main.WriteLine("$Del_DNS_A = Test-Path -Path ./Del_DNS_A.ps1 -PathType Leaf");
                    DNSmalicious_main.WriteLine("$Add_DNS_Zone = Test-Path -Path ./Add_DNS_Zone.ps1 -PathType Leaf");
                    DNSmalicious_main.WriteLine("$Add_DNS_A = Test-Path -Path ./Add_DNS_A.ps1 -PathType Leaf");
                    DNSmalicious_main.WriteLine("$Add_DNS_CNAME = Test-Path -Path ./Add_DNS_CNAME.ps1 -PathType Leaf");
                    DNSmalicious_main.WriteLine("if($Del_DNS_A -eq $true){");
                    DNSmalicious_main.WriteLine("   echo 'Execute Del_DNS_A.ps1 please wait...' >> start.log");
                    DNSmalicious_main.WriteLine("   & .\\Del_DNS_A.ps1");
                    DNSmalicious_main.WriteLine("   echo 'done' >> start.log");
                    DNSmalicious_main.WriteLine("   Start-Sleep 5");
                    DNSmalicious_main.WriteLine("}else{");
                    DNSmalicious_main.WriteLine("   echo 'Del_DNS_A.ps1 not exists' >> start.log");
                    DNSmalicious_main.WriteLine("}");
                    DNSmalicious_main.WriteLine("if($Add_DNS_Zone -eq $true){");
                    DNSmalicious_main.WriteLine("   echo 'Execute Add_DNS_Zone.ps1 please wait...' >> start.log");
                    DNSmalicious_main.WriteLine("   & .\\Add_DNS_Zone.ps1");
                    DNSmalicious_main.WriteLine("   echo 'done' >> start.log");
                    DNSmalicious_main.WriteLine("   Start-Sleep 5");
                    DNSmalicious_main.WriteLine("}else{");
                    DNSmalicious_main.WriteLine("   echo 'Add_DNS_Zone.ps1 not exists' >> start.log");
                    DNSmalicious_main.WriteLine("}");
                    DNSmalicious_main.WriteLine("if($Add_DNS_A -eq $true){");
                    DNSmalicious_main.WriteLine("   echo 'Execute Add_DNS_A.ps1 please wait...' >> start.log");
                    DNSmalicious_main.WriteLine("   & .\\Add_DNS_A.ps1");
                    DNSmalicious_main.WriteLine("   echo 'done' >> start.log");
                    DNSmalicious_main.WriteLine("   Start-Sleep 5");
                    DNSmalicious_main.WriteLine("}else{");
                    DNSmalicious_main.WriteLine("   echo 'Add_DNS_A.ps1 not exists' >> start.log");
                    DNSmalicious_main.WriteLine("}");
                    DNSmalicious_main.WriteLine("if($Add_DNS_CNAME -eq $true){");
                    DNSmalicious_main.WriteLine("   echo 'Execute Add_DNS_CNAME.ps1 please wait...' >> start.log");
                    DNSmalicious_main.WriteLine("   & .\\Add_DNS_CNAME.ps1");
                    DNSmalicious_main.WriteLine("   echo 'done' >> start.log");
                    DNSmalicious_main.WriteLine("   Start-Sleep 5");
                    DNSmalicious_main.WriteLine("}else{");
                    DNSmalicious_main.WriteLine("   echo 'Add_DNS_CNAME.ps1 not exists' >> start.log");
                    DNSmalicious_main.WriteLine("}");
                    DNSmalicious_main.WriteLine("For ($Time; $Time -gt 0; $Time--) {");
                    DNSmalicious_main.WriteLine("$min = [int](([string]($Time/60)).split('.')[0])");
                    DNSmalicious_main.WriteLine("$text = \" \" + $min + \" minutes \" + ($Time % 60) + \" seconds left\"");
                    DNSmalicious_main.WriteLine("Write-Progress -Activity \"Watiting for to start test nslookup\" -Status $Text -PercentComplete ($Time / $Lenght)");
                    DNSmalicious_main.WriteLine("Start-Sleep 1");
                    DNSmalicious_main.WriteLine("}");
                    DNSmalicious_main.WriteLine("$nslookupdeltest = Test-Path -Path ./nslookupdeltext.ps1 -PathType Leaf");
                    DNSmalicious_main.WriteLine("if($nslookupdeltest -eq $true){");
                    DNSmalicious_main.WriteLine("   echo 'Execute nslookupdeltext.ps1 please wait...' >> start.log");
                    DNSmalicious_main.WriteLine("   & ./nslookupdeltext.ps1");
                    DNSmalicious_main.WriteLine("   echo 'done' >> start.log");
                    DNSmalicious_main.WriteLine("}else{");
                    DNSmalicious_main.WriteLine("   echo 'nslookupdeltext.ps1 not exists' >> start.log");
                    DNSmalicious_main.WriteLine("}");
                    DN.WriteLine("新增:");
                    for (int i = 0; i < rowCount; i++)
                    {
                        IRow Getthismonthmalicious = sheet.GetRow(i);
                        if (Getthismonthmalicious.GetCell(0).ToString() != "")
                        {
                            Zone.WriteLine("dnscmd stsbow08 /zoneadd " + Getthismonthmalicious.GetCell(0).ToString() + " /dsprimary");
                            A.WriteLine("dnscmd stsbow08 /Recordadd " + Getthismonthmalicious.GetCell(0).ToString() + " @ A 123.123.123.123");
                            CNAME.WriteLine("dnscmd stsbow08 /Recordadd " + Getthismonthmalicious.GetCell(0).ToString() + " * CNAME "+ Getthismonthmalicious.GetCell(0).ToString());
                            DN.WriteLine(Getthismonthmalicious.GetCell(0).ToString());
                            DNSmalicious_main.WriteLine("nslookup " + Getthismonthmalicious.GetCell(0).ToString() + " >> nslookupfortest.log");
                        }
                        else
                        {
                            break;
                        }                   
                    }
                    Zone.Close();
                    A.Close();
                    CNAME.Close();
                    DNSmalicious_main.Close();
                    ISheet Deldnssheet = workbook.GetSheetAt(1);
                    int rowCountDeldns = Deldnssheet.LastRowNum;
                    DN.WriteLine("-----------------------------------");
                    DN.WriteLine("移除:");
                    if (rowCountDeldns != 0)
                    {                        
                        StreamWriter DelDNSA = new StreamWriter("C:\\temp\\" + NowDate + "\\Del_DNS_A.ps1");
                        StreamWriter DelDNSnslookuptest = new StreamWriter("C:\\temp\\" + NowDate + "\\nslookupdeltext.ps1");

                        for (int i = 0; i < rowCountDeldns+1; i++)
                        {
                            IRow deldnsrow = Deldnssheet.GetRow(i);
                            DelDNSA.WriteLine("dnscmd stsbow08 /zonedelete " + deldnsrow.GetCell(0).ToString() + " /DsDel /f");
                            DN.WriteLine(deldnsrow.GetCell(0).ToString());
                            DelDNSnslookuptest.WriteLine("nslookup " + deldnsrow.GetCell(0).ToString()+ " >> nslookupfortest.log");
                        }
                        DelDNSA.Close();
                        DelDNSnslookuptest.Close();
                        DN.Close();
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
                //Pass the filepath and filename to the StreamWriter Constructor
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
