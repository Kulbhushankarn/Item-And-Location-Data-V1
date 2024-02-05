using System;
using App=  Microsoft.Office.Interop.Excel.Application;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Item_And_Location_Data_V1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        private static bool IsWindowVisible(IntPtr hWnd)
        {
            return IsWindowVisible(hWnd.ToInt32());
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(int hWnd);
        private void ReleaseResources()
        {
            List<int> excelPID = new List<int>();

            // Get all processes
            Process[] prs = Process.GetProcesses();

            foreach (Process p in prs)
            {
                if (p.ProcessName == "EXCEL.EXE")
                {
                    // Check if the Excel process has a main window and is visible
                    if (IsWindowVisible(p.MainWindowHandle))
                    {
                        Console.WriteLine($"Excel process with PID {p.Id} is visible.");
                    }
                    else
                    {
                        excelPID.Add(p.Id);
                    }
                }
            }

            prs = Process.GetProcesses();

            foreach (Process p in prs)
            {
                if (p.ProcessName == "EXCEL" && !excelPID.Contains(p.Id))
                {
                    // Check if the Excel process has a main window and is visible
                    if (IsWindowVisible(p.MainWindowHandle))
                    {
                        Console.WriteLine($"Excel process with PID {p.Id} is visible.");
                    }
                    else
                    {
                        try
                        {
                            p.Kill();

                        }

                        catch
                        {
                            MessageBox.Show("Excel File not running in Background");
                            System.Windows.Forms.Application.Restart();
                        }
                    }
                }
            }
        }

        private void btn_selectExcelfile(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox_inputFile.Text = openFileDialog1.FileName;

        }

        private void btn_selectOutputfile(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            textBox_outputFile.Text = folderBrowserDialog1.SelectedPath;

        }

        private void btn_process(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            string inputFilePath = textBox_inputFile.Text;
            string output_SparePartsPath = Path.Combine(textBox_outputFile.Text,"Spare Parts.xlsx");
            string output_LocationsPath = Path.Combine(textBox_outputFile.Text,"Location.xlsx");
            App excelApp = new App();

            Workbook inputWorkbook = excelApp.Workbooks.Open(inputFilePath);
            Worksheet locationsWorksheet = inputWorkbook.Sheets["Location"];
            Worksheet sparePartWorksheet = inputWorkbook.Sheets["SparePart"];
            
            RowData_Locations obj_Locations = new RowData_Locations();
            RowData_SpareParts obj_SpareParts= new RowData_SpareParts();
            progressBar1.Value = 25;

            List<RowData_Locations> dataToBeWrittenInLocations = obj_Locations.ReadDataFromLocationSheet(locationsWorksheet);
            progressBar1.Value = 50;
            List<RowData_SpareParts> dataToBeWrittenInSpareParts=obj_SpareParts.ReadDataFromSparePartsSheet(sparePartWorksheet);

            Workbook outputWorkbook_Locations = excelApp.Workbooks.Add();
            progressBar1.Value = 75;
            Worksheet outputLocationsWorksheet = outputWorkbook_Locations.Worksheets.Add();

            obj_Locations.WriteDataInLocationSheet(outputLocationsWorksheet,dataToBeWrittenInLocations);
            outputWorkbook_Locations.SaveAs(output_LocationsPath);
            Workbook outputWorkbook_SpareParts = excelApp.Workbooks.Add();
            Worksheet outputSparePartsWorksheet=outputWorkbook_SpareParts.Worksheets.Add();

            obj_SpareParts.WriteDataInSparePartsSheet(outputSparePartsWorksheet, dataToBeWrittenInSpareParts);
            progressBar1.Value = 100;
            outputWorkbook_SpareParts.SaveAs(output_SparePartsPath);
            MessageBox.Show("Excel sheet create successfully! Please check in selected path.");

        }

        private void btn_exit(object sender, EventArgs e)
        {
            ReleaseResources();
            Environment.Exit(0);
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox_inputFile_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
