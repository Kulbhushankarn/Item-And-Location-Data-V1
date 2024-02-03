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

namespace Item_And_Location_Data_V1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_selectExcelfile(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();

        }

        private void btn_selectOutputfile(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            textBox_outputFile.Text = folderBrowserDialog1.SelectedPath;

        }

        private void btn_process(object sender, EventArgs e)
        {
            string inputFilePath = openFileDialog1.FileName;
            string output_SparePartsPath = Path.Combine(textBox_outputFile.Text,"Spare Parts.xlsx");
            App excelApp = new App();

            Workbook inputWorkbook = excelApp.Workbooks.Open(inputFilePath);
            Worksheet sparePartWorksheet = inputWorkbook.Sheets["SparePart"];

            RowData_SpareParts obj_SpareParts= new RowData_SpareParts();

            List<RowData_SpareParts> dataToBeWrittenInSpareParts=obj_SpareParts.ReadDataFromSparePartsSheet(sparePartWorksheet);

            Workbook outputWorkbook_SpareParts = excelApp.Workbooks.Add();
            Worksheet outputSparePartsWorksheet=outputWorkbook_SpareParts.Worksheets.Add();

            obj_SpareParts.WriteDataInSparePartsSheet(outputSparePartsWorksheet, dataToBeWrittenInSpareParts);
            outputWorkbook_SpareParts.SaveAs(output_SparePartsPath);

        }

        private void btn_exit(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }
    }
}
