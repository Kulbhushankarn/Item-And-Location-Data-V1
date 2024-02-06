using System;
using App = Microsoft.Office.Interop.Excel.Application;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.CodeDom;
using System.Threading;

namespace Item_And_Location_Data_V1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        /*private static bool IsWindowVisible(IntPtr hWnd)
        {
            return IsWindowVisible(hWnd.ToInt32());
        }*/

        /*[DllImport("user32.dll")]
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
        }*/

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
            string inputFilePath = textBox_inputFile.Text; // Path where input File is present
            string output_SparePartsPath = Path.Combine(textBox_outputFile.Text, "Item stock bulk import format.xlsx"); // path where Workbook Location bulk import format.xlsx is to be saved
            string output_LocationsPath = Path.Combine(textBox_outputFile.Text, "Location bulk import format.xlsx");   // path where Location bulk import format.xlsx is to be saved
            App excelApp = new App();  // creating instance of Excel Application
            IntPtr pStm;
            Program.CoMarshalInterThreadInterfaceInStream(Guid.Empty, excelApp, out pStm);

            // Assuming threadId is the thread ID of the COM object

            Workbook inputWorkbook = TryOpenAndRecoverWorkbook(excelApp, inputFilePath);  // This is INPUT SHEET WHICH CAN BE CORRUPTED

            if (inputWorkbook != null)
            {
                Worksheet locationsWorksheet = inputWorkbook.Sheets["Location"];  //This is LOCATIONS WORKSHEET
                Worksheet sparePartWorksheet = inputWorkbook.Sheets["SparePart"]; // This is SPARE PARTS [ASSET MIGRATION] WORKSHEET

                RowData_Locations obj_Locations = new RowData_Locations();        // Creating Object of  RowData_Locations Class 
                RowData_SpareParts obj_SpareParts = new RowData_SpareParts();     // Creating Object og RowData_SpareParts Class

                progressBar1.Value = 25;

                List<RowData_Locations> dataToBeWrittenInLocations = obj_Locations.ReadDataFromLocationSheet(locationsWorksheet);  // This list holds the data that is to be written in Locations worksheet
                List<RowData_SpareParts> dataToBeWrittenInSpareParts = obj_SpareParts.ReadDataFromSparePartsSheet(sparePartWorksheet);  // This list holds the data that is to be written in assetMigration[spare parts]
                
                progressBar1.Value = 50;

                Marshal.ReleaseComObject(locationsWorksheet);
                Marshal.ReleaseComObject(sparePartWorksheet);
                inputWorkbook.Close(true);
                Marshal.ReleaseComObject(inputWorkbook);
                if (inputWorkbook != null)
                {
                    inputWorkbook = null;
                }
                if (locationsWorksheet != null)
                {
                    locationsWorksheet = null;
                }
                if (sparePartWorksheet != null)
                {
                    sparePartWorksheet = null;
                }
                GC.Collect();

                // --------------------- Handling Workbook Location bulk import format.xlsx ----------------------\\

                Workbook outputWorkbook_Locations = excelApp.Workbooks.Add();    // This workbook is for Locations file
                Worksheet outputLocationsWorksheet = outputWorkbook_Locations.Sheets[1];  // This worksheet is For LOCATIONS WORKSHEET
                Worksheet outputIntroLocationsWorksheet = outputWorkbook_Locations.Sheets.Add();  // This worksheet is For Intorduction WORKSHEET
                

                progressBar1.Value = 75;
                obj_Locations.WriteIntroduction(outputIntroLocationsWorksheet);
                obj_Locations.WriteDataInLocationSheet(outputLocationsWorksheet, dataToBeWrittenInLocations); // Writing Data in LOCATIONS WORKSHEET
                outputWorkbook_Locations.SaveAs(output_LocationsPath); // Saving Workbook Location bulk import format.xlsx

                Marshal.ReleaseComObject(outputLocationsWorksheet);
                Marshal.ReleaseComObject(outputIntroLocationsWorksheet);
                outputWorkbook_Locations.Close(true);
                Marshal.ReleaseComObject(outputWorkbook_Locations);
                if (outputWorkbook_Locations != null)
                {
                    outputWorkbook_Locations = null;
                }
                if (outputLocationsWorksheet != null)
                {
                    outputLocationsWorksheet = null;
                }
                if (outputIntroLocationsWorksheet != null)
                {
                    outputIntroLocationsWorksheet = null;
                }
                GC.Collect();

                // --------------------- Handling Workbook Item stock bulk import format.xlsx ----------------------\\

                Workbook outputWorkbook_SpareParts = excelApp.Workbooks.Add();// This workbook is for Spare Parts file
                Worksheet outputSparePartsWorksheet = outputWorkbook_SpareParts.Sheets[1];  // This worksheet is For Asset Migration [Spare Parts] WORKSHEET

                obj_SpareParts.WriteDataInSparePartsSheet(outputSparePartsWorksheet, dataToBeWrittenInSpareParts); // Writing Data in LOCATIONS WORKSHEET
                outputWorkbook_SpareParts.SaveAs(output_SparePartsPath); // Saving Workbook Item stock bulk import format.xlsx

                progressBar1.Value = 100;

                Marshal.ReleaseComObject(outputSparePartsWorksheet);
                outputWorkbook_SpareParts.Close(true);
                Marshal.ReleaseComObject(outputWorkbook_SpareParts);
                /*excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);*/
                if (outputWorkbook_SpareParts != null)
                {
                    outputWorkbook_SpareParts = null;
                }
                if (outputSparePartsWorksheet != null)
                {
                    outputSparePartsWorksheet = null;
                }
                /*if(excelApp!= null)
                {
                    excelApp = null;
                }*/
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Process p = Program.GetExcelProcess(excelApp);
                if (p != null)
                {
                    p.Kill();
                }
                else
                {
                    MessageBox.Show("No P");
                }
                // Get the process ID of the Excel process
                /*object unmarshaledObject;
                Program.CoGetInterfaceAndReleaseStream(pStm, Guid.Empty, out unmarshaledObject);*/

                // Release the unmarshaled interface
                /*if(unmarshaledObject != null)
                {
                    Marshal.ReleaseComObject(unmarshaledObject);

                }*/

                // Now, you can attempt to terminate the Excel process without elevated permissions
                /*IntPtr threadId = Program.GetCurrentThread();
                uint tid = Program.GetThreadId(threadId);

                Console.WriteLine($"Thread ID: {tid}");

                Process currentProcess = Process.GetCurrentProcess();
                uint excelProcessId = GetExcelProcessId(excelApp);
                MessageBox.Show(excelProcessId.ToString());*/
                //Process.GetProcessById(Program.GetProcessIdOfThread(threadId)).Kill();

                /*if (processId != 0)
                {
                    try
                    {
                        // Attempt to get the Excel process by ID
                        Process p = Process.GetProcessById(processId);

                        // Close the Excel process gracefully
                        if (p.CloseMainWindow())
                        {
                            // Wait for the process to exit or a timeout, then check if it's still running
                            if (!p.WaitForExit(5000) || !p.HasExited)
                            {
                                // Process did not exit gracefully, consider using Kill
                                p.Kill();
                            }
                        }
                        else
                        {
                            // CloseMainWindow failed, consider using Kill
                            p.Kill();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error closing Excel process: {ex.Message}");
                    }
                }*/
                MessageBox.Show($"Excel sheet create successfully! Please check atr {textBox_outputFile.Text}");
            } // If inputWorkbook is not null 

            else
            {
                MessageBox.Show("File Corrupted");
            } // If inputWorkbook is null

        }

        /*static uint GetExcelProcessId(Excel.Application excelApp)
        {
            IntPtr excelHwnd = (IntPtr)excelApp.Hwnd;
            if (excelHwnd != IntPtr.Zero)
            {
                // Get the process ID using the window handle
                uint excelProcessId = Program.GetProcessId(excelHwnd);
                return excelProcessId;
            }
            else
            {
                // Handle the case where the window handle is not valid
                return 0;
            }
        }*/
        private Workbook TryOpenAndRecoverWorkbook(App excelApp, string filePath)
        {
            Workbook workbook = null;

            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                // Attempt to recover if the file is corrupted
                if (comEx.ErrorCode == -2146827284)
                {
                    try
                    {
                        MessageBox.Show("Your File is Being Repaired","Corrupted Data Found",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        workbook = excelApp.Workbooks.Open(filePath, CorruptLoad: XlCorruptLoad.xlRepairFile);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        MessageBox.Show("Data is corrupted. File Cannot be repaired! Please repair the file manually then upload it");
                    }
                }
                else
                {
                    MessageBox.Show("Com exception occured");
                }
            }

            return workbook;
        }
        private void btn_exit(object sender, EventArgs e)
        {
            //ReleaseResources();
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
