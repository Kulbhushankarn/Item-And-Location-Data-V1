using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Item_And_Location_Data_V1
{
    internal class RowData_Locations
    {

        public string LocationID { get; set; }
        public string LocationName { get; set; }


        public List<RowData_Locations> ReadDataFromLocationSheet(Worksheet inputWorksheet)
        {

            Microsoft.Office.Interop.Excel.Range usedRange = inputWorksheet.UsedRange;
            object[,] data = usedRange.Value;
            int rowCount = data.GetLength(0);

            //---------------------Reading the data and storing it in a list of RowData------------------------//
            List<RowData_Locations> rows = new List<RowData_Locations>();

            int chunkSize = 1000;

            for (int rowIdx = 10; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    RowData_Locations singleRow = new RowData_Locations();  // Holds data for single Row

                    singleRow.LocationName = Convert.ToString(data[i, 1]);
                    singleRow.LocationID = Convert.ToString(data[i, 11]);

                    rows.Add(singleRow);

                }
            }
            return rows;

        }

        public void WriteHeaders(Worksheet worksheet)
        {
            // Headers for Row 1
            string[] row1Headers = { "ID", "Code", "Name", "Parent", "Remarks", "System Location", "Item Location", "Component Location", "SHEQ Location", "ReadOnly", "Disabled", "Mark As delete", "Bunker Location", "Bunker Capacity" };

            // Headers for Row 2
            string[] row2Headers = { "LocationID", "LocationCode", "Name", "ParentNo", "Remarks", "IsSystemLocationREQ", "IsItemLocationREQ", "IsComponentLocationREQ", "IsSHEQLocationREQ", "ReadOnly", "Disabled", "IsDeleted", "IsBunkerLocation", "BunkerCapacity" };

            worksheet.Columns.ColumnWidth = 20;

            // Set headers in bold for Row 1
            Range row1HeaderRange = worksheet.Rows[1];
            row1HeaderRange.Font.Bold = true;
            row1HeaderRange.Interior.Color = XlRgbColor.rgbSkyBlue;

            for (int i = 0; i < row1Headers.Length; i++)
            {
                worksheet.Cells[1, i + 1] = row1Headers[i];
            }

            // Set headers for Row 2
            Range row2HeaderRange = worksheet.Rows[2];
            row2HeaderRange.Font.Bold = true;
            row2HeaderRange.Interior.Color = XlRgbColor.rgbPaleGreen;

            for (int i = 0; i < row2Headers.Length; i++)
            {
                worksheet.Cells[2, i + 1] = row2Headers[i];
            }

            Range additionalTextRange = worksheet.Range["A9:D9"];
            additionalTextRange.Merge();
            additionalTextRange.Value = "Migration of Data Will Start from 10 Row Only.";
            additionalTextRange.Font.Bold = true;
            additionalTextRange.Font.Color = XlRgbColor.rgbBlue;
            additionalTextRange.Interior.Color = XlRgbColor.rgbYellow;
        }

        public void WriteDataInLocationSheet(Worksheet worksheet, List<RowData_Locations> locationsData)
        {
            WriteHeaders(worksheet);

            int startRow = 10; // Start writing from the second row
            int numRows = locationsData.Count;
            int numColumns = 14; // Adjust the number of columns as needed

            Range dataRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow + numRows - 1, numColumns]];

            dataRange.NumberFormat = "@";

            string[,] dataArray = new string[numRows, numColumns];

            for (int i = 0; i < numRows; i++)
            {
                RowData_Locations rowData = locationsData[i];
                dataArray[i, 1] = rowData.LocationID;
                dataArray[i, 2] = rowData.LocationName;
                dataArray[i, 5] = "0";
                dataArray[i, 6] = "1";
                dataArray[i, 7] = "0"; 
                dataArray[i, 8] = "0";
                dataArray[i, 12] = "0";
            }

            dataRange.Value = dataArray;
        }
    }
}
