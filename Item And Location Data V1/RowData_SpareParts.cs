using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Item_And_Location_Data_V1
{
    internal class RowData_SpareParts
    {
        public string ItemId { get; set; }  //BI column
        public string DrawingNo1 { get; set; }  //BE
        public string DrawingNo2 { get; set; }   //BF
        public string PositionNo1 { get; set; }   //BG
        public string PositionNo2 { get; set; }   //BH
        public string StockMin { get; set; }      //K
        public string StockMax { get; set; }      //L
        public string OpeningStockNumber { get; set; }   //O
        public string OpeningStockDate { get; set; }    //T
        public List<string> Locations { get; set; }     //BJ
        public List<string> CurrentStockNo { get; set; }      //R

        public RowData_SpareParts()
        {
            Locations = new List<string> { "", "", "" };
            CurrentStockNo = new List<string> { "", "", "" };
        }
        public List<RowData_SpareParts> ReadDataFromSparePartsSheet(Worksheet inputWorksheet)
        {

            Microsoft.Office.Interop.Excel.Range usedRange = inputWorksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);

            //---------------------Reading the data and storing it in a list of RowData------------------------//
            List<RowData_SpareParts> rows = new List<RowData_SpareParts>();

            int chunkSize = 1000;

            for (int rowIdx = 10; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    RowData_SpareParts singleRow = new RowData_SpareParts();  // Holds data for single Row

                    int locationCount = 1;
                    int temp = i;

                    singleRow.ItemId = Convert.ToString(data[i, 61]);
                    singleRow.DrawingNo1 = Convert.ToString(data[i, 57]);
                    singleRow.DrawingNo2 = Convert.ToString(data[i, 58]);
                    singleRow.PositionNo1 = Convert.ToString(data[i, 59]);
                    singleRow.PositionNo2 = Convert.ToString(data[i, 60]);
                    singleRow.StockMin = Convert.ToString(data[i, 11]);
                    singleRow.StockMax = Convert.ToString(data[i, 12]);
                    singleRow.OpeningStockNumber = Convert.ToString(data[i, 15]);
                    singleRow.OpeningStockDate = Convert.ToString(data[i, 16]);
                    singleRow.Locations[0] = (Convert.ToString(data[i, 62]));
                    singleRow.CurrentStockNo[0] = (Convert.ToString(data[i, 18]));

                    while (temp < rowCount - 1 && Convert.ToString(data[temp + 1, 61]) == "")
                    {
                        singleRow.Locations[locationCount] = (Convert.ToString(data[temp + 1, 62]));
                        singleRow.CurrentStockNo[locationCount] = (Convert.ToString(data[temp + 1, 18]));
                        locationCount++;
                        temp++;
                        if (locationCount == 3)
                        {
                            break;
                        }
                    }

                    if (singleRow.ItemId != "")
                    {

                        rows.Add(singleRow);
                    }

                }
            }
            return rows;

        }


        public void WriteDataInSparePartsSheet(Worksheet worksheet, List<RowData_SpareParts> sparePartsData)
        {
            WriteHeaders(worksheet);
            worksheet.Name = "Asset Migration";

            int startRow = 10;
            int numRows = sparePartsData.Count;
            int numColumns = 18;

            Range dataRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow + numRows - 1, numColumns]];

            dataRange.NumberFormat = "@";

            string[,] dataArray = new string[numRows, numColumns];

            for (int i = 0; i < numRows; i++)
            {
                RowData_SpareParts rowData = sparePartsData[i];
                dataArray[i, 0] = rowData.ItemId;
                dataArray[i, 1] = rowData.DrawingNo1;
                dataArray[i, 2] = rowData.PositionNo1;
                dataArray[i, 3] = rowData.DrawingNo2;
                dataArray[i, 4] = rowData.PositionNo2;
                dataArray[i, 8] = rowData.StockMin;
                dataArray[i, 9] = rowData.StockMax;
                dataArray[i, 10] = rowData.OpeningStockNumber;
                dataArray[i, 11] = rowData.OpeningStockDate;


                if (rowData.Locations.Count > 2)
                {
                    dataArray[i, 12] = rowData.Locations[0];
                    dataArray[i, 14] = rowData.Locations[1];
                    dataArray[i, 16] = rowData.Locations[2];

                }

                if (rowData.CurrentStockNo.Count > 2)
                {
                    dataArray[i, 13] = rowData.CurrentStockNo[0];
                    dataArray[i, 15] = rowData.CurrentStockNo[1];
                    dataArray[i, 17] = rowData.CurrentStockNo[2];

                }


            }

            dataRange.Value = dataArray;
        }
        public void WriteHeaders(Worksheet worksheet)
        {
            string[] headers = { "Item No", "Drawing No.1", "Position No.1", "Drawing No.2", "Position No.2", "Criticality", "Phase Out", "Wear & Tear", "Min Stock", "Max Stock", "Opening stock", "Opening stock date", "Location Code 1", "Location Stock 1", "Location Code 2", "Location Stock 2", "Location Code 3", "Location Stock 3" };

            worksheet.Columns.ColumnWidth = 27;

            Range headerRange = worksheet.Rows[1];

            // Set headers in bold and red color
            headerRange.Font.Bold = true;
            headerRange.Font.Color = XlRgbColor.rgbRed;
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1] = headers[i];
            }
        }

        public void WriteContent(Worksheet worksheet)
        {
            // Set data in non-bold and non-colored cells
            worksheet.Cells[3, 6] = "Safety or Operational";
            worksheet.Cells[3, 7] = "Yes";
            worksheet.Cells[3, 8] = "Yes";
            worksheet.Range["F5:H5"].Merge();
            worksheet.Cells[5, 6] = "Default: No";
            worksheet.Range["L5:M5"].Merge();
            worksheet.Cells[5, 12] = "DD-MM-YYYY";
            worksheet.Cells[5, 13] = "Set Location 1 as default location";

            // Set bold and colored cells
            Range boldBlueRange = worksheet.Range["A8:D9"];
            boldBlueRange.Font.Bold = true;
            boldBlueRange.Font.Color = XlRgbColor.rgbBlue;

            // Text for row 8
            worksheet.Cells[8, 1] = "Date Format Should be DD-MM-YYYY as Text";

            // Text for row 9
            worksheet.Cells[9, 1] = "Migration of Data Will Start from 10 Row Only.";
        }




    }
}
