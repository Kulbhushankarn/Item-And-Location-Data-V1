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
        public List<string> Locations { get; set; }     //H
        public string CurrentStockNo { get; set; }      //R


        public List<RowData_SpareParts> ReadDataFromSparePartsSheet(Worksheet inputWorksheet)
        {

            Microsoft.Office.Interop.Excel.Range usedRange = inputWorksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);

            //---------------------Reading the data and storing it in a list of RowData------------------------//
            List<RowData_SpareParts> rows = new List<RowData_SpareParts>();

            int chunkSize = 1000;

            for (int rowIdx = 2; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    RowData_SpareParts singleRow = new RowData_SpareParts();  // Holds data for single Row

                    singleRow.ItemId = Convert.ToString(data[i, 61]);
                    singleRow.DrawingNo1 = Convert.ToString(data[i, 23]);
                    singleRow.DrawingNo2 = Convert.ToString(data[i, 28]);
                    singleRow.PositionNo1 = Convert.ToString(data[i, 29]);
                    singleRow.PositionNo2 = Convert.ToString(data[i, 30]);
                    singleRow.StockMin = Convert.ToString(data[i, 31]);
                    singleRow.StockMax = Convert.ToString(data[i, 32]);
                    singleRow.OpeningStockNumber = Convert.ToString(data[i, 33]);
                    singleRow.OpeningStockDate = Convert.ToString(data[i, 37]);
                   // singleRow.Locations = Convert.ToString(data[i, 35]);
                    singleRow.CurrentStockNo = Convert.ToString(data[i, 36]);



                    rows.Add(singleRow);

                }
            }
            return rows;

        }



    }
}
