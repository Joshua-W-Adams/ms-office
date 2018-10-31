using Microsoft.Office.Interop.Excel;

namespace ms_office {

    class excelFunctions {

        public static void outputDataToWordChart (Microsoft.Office.Interop.Word.Chart chart, dataStructures.row row) {
            
            // Get charts connected workbook
            Workbook wb = chart.ChartData.Workbook;
            Worksheet sht = wb.Worksheets[1];

            outputDataToSheet(sht, row);

            // refresh chart
            chart.ChartData.Activate();
            chart.ChartData.Workbook.Application.WindowState = -4140;

            // close workbook when finished processing data
            wb.Close();

        }

        public static void outputDataToSheet (Worksheet sht, dataStructures.row row) {

            var rowInsertpoint = row.insertPoint[0];
            var colInsertPoint = row.insertPoint[1];

            // loop through input data and output to excel
            for (int k = 0; k < row.twoDimArray.GetLength(0); k++) {

                for (int l = 0; l < row.twoDimArray.GetLength(1); l++) {

                    var val = row.twoDimArray[k, l];
                    sht.Cells[k + rowInsertpoint, l + colInsertPoint] = val;

                }

            }

        }
    }

}
