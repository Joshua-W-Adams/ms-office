using System;
using Microsoft.Office.Interop.Word;
using Ghostscript.NET.Samples;
using System.IO;

namespace ms_office {

    class wordFunctions {

        // Function to insert a photo at a specific location (range) within a word document
        public static void insertImage(Document doc, Range rng, String path) {

                // Create an InlineShape in the InlineShapes collection where the picture should be added later
                // It is used to get automatically scaled sizes.
                InlineShape autoScaledInlineShape = rng.InlineShapes.AddPicture(path);
                float scaledWidth = autoScaledInlineShape.Width;
                float scaledHeight = autoScaledInlineShape.Height;
                autoScaledInlineShape.Delete();

                // Create a new Shape and fill it with the picture
                Shape newShape = doc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
                newShape.Fill.UserPicture(path);

                // Convert the Shape to an InlineShape and optional disable Border
                InlineShape finalInlineShape = newShape.ConvertToInlineShape();
                finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Cut the range of the InlineShape to clipboard
                finalInlineShape.Range.Cut();

                // And paste it to the target Range
                rng.Paste();

        }

        public static void insertPdf(Document doc, Range rng, String pdfPath, String tempFilePath) {

            // confirm file exists before attempting to do any operations
            if (File.Exists(pdfPath) == true) {

                Rasterizer pdfConverter = new Rasterizer();

                // convert pdf to a jpg
                string photoPath = pdfConverter.pdfToJpeg(pdfPath, tempFilePath);

                // insert pdf image raster into word document
                insertImage(doc, rng, photoPath);

                // clean up photo generated
                generalFunctions.deleteFile(photoPath);

            }

        }

        public static void insertFile(Document doc, Range rng, string filename, string fileExt, dataStructures.row row, dataStructures.Parameters parameters) {

            string path = null;

            // define file path
            if (fileExt == ".jpg") {

                path = row.photoAssetsPath + @"\" + filename + ".jpg";

            } else if (fileExt == ".pdf") {

                path = row.pdfAssetsPath + @"\" + filename + ".pdf";

            }

            // confirm file exists before attempting to do any operations
            if (File.Exists(path) == true) {

                if (fileExt == ".jpg") {

                    insertImage(doc, rng, path);

                } else if (fileExt == ".pdf") {

                    insertPdf(doc, rng, path, parameters.temporaryFilesPath);

                }

            } else {

                // insert text indicating no file was found
                rng.Text = filename + fileExt + " was not found";

            }

        }

        public static void insertFileAtTable(Document doc, dataStructures.Parameters parameters, dataStructures.row row, string fileExt) {

            var count = doc.Tables.Count;
            object oMissing = System.Reflection.Missing.Value;

            // loop through all tables in word document
            for (var i = 1; i <= count; i++) {

                // Find table specific to current variable to be inserted into document
                if (doc.Tables[i].Title == row.insertTableName) {

                    var tbl = doc.Tables[i];
                    var rowInsertpoint = row.insertPoint[0];
                    var colInsertPoint = row.insertPoint[1];

                    // loop through file list and output to word document
                    for (int k = 0; k < row.oneDimArray.GetLength(0); k++) {

                        // add new row to table if applicable
                        if (tbl.Rows.Count - (rowInsertpoint - 1) < k + 1) {
                            tbl.Rows.Add(ref oMissing);
                        }

                        // get filename
                        string filename = row.oneDimArray[k].ToString();

                        // define insert range
                        Range rng = doc.Range(0, 0);
                        rng = tbl.Cell(k + rowInsertpoint, colInsertPoint).Range;

                        insertFile(doc, rng, filename, fileExt, row, parameters);

                    }

                }

            }

        }

        public static void insertFileAtTag(Document doc, dataStructures.Parameters parameters, dataStructures.row row, string fileExt) {

            // select empty range in document
            Range rng = doc.Range(0, 0);
            Range tmp = doc.Range(0, 0);

            // Find location to insert images
            if (rng.Find.Execute("<" + row.param + ">")) {

                // range is now set to bounds of the word "<" + parameter name + ">"

                // loop through file list and output to word document
                for (int k = 0; k < row.oneDimArray.GetLength(0); k++) {

                    string filename = row.oneDimArray[k].ToString();

                    // insert a new text value for the image to be pasted over
                    rng.InsertBefore("<" + filename + ">");

                    // define temporary range as inserted tag
                    tmp = doc.Range(0, 0);
                    tmp.Find.Execute("<" + filename + ">");

                    // insert File into document
                    insertFile(doc, tmp, filename, fileExt, row, parameters);

                }

                // clean up document tags
                rng.Find.Execute("<" + row.param + ">");
                rng.Delete();
                
            }

        }

        // lookup variable in word document and insert value
        public static void updateParameter(Document doc, dataStructures.row row) {

            doc.Variables[row.param].Value = row.value.ToString();
            doc.Fields.Update();

        }

        public static void updateTable(Document doc, dataStructures.row row) {

            var count = doc.Tables.Count;
            object oMissing = System.Reflection.Missing.Value;

            for (var i = 1; i <= count; i++) {

                // table found in word
                if (doc.Tables[i].Title == row.param) {

                    var tbl = doc.Tables[i];
                    var rowInsertpoint = row.insertPoint[0];
                    var colInsertPoint = row.insertPoint[1];

                    // loop through input table and output to word document
                    for (int k = 0; k < row.twoDimArray.GetLength(0); k++) {

                        // check if row exists
                        if (tbl.Rows.Count - (rowInsertpoint - 1) < k + 1) {
                            tbl.Rows.Add(ref oMissing);
                        }

                        for (int l = 0; l < row.twoDimArray.GetLength(1); l++) {

                            var val = row.twoDimArray[k, l];
                            string strVal;

                            // handle nulls in dataset
                            if (val != null) {

                                strVal = val.ToString();

                            } else {

                                strVal = val;

                            }

                            tbl.Cell(k + rowInsertpoint, l + colInsertPoint).Range.Text = strVal;

                        }

                    }

                }

            }

        }

        // Find a chart in a document and update its data
        public static void updateChartData(Document doc, dataStructures.row row) {

            var count = doc.InlineShapes.Count;

            // Loop through all inline shapes
            for (var i = 1; i <= count; i++) {

                String shapeType = doc.InlineShapes[i].Type.ToString();

                // chart found
                if (shapeType == "wdInlineShapeChart") {

                    Chart chart = doc.InlineShapes[i].Chart;
                    String title = chart.ChartTitle.Text.ToString();

                    // Specific chart found
                    if (title == row.param) {

                        excelFunctions.outputDataToWordChart(chart, row);

                    }

                }

            }

        }

    }

}