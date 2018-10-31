// shortens namespace references in code
using System;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System.IO;
using System.Diagnostics;

// set namespace that all code in this file will be associated to
namespace ms_office
{   

    // define container class for main application
    class Program
    {

        // Main method of application
        // Entry point of code
        // static = do not need to construct an instance of the parent class to run this method.
        // eg. Program.Main() works
        // void = method does not return any results
        // string[] args = arguements passed to application from shell console.
        static void Main(string[] args) {

            // Ensure json arguements have been passed to the application
            //if (args.Length != 0) {

                //var json = args[0];
                var json = testCases(1);
                var json2 = testCases(2);

                // Convert JSON text string to C# object
                dataStructures.Parameters parameters = JsonConvert.DeserializeObject<dataStructures.Parameters>(json);

                dataStructures.Parameters parameters2 = JsonConvert.DeserializeObject<dataStructures.Parameters>(json2);
                // Call main loop for creating word document
                createWordDocument(parameters);
                createWordDocument(parameters2);


            //}

        }

        // returns a json string to be used as a test case for the application
        public static string testCases(int caseNo) {
            string testCase = "";
            // values, tables and charts
            if (caseNo == 1) {
                // @ = verbatum string literal command. This means do not apply any interpretations to characters until the
                // next quotation is reached.
                testCase = @"{
                    template: 'C:\\solidintegrity\\server\\templates\\word\\mech-integrity.docm',
                    documentNumber: 'SO-TEST-001',
                    revision: 'A',
                    temporaryFilesPath: 'C:\\solidintegrity\\assets\\generated_files',
                    documentationAssetsPath: 'C:\\solidintegrity\\assets\\generated_documents\\1',
                    inputs: [{
                        param: 'var_title',
                        type: 'value',
                        value: 'Test Report',
                        insertPoint: null,
                        insertTableName: null,
                        twoDimArray: null,
                        oneDimArray: null,
                        photoAssetsPath: null,
                        pdfAssetsPath: null
                    }, {
                        param: 'var_risk_matrix',
                        type: 'table',
                        value: null,
                        insertPoint: [1,3],
                        insertTableName: null,
                        twoDimArray: [['1', '2', '3'], ['7', '8', '9'], ['4', '5', '6']],
                        oneDimArray: null,
                        photoAssetsPath: null,
                        pdfAssetsPath: null
                    }, {
                        param: 'var_chart_inspection_and_remediation',
                        type: 'chart',
                        value: null,
                        insertPoint: [2,1],
                        insertTableName: null,
                        twoDimArray: [['Test1', '1', '1'], ['Test2', '1', '1']],
                        oneDimArray: null,
                        photoAssetsPath: null,
                        pdfAssetsPath: null
                    }]
                  }";
            // pdfs and photos
            } else if (caseNo == 2) {
                testCase = @"{
                    template: 'C:\\solidintegrity\\server\\templates\\word\\rt-inspection-workpack.docm',
                    documentNumber: 'SO-TEST-002',
                    revision: 'A',
                    temporaryFilesPath: 'C:\\solidintegrity\\assets\\generated_files',
                    documentationAssetsPath: 'C:\\solidintegrity\\assets\\generated_documents\\1',
                    inputs: [{
                        param: 'var_photos',
                        type: 'photo',
                        value: null,
                        insertPoint: [2,7],
                        insertTableName: 'var_cml_details',
                        twoDimArray: null,
                        oneDimArray: ['1', '2', '3'],
                        photoAssetsPath: 'C:\\solidintegrity\\assets\\photos',
                        pdfAssetsPath: null
                    }, {
                        param: 'var_pids',
                        type: 'pdf',
                        value: null,
                        insertPoint: null,
                        insertTableName: null,
                        twoDimArray: null,
                        oneDimArray: ['DWG-PID-01'],
                        photoAssetsPath: null,
                        pdfAssetsPath: 'C:\\solidintegrity\\assets\\pids'
                    }]
                  }";
            }
            return testCase;
        }

        // Send application outputs to file for debugging purposes
        public static void printTodebugFile (String value, String logFile) {
            using (StreamWriter file = new StreamWriter(logFile, true)) {
                // write passed value to end of file
                file.WriteLine(value);
            }
        }

        // Determine if application is running
        public static Boolean appRunning(string processName) {
            Process[] processes = Process.GetProcessesByName(processName);
            if (processes.Length == 0) {
                return false;
            } else { 
                return true;
            }
        }

        // get runnign instance of application from operating system
        // <T> = application object type
        public static Object getApplication <T>(string appName) {
            T app = (T) System.Runtime.InteropServices.Marshal.GetActiveObject(appName);
            return app;
        } 

        // Main loop to construct the word document from a template based on the passed JSON
        public static void createWordDocument (dataStructures.Parameters parameters) {

            Microsoft.Office.Interop.Word.Application app;

            // Create word application
            if (appRunning("winword") == true) {
                app = (Microsoft.Office.Interop.Word.Application) getApplication<Microsoft.Office.Interop.Word.Application>("Word.Application");
            } else {
                app = new Microsoft.Office.Interop.Word.Application();
            }

            // Configure app for improved processing speed
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            app.ScreenUpdating = false;

            // set parameters for more reliable code execution
            // setting app.Visible = true in the past has resulted in more reliable generation of charts
            // however this can result in the c# app becoming stuck when word attempts to close the document if it is not focused
            // by the operating system.
            app.Visible = false;

            // Open Template
            Document doc = app.Documents.Open(parameters.template);

            // loop through all rows in the inputs object and perform operations
            foreach (var row in parameters.inputs) {

                if (row.type == "value") {

                    wordFunctions.updateParameter(doc, row);

                } else if (row.type == "table") {

                    wordFunctions.updateTable(doc, row);

                } else if (row.type == "photo") {

                    if (row.insertTableName != null) {

                        wordFunctions.insertFileAtTable(doc, parameters, row, ".jpg");

                    } else {

                        wordFunctions.insertFileAtTag(doc, parameters, row, ".jpg");

                    }
                    
                } else if (row.type == "pdf") {

                    if (row.insertTableName != null) {

                        wordFunctions.insertFileAtTable(doc, parameters, row, ".pdf");

                    } else {

                        wordFunctions.insertFileAtTag(doc, parameters, row, ".pdf");

                    }

                } else if (row.type == "chart") {

                    wordFunctions.updateChartData(doc, row);

                }

            }

            // Copy new item to clipboard to supress clipboard message
            // doc.Paragraphs[doc.Paragraphs.Count].Range.Copy();
            doc.Characters[doc.Characters.Count].Copy();

            // Create folder to store document
            // will do nothing if the directory already exists
            System.IO.Directory.CreateDirectory(parameters.documentationAssetsPath);

            string wordPath = parameters.documentationAssetsPath + @"\" + parameters.documentNumber + "-" + parameters.revision + ".docm";
            string pdfPath = parameters.documentationAssetsPath + @"\" + parameters.documentNumber + "-" + parameters.revision + ".pdf";

            // delete files if they already exist
            generalFunctions.deleteFile(wordPath);
            generalFunctions.deleteFile(pdfPath);

            // Save changes
            doc.SaveAs(wordPath);
            doc.SaveAs2(pdfPath, WdSaveFormat.wdFormatPDF);
            doc.Close();

            // return performance defaults
            app.DisplayAlerts = WdAlertLevel.wdAlertsAll;
            app.ScreenUpdating = true;

            // Close word.
            // app.Quit();

        }

    }
}
