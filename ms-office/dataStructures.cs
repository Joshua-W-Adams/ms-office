using System.Collections.Generic;

namespace ms_office {

    class dataStructures {

        // define class for storing passed JSON data
        public class Parameters {
            public string template { get; set; }
            public string documentNumber { get; set; }
            public string revision { get; set; }
            public string temporaryFilesPath { get; set; }
            public string documentationAssetsPath { get; set; }
            public List<row> inputs { get; set; }
        }

        public class row {
            public string param { get; set; }
            public string type { get; set; }
            public object value { get; set; }
            public int[] insertPoint { get; set; }
            public string insertTableName { get; set; }
            public string[,] twoDimArray { get; set; }
            public string[] oneDimArray { get; set; }
            public string photoAssetsPath { get; set; }
            public string pdfAssetsPath { get; set; }
        }

    }

}
