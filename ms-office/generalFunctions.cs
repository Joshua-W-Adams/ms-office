using System.IO;

namespace ms_office {

    class generalFunctions {

        public static void deleteFile(string fullFileName) {
            if (File.Exists(fullFileName)) {
                File.Delete(fullFileName);
            }
        }

    }

}
