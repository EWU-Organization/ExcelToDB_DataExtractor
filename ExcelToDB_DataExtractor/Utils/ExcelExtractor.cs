using ExcelToDB_DataExtractor.Database;
using ExcelToDB_DataExtractor.ExcelSheets;
using MongoDB.Bson;

namespace ExcelToDB_DataExtractor.Utils
{
    internal static class ExcelExtractor
    {
        public static void ProcessExcelFiles(string[] directoryPath)
        {
            Task.Run(() =>
            {
                foreach (var file in directoryPath)
                {
                    ProcessExcelFile(file);
                    ExcelUtility.Dispose();
                    Entry.Dispose();
                }
            });
        }

        static void ProcessExcelFile(string file)
        {
            Entry.ExtractInfo(file);

            if (Entry.IsExtractionComplete) SaveStudentsToMongoDB();
        }

        public static void SaveStudentsToMongoDB()
        {
            // Establish a connection to the MongoDB server
            var client = new MongoDatabase("mongodb://localhost:27017", "test");
            if (!client.IsConnectionEstablished()) return;

            // Convert Globals.Students to BsonDocument and insert into the collection
            foreach (var student in Globals.Students)
            {
                var document = student.ToBsonDocument();
                client.InsertDocument("students", document);
            }
        }
    }
}
