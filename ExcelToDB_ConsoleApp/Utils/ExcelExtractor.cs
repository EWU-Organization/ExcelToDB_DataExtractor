using ExcelToDB_ConsoleApp.Database;
using ExcelToDB_ConsoleApp.ExcelSheets;
using ExcelToDB_ConsoleApp.Models;
using MongoDB.Bson;
using System.Configuration;

namespace ExcelToDB_ConsoleApp.Utils
{
    internal static class ExcelExtractor
    {
        private static MongoDatabase? _Client;
        public static async Task ProcessExcelFiles(string[] directoryPath)
        {
            if (ConnectToMongoDB())
            {
                foreach (var file in directoryPath)
                {
                    ProcessExcelFile(file);
                    ExcelUtility.Dispose();
                    Entry.Dispose();
                }
            }
        }

        static void ProcessExcelFile(string file)
        {
            Entry.ExtractInfo(file);

            if (Entry.IsExtractionComplete) SaveStudentsToMongoDB();
        }

        private static bool ConnectToMongoDB()
        {
            string ConnectionString = ConfigurationManager.AppSettings["ConnectionString"];
            string DatabaseName = ConfigurationManager.AppSettings["DatabaseName"];
            string CollectionName = ConfigurationManager.AppSettings["CollectionName"];
            // Establish a connection to the MongoDB server
            _Client = new MongoDatabase(ConnectionString, DatabaseName, CollectionName);
            return _Client.IsConnectionEstablished();
        }

        public static void SaveStudentsToMongoDB()
        {
            if (!_Client.IsConnectionEstablished()) return;

            // Convert Globals.Students to BsonDocument and insert into the collection
            foreach (Student student in Globals.Students)
            {
                BsonDocument document = student.ToBsonDocument();
                _Client.InsertDocument(document);
            }
        }
    }
}
