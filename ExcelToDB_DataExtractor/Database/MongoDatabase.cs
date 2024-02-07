using MongoDB.Bson;
using MongoDB.Driver;

namespace ExcelToDB_DataExtractor.Database
{
    internal class MongoDatabase
    {
        private IMongoDatabase _Database;
        private MongoClient _Client;

        public MongoDatabase(string connectionString, string databaseName)
        {
            // Establish a connection to the MongoDB server
            _Client = new MongoClient(connectionString);

            // Get a reference to the database
            _Database = _Client.GetDatabase(databaseName);
        }

        public bool IsConnectionEstablished()
        {
            try
            {
                // Use the Ping command to check if the server is responding
                _Database.RunCommandAsync((Command<BsonDocument>)"{ping:1}").Wait();
                return true;
            }
            catch (MongoCommandException ex)
            {
                // Handle expected exceptions here
                Console.Error.WriteLine($"Failed to connect to MongoDB server: {ex.Message}");
                return false;
            }
        }

        public void InsertDocument(string collectionName, BsonDocument document)
        {
            // Get a reference to the collection
            var collection = _Database.GetCollection<BsonDocument>(collectionName);

            // Insert the document into the collection
            collection.InsertOne(document);
        }

        public List<BsonDocument> GetDocuments(string collectionName)
        {
            // Get a reference to the collection
            var collection = _Database.GetCollection<BsonDocument>(collectionName);

            // Retrieve all documents from the collection
            return collection.Find(new BsonDocument()).ToList();
        }
    }
}
