using MongoDB.Bson;
using MongoDB.Driver;

namespace ExcelToDB_ConsoleApp.Database
{
    internal class MongoDatabase
    {
        private IMongoDatabase _Database;
        private MongoClient _Client;
        private IMongoCollection<BsonDocument> _Collection;

        public MongoDatabase(string connectionString, string databaseName, string collectionName)
        {
            try
            {
                // Establish a connection to the MongoDB server
                _Client = new MongoClient(connectionString);

                // Check if the database exists
                List<string> databaseList = _Client.ListDatabaseNames().ToList();
                if (!databaseList.Contains(databaseName))
                {
                    throw new MongoException($"Database '{databaseName}' does not exist.");
                }

                // Get a reference to the database
                _Database = _Client.GetDatabase(databaseName);

                // Get a reference to the collection
                _Collection = _Database.GetCollection<BsonDocument>(collectionName);
            }
            catch (MongoClientException ex)
            {
                Console.Error.WriteLine(ex.Message);
                Environment.Exit(-1);
            }
            catch (MongoException ex)
            {
                Console.Error.WriteLine(ex.Message);
                Environment.Exit(-2);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                Environment.Exit(-3);
            }
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
        
        public List<BsonDocument> GetDocuments(string collectionName)
        {
            // Retrieve all documents from the collection
            return _Collection.Find(new BsonDocument()).ToList();
        }

        public void InsertDocument(BsonDocument document)
        {
            // Insert the document into the collection
            _Collection.InsertOne(document);
        }

        public void UpdateDocument(string field, string oldValue, string newValue)
        {
            // Define your search criteria
            var filter = Builders<BsonDocument>.Filter.Eq(field, oldValue);

            // Check if the document exists
            var existingDocument = _Collection.Find(filter).FirstOrDefault();

            if (existingDocument != null)
            {
                // Document exists, update it
                var update = Builders<BsonDocument>.Update.Set(field, newValue);
                _Collection.UpdateOne(filter, update);
                Console.WriteLine("Document updated successfully.");
            }
            else
            {
                // Document doesn't exist, insert a new one
                var documentToInsert = new BsonDocument
                {
                    { "yourField", "yourValue" },
                    { "anotherField", "anotherValue" }
                    // Add more fields as needed
                };
                _Collection.InsertOne(documentToInsert);
                Console.WriteLine("New document inserted successfully.");
            }
        }
    }
}
