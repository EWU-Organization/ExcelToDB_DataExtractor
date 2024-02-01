namespace ExcelToDB_DataExtractor.Models
{
    public class Student
    {
        public required string ID { get; set; }
        public required string Name { get; set; }
        public required List<Course> Course { get; set; }
    }
}
