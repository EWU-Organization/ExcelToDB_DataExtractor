namespace ExcelToDB_DataExtractor.Models
{
    public class Student
    {
        public required string StudentID { get; set; }
        public required string Name { get; set; }
        public required Course Course { get; set; }
        public required MarkDistribution MarkDistribution { get; set; }
    }
}
