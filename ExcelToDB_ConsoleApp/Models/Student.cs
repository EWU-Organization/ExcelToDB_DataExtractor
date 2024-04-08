namespace ExcelToDB_ConsoleApp.Models
{
    public class Student
    {
        public required string ID { get; set; }
        public required string Name { get; set; }
        public required Course Course { get; set; }
    }
}
