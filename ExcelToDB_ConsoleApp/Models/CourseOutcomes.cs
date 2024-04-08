namespace ExcelToDB_ConsoleApp.Models
{
    public class CourseOutcomes
    {
        public Dictionary<string, double> CO { get; set; }

        public CourseOutcomes()
        {
            CO = new Dictionary<string, double>();
        }
    }
}
