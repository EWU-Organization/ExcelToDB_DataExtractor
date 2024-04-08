namespace ExcelToDB_ConsoleApp.Models
{
    public class Course
    {
        public required string ID { get; set; }
        public required string Semester { get; set; }
        public required string CourseCode { get; set; }
        public required int Section { get; set; }
        public required string Instructor { get; set; }
        public required string CourseTitle { get; set; }
        public required double Credit { get; set; }
        public required string LetterGrade { get; set; }
        //public required bool IsCourseCompleted { get; set; }
        public required MarkDistribution MarkDistribution { get; set; }
    }
}
