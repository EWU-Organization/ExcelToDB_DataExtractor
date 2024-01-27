namespace ExcelToDB_DataExtractor.Models
{
    public class Course
    {
        public required string Semester { get; set; }
        public required string CourseCode { get; set; }
        public required string Section { get; set; }
        public required string Instructor { get; set; }
        public required string CourseTitile { get; set; }
        public required string Credit { get; set; }
        public required string LetterGrade { get; set; }
        public required bool IsCourseCompleted { get; set; }
    }
}
