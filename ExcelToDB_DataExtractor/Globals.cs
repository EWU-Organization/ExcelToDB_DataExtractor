using ExcelToDB_DataExtractor.Models;

namespace ExcelToDB_DataExtractor
{
    internal class Globals
    {
        public static List<Student> Students = new List<Student>();
        // Course Details
        public static string Instructor;
        public static string CourseCode;
        public static string CourseTitle;
        public static string Semester;
        // public static string LetterGrade;
        public static int Section;
        public static int Credit;
        public static bool IsCourseCompleted;

        //Exam details
        public static double Midterm1Weight;
        public static double Midterm2Weight;
        public static double FinalWeight;
        public static double ProjectWeight;
        public static double LabWeight;
        public static double VivaWeight;
        public static double AssignmentQuizWeight;
        public static double ClassParticipationWeight;
        public static double TotalWeight;

        // Student Details
        public static string ID;
        public static string Name;
        public static string LetterGrade;
        public static double SI_NO;
        public static double ClassParticipation;
        public static double ClassTest;
        public static double GrandTotal;

        // MIDTERM 1
        public static List<double> Mid1_CO = new List<double>();
        public static double Mid1_Total;
        public static double Mid1_Converted;

        // MIDTERM 2
        public static List<double> Mid2_CO = new List<double>();
        public static double Mid2_Total;
        public static double Mid2_Converted;

        // FINAL
        public static List<double> Final_CO = new List<double>();
        public static double Final_Total;
        public static double Final_Converted;

        // PROJECT
        public static List<double> Project_CO = new List<double>();
        public static double Project_Total;
        public static double Project_Converted;

        // LAB
        public static List<double> Lab_CO = new List<double>();
        public static double Lab_Total;
        public static double Lab_Converted;

        // CO Individual
        public static List<double> Indiv_CO = new List<double>();

        // COIndividualBianry
        public static List<double> Indiv_BIN_CO = new List<double>();

        //PO INDIVIDUAL
        public static List<double> Indiv_PO = new List<double>();
    }
}
