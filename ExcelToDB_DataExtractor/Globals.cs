using ExcelToDB_DataExtractor.Models;

namespace ExcelToDB_DataExtractor
{
    internal class Globals
    {
        public static List<Student> Students = new List<Student>();

        // Course Details
        public static string CourseID;
        public static string Instructor;
        public static string CourseCode;
        public static string CourseTitle;
        public static string Semester;
        // public static string LetterGrade;
        public static int Section;
        public static double Credit;
        public static bool IsCourseCompleted;

        // Exam details
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
        public static double ClassParticipation;
        public static double ClassTest;
        public static double GrandTotal;

        // MIDTERM 1
        public static List<double> Mid1CO = new List<double>();
        public static double Mid1Total;
        public static double Mid1Converted;

        // MIDTERM 2
        public static List<double> Mid2CO = new List<double>();
        public static double Mid2Total;
        public static double Mid2Converted;

        // FINAL
        public static List<double> FinalCO = new List<double>();
        public static double FinalTotal;
        public static double FinalConverted;

        // PROJECT
        public static List<double> ProjectCO = new List<double>();
        public static double ProjectTotal;
        public static double ProjectConverted;

        // LAB
        public static List<double> LabCO = new List<double>();
        public static double LabTotal;
        public static double LabConverted;

        // VIVA
        public static List<double> VivaCO = new List<double>();
        public static double VivaTotal;
        public static double VivaConverted;

        // CO Contribution
        public static List<double> COContribution = new List<double>();

        //// CO Individual
        //public static List<double> Indiv_CO = new List<double>();

        //// COIndividualBianry
        //public static List<double> Indiv_BIN_CO = new List<double>();

        ////PO INDIVIDUAL
        //public static List<double> Indiv_PO = new List<double>();
    }
}
