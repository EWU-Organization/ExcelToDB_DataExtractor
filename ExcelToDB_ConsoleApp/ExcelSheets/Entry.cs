using ExcelToDB_ConsoleApp.Models;
using ExcelToDB_ConsoleApp.Utils;

namespace ExcelToDB_ConsoleApp.ExcelSheets
{
    internal class Entry
    {
        public static bool IsExtractionComplete = false;
        public static void ExtractInfo(string file)
        {
            IsExtractionComplete = false;
            ExcelUtility.Initialize(file, "Entry");

            ExtractCourseInfo();
            if (string.IsNullOrEmpty(Globals.Instructor) && string.IsNullOrEmpty(Globals.CourseCode) && Globals.Section == 0 && string.IsNullOrEmpty(Globals.CourseTitle) && Globals.Credit == 0 && string.IsNullOrEmpty(Globals.Semester))
            {
                return;
            }
            ExtractStudentInfo();
            //ExtractExamWeightsInfo();

            IsExtractionComplete = true;
        }
        static void ExtractCourseInfo()
        {
            Globals.Instructor = ExcelUtility.CellValue(19, 2);
            Globals.CourseCode = ExcelUtility.CellValue(20, 2);
            Globals.Section = ExcelUtility.ParseInteger(ExcelUtility.CellValue(20, 3));
            Globals.CourseTitle = ExcelUtility.CellValue(21, 2);
            Globals.Credit = ExcelUtility.ParseDouble(ExcelUtility.CellValue(22, 2));
            Globals.Semester = ExcelUtility.CellValue(22, 3);
            Globals.CourseID = $"{Globals.Semester}-{Globals.CourseCode}-{Globals.Section}";
        }
        static void ExtractExamWeightsInfo()
        {
            Globals.ClassParticipationWeight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(7, 17));
            Globals.AssignmentQuizWeight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(8, 17));
            Globals.Midterm1Weight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(9, 17));
            Globals.Midterm2Weight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(10, 17));
            Globals.FinalWeight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(11, 17));
            Globals.ProjectWeight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(12, 17));
            Globals.LabWeight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(13, 17));
            Globals.TotalWeight = ExcelUtility.ParseDouble(ExcelUtility.CellValue(14, 17));
        }
        static void ExtractStudentInfo()
        {
            int rows = ExcelUtility.GetNumberOfRows();
            for (int row = 24; row < rows; row++)
            {
                Globals.ID = ExcelUtility.CellValue(row, 2);
                Globals.Name = ExcelUtility.CellValue(row, 3);
                Globals.ClassParticipation = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, 74));
                Globals.ClassTest = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, 75));
                Globals.GrandTotal = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, 76));
                Globals.LetterGrade = ExcelUtility.CellValue(row, 77);

                ExtractExamInfo(row);
                AddAllInfoToStudentsList();
            }
        }
        static void ExtractExamInfo(int row)
        {
            for (int column = 4; column <= 160; column++)
            {
                double cell = double.Parse(ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, column)).ToString("0.00"));

                if (column <= 15) Globals.Mid1CO.Add(cell);
                else if (column == 16) Globals.Mid1Total = cell;
                else if (column == 17) Globals.Mid1Converted = cell;

                else if (column <= 29) Globals.Mid2CO.Add(cell);
                else if (column == 30) Globals.Mid2Total = cell;
                else if (column == 31) Globals.Mid2Converted = cell;

                else if (column <= 43) Globals.FinalCO.Add(cell);
                else if (column == 44) Globals.FinalTotal = cell;
                else if (column == 45) Globals.FinalConverted = cell;

                else if (column <= 57) Globals.ProjectCO.Add(cell);
                else if (column == 58) Globals.ProjectTotal = cell;
                else if (column == 59) Globals.ProjectConverted = cell;

                else if (column <= 71) Globals.LabCO.Add(cell);
                else if (column == 72) Globals.LabTotal = cell;
                else if (column == 73) Globals.LabConverted = cell;

                else if (column >= 79 && column <= 90) Globals.COContribution.Add(cell);
                //else if (column >= 91 && column <= 102) Globals.Indiv_CO.Add(cell);
                //else if (column >= 103 && column <= 114) Globals.Indiv_BIN_CO.Add(cell);
                //else if (column >= 129 && column <= 140) Globals.Indiv_PO.Add(cell);
            }
        }
        static void AddAllInfoToStudentsList()
        {
            Globals.Students.Add(new Student
            {
                ID = Globals.ID,
                Name = Globals.Name,
                Course = new Course
                {
                    ID = Globals.CourseID, // Need to change this
                    Instructor = Globals.Instructor,
                    CourseCode = Globals.CourseCode,
                    Section = Globals.Section,
                    CourseTitle = Globals.CourseTitle,
                    Credit = Globals.Credit,
                    LetterGrade = Globals.LetterGrade,
                    Semester = Globals.Semester,
                    MarkDistribution = new MarkDistribution
                    {
                        ClassParticipation = Globals.ClassParticipation,
                        ClassTest = Globals.ClassTest,
                        GrandTotal = Globals.GrandTotal,

                        Midterm1 = ExtractCO(Globals.Mid1CO),
                        Midterm1Total = Globals.Mid1Total,
                        Midterm1Converted = Globals.Mid1Converted,

                        Midterm2 = ExtractCO(Globals.Mid2CO),
                        Midterm2Total = Globals.Mid2Total,
                        Midterm2Converted = Globals.Mid2Converted,

                        Final = ExtractCO(Globals.FinalCO),
                        FinalTotal = Globals.FinalTotal,
                        FinalConverted = Globals.FinalConverted,

                        Project = ExtractCO(Globals.ProjectCO),
                        ProjectTotal = Globals.ProjectTotal,
                        ProjectConverted = Globals.ProjectConverted,

                        Lab = ExtractCO(Globals.LabCO),
                        LabTotal = Globals.LabTotal,
                        LabConverted = Globals.LabConverted,

                        Viva = ExtractCO(Globals.VivaCO),
                        VivaTotal = Globals.VivaTotal,
                        VivaConverted = Globals.VivaConverted,

                        COContribution = ExtractCO(Globals.COContribution),
                        //Weights = new MarkDistribution.Weight
                        //{
                        //    Midterm1Weight = Globals.Midterm1Weight,
                        //    Midterm2Weight = Globals.Midterm2Weight,
                        //    FinalWeight = Globals.FinalWeight,
                        //    ProjectWeight = Globals.ProjectWeight,
                        //    LabWeight = Globals.LabWeight,
                        //    VivaWeight = Globals.VivaWeight
                        //}
                    }
                }
            });

            Globals.Mid1CO.Clear();
            Globals.Mid2CO.Clear();
            Globals.FinalCO.Clear();
            Globals.ProjectCO.Clear();
            Globals.LabCO.Clear();
            Globals.VivaCO.Clear();
            Globals.COContribution.Clear();

            //Globals.Indiv_CO.Clear();
            //Globals.Indiv_BIN_CO.Clear();
            //Globals.Indiv_PO.Clear();
        }
        static CourseOutcomes ExtractCO(List<double> co)
        {
            CourseOutcomes result = new CourseOutcomes();

            for (int i = 0; i < co.Count; i++)
            {
                result.CO[$"CO{i + 1}"] = co[i];
            }

            return result;
        }

        public static void Dispose()
        {
            Globals.Students.Clear();
            Globals.Mid1CO.Clear();
            Globals.Mid2CO.Clear();
            Globals.FinalCO.Clear();
            Globals.ProjectCO.Clear();
            Globals.LabCO.Clear();
            Globals.VivaCO.Clear();
            Globals.COContribution.Clear();

            Globals.Instructor = null;
            Globals.CourseCode = null;
            Globals.Section = 0;
            Globals.CourseTitle = null;
            Globals.Credit = 0;
            Globals.Semester = null;
            Globals.CourseID = null;

            Globals.ID = null;
            Globals.Name = null;
            Globals.ClassParticipation = 0;
            Globals.ClassTest = 0;
            Globals.GrandTotal = 0;
            Globals.LetterGrade = null;
        }
    }
}
