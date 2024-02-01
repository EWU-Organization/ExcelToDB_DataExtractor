using ExcelToDB_DataExtractor.Models;
using System.Windows.Input;

namespace ExcelToDB_DataExtractor.Utils
{
    internal static class ExcelExtractor
    {
        public static void ProcessExcelFiles(string[] directoryPath)
        {
            Task.Run(() =>
            {
                foreach (var file in directoryPath)
                {
                    ProcessExcelFile(file);
                    ExcelUtility.Dispose();
                }
            });
            
            //try
            //{

            //}
            //catch (Exception ex)
            //{
            //    Console.Error.WriteLine(ex.Message);
            //}
        }

        static void ProcessExcelFile(string file)
        {
            ExcelUtility.Initialize(file, "Entry");

            ExtractCourseInfo();
            ExtractStudentInfo();
            ExtractExamWeightsInfo();
        }

        static void ExtractCourseInfo()
        {
            Globals.Instructor = ExcelUtility.CellValue(19, 2);
            Globals.CourseCode = ExcelUtility.CellValue(20, 2);
            Globals.Section = ExcelUtility.ParseInteger(ExcelUtility.CellValue(20, 3));
            Globals.CourseTitle = ExcelUtility.CellValue(21, 2);
            Globals.Credit = ExcelUtility.ParseInteger(ExcelUtility.CellValue(22, 2));
            Globals.Semester = ExcelUtility.CellValue(22, 3);
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
                Globals.Name = ExcelUtility.CellValue(row, 3);
                Globals.ClassParticipation = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, 74));
                Globals.ClassTest = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, 75));
                Globals.GrandTotal = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, 76));
                Globals.LetterGrade = ExcelUtility.CellValue(row, 77);

                ExtractExamInfo(row);
            }
        }
        static void ExtractExamInfo(int row)
        {
            for (int i = 4; i <= 160; i++)
            {
                double cell = ExcelUtility.ParseDouble(ExcelUtility.CellValue(row, i));
                if (i <= 15) Globals.Mid1_CO.Add(cell);
                else if (i == 16) Globals.Mid1_Total = cell;
                else if (i == 17) Globals.Mid1_Converted = cell;
                else if (i <= 29) Globals.Mid2_CO.Add(cell);
                else if (i == 30) Globals.Mid2_CO.Add(cell);
                else if (i == 31) Globals.Mid2_Converted = cell;
                else if (i <= 43) Globals.Final_CO.Add(cell);
                else if (i == 44) Globals.Final_Total = cell;
                else if (i == 45) Globals.Final_Converted = cell;
                else if (i <= 57) Globals.Project_CO.Add(cell);
                else if (i == 58) Globals.Project_Total = cell;
                else if (i == 59) Globals.Project_Converted = cell;
                else if (i <= 71) Globals.Lab_CO.Add(cell);
                else if (i == 72) Globals.Lab_Converted = cell;
                else if (i == 73) Globals.Lab_Total = cell;
                else if (i >= 91 && i <= 102) Globals.Indiv_CO.Add(cell);
                else if (i >= 103 && i <= 114) Globals.Indiv_BIN_CO.Add(cell);
                else if (i >= 129 && i <= 140) Globals.Indiv_PO.Add(cell);
            }
        }
        static void StudentJSON(string id, string name, Course course)
        {
            Student student = new Student()
            {
                ID = id,
                Name = name,
                Course = new List<Course>()
            };
            student.Course.Add(course); // Needs to be inside a loop
        }

        static Course Course(string id, string courseCode, string courseTitle, double credit, string instructor, bool isCourseCompleted, string letterGrade, int section, string semester, MarkDistribution markDistribution)
        {
            return new Course
            {
                ID = id,
                CourseCode = courseCode,
                CourseTitle = courseTitle,
                Credit = credit,
                Instructor = instructor,
                IsCourseCompleted = isCourseCompleted,
                LetterGrade = letterGrade,
                Section = section,
                Semester = semester,
                MarkDistribution = markDistribution
            };
        }
        static MarkDistribution MarkDistribution(double classParticipation, double classTest, double grandTotal, CO midterm1, CO midterm2, CO final, CO project, CO lab, CO viva, MarkDistribution.Weight weights)
        {
            return new MarkDistribution()
            {
                ClassParticipation = classParticipation,
                ClassTest = classTest,
                Midterm1 = midterm1,
                Midterm2 = midterm2,
                Final = final,
                Project = project,
                Lab = lab,
                Viva = viva,
                GrandTotal = grandTotal,
                Weights = weights
            };
        }
        static MarkDistribution.Weight MarkDistributionWeight(double midterm1Weight, double midterm2Weight, double finalWeight, double projectWeight, double labWeight, double vivaWeight)
        {
            return new MarkDistribution.Weight()
            {
                Midterm1Weight = midterm1Weight,
                Midterm2Weight = midterm2Weight,
                FinalWeight = finalWeight,
                ProjectWeight = projectWeight,
                LabWeight = labWeight,
                VivaWeight = vivaWeight,
            };
        }
        //static void AddAllInfoToStudentsList()
        //{
        //    Term mid1 = new Term();
        //    Term mid2 = new Term();
        //    Term final = new Term();
        //    Term project = new Term();
        //    Term lab = new Term();
        //    CO_PO co_Indiv = new CO_PO();
        //    CO_PO co_Indiv_BIN = new CO_PO();
        //    CO_PO po_Indiv = new CO_PO();

        //    for (int i = 0; i < 12; i++)
        //    {
        //        mid1.GetType().GetProperty($"CO_PO{i + 1}").SetValue(mid1, Globals.Mid1_CO[i]);
        //        mid2.GetType().GetProperty($"CO_PO{i + 1}").SetValue(mid2, Globals.Mid2_CO[i]);
        //        final.GetType().GetProperty($"CO_PO{i + 1}").SetValue(final, Globals.Final_CO[i]);
        //        project.GetType().GetProperty($"CO_PO{i + 1}").SetValue(project, Globals.Project_CO[i]);
        //        lab.GetType().GetProperty($"CO_PO{i + 1}").SetValue(lab, Globals.Lab_CO[i]);
        //        co_Indiv.GetType().GetProperty($"CO_PO{i + 1}").SetValue(co_Indiv, Globals.Indiv_CO[i]);
        //        co_Indiv.GetType().GetProperty($"CO_PO{i + 1}").SetValue(co_Indiv, Globals.Indiv_CO[i]);
        //        co_Indiv_BIN.GetType().GetProperty($"CO_PO{i + 1}").SetValue(co_Indiv_BIN, Globals.Indiv_BIN_CO[i]);
        //        po_Indiv.GetType().GetProperty($"CO_PO{i + 1}").SetValue(po_Indiv, Globals.Indiv_PO[i]);
        //    }

        //    //Midterm 1
        //    mid1.Total = Globals.Mid1_Total;
        //    mid1.Converted = Globals.Mid1_Converted;
        //    //Midterm 2
        //    mid2.Total = Globals.Mid2_Total;
        //    mid2.Converted = Globals.Mid2_Converted;
        //    // Final
        //    final.Total = Globals.Final_Total;
        //    final.Converted = Globals.Final_Converted;
        //    // Project
        //    project.Total = Globals.Project_Total;
        //    project.Converted = Globals.Project_Converted;
        //    // Lab
        //    lab.Total = Globals.Lab_Total;
        //    lab.Converted = Globals.Lab_Converted;

        //    Globals.Students.Add(new Student
        //    {
        //        Instructor = Globals.Instructor,
        //        CourseCode = Globals.CourseCode,
        //        Section = Globals.Section,
        //        CourseTitile = Globals.CourseTitle,
        //        Credit = Globals.Credit,
        //        Semester = Globals.Semester,
        //        SI_NO = Globals.SI_NO,
        //        StudentID = Globals.StudentID,
        //        Name = Globals.Name,
        //        ClassParticipation = Globals.ClassParticipation,
        //        ClassTest = Globals.ClassTest,
        //        LetterGrade = Globals.LetterGrade,
        //        Midterm1 = mid1,
        //        Midterm2 = mid2,
        //        Final = final,
        //        Project = project,
        //        Lab = lab,
        //        GrandTotal = Globals.GrandTotal,
        //        COIndividual = co_Indiv,
        //        COIndividualBianry = co_Indiv_BIN,
        //        POIndividual = po_Indiv
        //    });

        //    Globals.Mid1_CO.Clear();
        //    Globals.Mid2_CO.Clear();
        //    Globals.Final_CO.Clear();
        //    Globals.Project_CO.Clear();
        //    Globals.Lab_CO.Clear();
        //    Globals.Indiv_CO.Clear();
        //    Globals.Indiv_BIN_CO.Clear();
        //    Globals.Indiv_PO.Clear();
        //}
    }
}
