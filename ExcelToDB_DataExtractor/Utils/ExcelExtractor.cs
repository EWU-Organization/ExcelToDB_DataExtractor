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
                string cell = ExcelUtility.CellValue(row, i);
                if (i <= 15)
                {
                    Globals.Mid1_CO.Add(ExcelUtility.ParseDouble(cell));
                }
                else if (i == 16)
                {
                    Globals.Mid1_Total = ExcelUtility.ParseDouble(cell);
                }
                else if (i == 17)
                {
                    Globals.Mid1_Converted = ExcelUtility.ParseDouble(cell);
                }
                else if (i <= 29)
                {
                    Globals.Mid2_CO.Add(ExcelUtility.ParseDouble(cell));
                }
                else if (i == 30)
                {
                    Globals.Mid2_Total = ExcelUtility.ParseDouble(cell);
                }
                else if (i == 31)
                {
                    Globals.Mid2_Converted = ExcelUtility.ParseDouble(cell);
                }
                else if (i <= 43)
                {
                    Globals.Final_CO.Add(ExcelUtility.ParseDouble(cell));
                }
                else if (i == 44)
                {
                    Globals.Final_Total = ExcelUtility.ParseDouble(cell);
                }
                else if (i == 45)
                {
                    Globals.Final_Converted = ExcelUtility.ParseDouble(cell);
                }
                else if (i <= 57)
                {
                    Globals.Project_CO.Add(ExcelUtility.ParseDouble(cell));
                }
                else if (i == 58)
                {
                    Globals.Project_Total = ExcelUtility.ParseDouble(cell);
                }
                else if (i == 59)
                {
                    Globals.Project_Converted = ExcelUtility.ParseDouble(cell);
                }
                else if (i <= 71)
                {
                    Globals.Lab_CO.Add(ExcelUtility.ParseDouble(cell));
                }
                else if (i == 72)
                {
                    Globals.Lab_Converted = ExcelUtility.ParseDouble(cell);
                }
                else if (i == 73)
                {
                    Globals.Lab_Total = ExcelUtility.ParseDouble(cell);
                }
                else if (i >= 91 && i <= 102)
                {
                    Globals.Indiv_CO.Add(ExcelUtility.ParseDouble(cell));
                }
                else if (i >= 103 && i <= 114)
                {
                    Globals.Indiv_BIN_CO.Add(ExcelUtility.ParseDouble(cell));
                }

                else if (i >= 129 && i <= 140)
                {
                    Globals.Indiv_PO.Add(ExcelUtility.ParseDouble(cell));
                }
            }
        }
    }
}
