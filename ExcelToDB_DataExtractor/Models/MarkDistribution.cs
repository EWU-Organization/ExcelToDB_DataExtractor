namespace ExcelToDB_DataExtractor.Models
{
    public class MarkDistribution
    {
        public double ClassParticipation { get; set; }
        public double ClassTest { get; set; }

        public CourseOutcomes Midterm1 { get; set; }
        public double Midterm1Total { get; set; }
        public double Midterm1Converted { get; set; }
        public CourseOutcomes Midterm2 { get; set; }
        public double Midterm2Total { get; set; }
        public double Midterm2Converted { get; set; }
        public CourseOutcomes Final { get; set; }
        public double FinalTotal { get; set; }
        public double FinalConverted { get; set; }
        public CourseOutcomes Project { get; set; }
        public double ProjectTotal { get; set; }
        public double ProjectConverted { get; set; }
        public CourseOutcomes Lab { get; set; }
        public double LabTotal { get; set; }
        public double LabConverted { get; set; }
        public CourseOutcomes Viva { get; set; }
        public double VivaTotal { get; set; }
        public double VivaConverted { get; set; }
        public CourseOutcomes COContribution { get; set; }
        
        public required double GrandTotal { get; set; }
        //public required Weight Weights { get; set; }

        public class Weight
        {
            public double Midterm1Weight { get; set; }
            public double Midterm2Weight { get; set; }
            public double FinalWeight { get; set; }
            public double ProjectWeight { get; set; }
            public double LabWeight { get; set; }
            public double VivaWeight { get; set; }
        }
    }
}
