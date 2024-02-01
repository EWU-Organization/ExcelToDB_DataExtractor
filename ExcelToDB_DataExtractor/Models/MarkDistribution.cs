namespace ExcelToDB_DataExtractor.Models
{
    public class MarkDistribution
    {
        public double ClassParticipation { get; set; }
        public double ClassTest { get; set; }
        public required CO Midterm1 { get; set; }
        public required CO Midterm2 { get; set; }
        public required CO Final { get; set; }
        public required CO Project { get; set; }
        public required CO Lab { get; set; }
        public required CO Viva { get; set; }
        public required double GrandTotal { get; set; }
        public required Weight Weights { get; set; }

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
