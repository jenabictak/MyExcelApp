namespace MyExcelApp.Models
{
    public class FormData
    {
        // اطلاعات کلی (خودکار)
        public string? Date { get; set; }
        public string? Time { get; set; }
        public string? DayOfWeek { get; set; }
        
        // اطلاعات کلی (دستی)
        public string? ResponsibleCook { get; set; }
        public string? TechnicalExpert { get; set; }
        public string? ProductType { get; set; }
        public int? PersonnelCount { get; set; }

        // مقدار مواد اولیه (همه به گرم)
        public decimal? CheeseFrozen { get; set; }
        public decimal? CheeseDefrost { get; set; }
        public decimal? Oil { get; set; }
        public decimal? Starch { get; set; }
        public decimal? Ricotta { get; set; }
        public decimal? Phosphate { get; set; }
        public decimal? Water { get; set; }
        public decimal? Stabilizer { get; set; }
        public decimal? Citrate { get; set; }
        public decimal? Salt { get; set; }
        public decimal? WasteReturn { get; set; }
        public decimal? Essence { get; set; }
        public decimal? CheeseGolpayegan { get; set; }
        public decimal? Sorbate { get; set; }
        public decimal? PowderCode24 { get; set; }
        public decimal? PowderCodeF { get; set; }
        public decimal? ColorAnnatto { get; set; }
        public decimal? PlasticUnderLayer { get; set; }

        // مواد اولیه متفرقه با نام و مقدار
        public string? RawMaterial1Name { get; set; }
        public decimal? RawMaterial1 { get; set; }
        public string? RawMaterial2Name { get; set; }
        public decimal? RawMaterial2 { get; set; }
        public string? RawMaterial3Name { get; set; }
        public decimal? RawMaterial3 { get; set; }

        // اطلاعات نهایی
        public int? MoldCount { get; set; }
        public decimal? MoldWeight { get; set; }
        public decimal? TotalMoldWeight { get; set; }
        public decimal? TotalInputWeight { get; set; }
        public decimal? ActualProductWeight { get; set; }

        // توضیحات
        public string? ProductionDescription { get; set; }
        public string? CookDescription { get; set; }
        public string? ExpertDescription { get; set; }
    }
}