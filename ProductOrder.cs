using System.ComponentModel.DataAnnotations;

namespace iSchedule.Data
{
    public class ProjectOrder
    {
        [Key]
        public string ProjectOrderId { get; set; }
        public string ProjectId { get; set; }
        public string ShipId { get; set; }
        public string ProductType { get; set; }
        public string ProductScale { get; set; }
        public string Unit { get; set; }
        public string GraphId { get; set; }
        public string ExtraInfo { get; set; }
        public int TotalAmount { get; set; }
        public int BackupAmount { get; set; }
        public int Cabin1Amount { get; set; }
        public int Cabin2Amount { get; set; }
        public int Cabin3Amount { get; set; }
        public int Cabin4Amount { get; set; }
        public int Cabin5Amount { get; set; }
        public int Cabin6Amount { get; set; }
        public int Cabin7Amount { get; set; }
        public int Cabin8Amount { get; set; }
    }
}
