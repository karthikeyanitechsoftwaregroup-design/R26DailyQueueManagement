using System;

namespace R26_DailyQueueWinForm.Models
{
    public class R26QueueModel
    {
        public int R26QueueUid { get; set; }
        public int CompanyUid { get; set; }
        public string CompanyName { get; set; } = "";
        public int? PatNumber { get; set; }
        public string PatFName { get; set; } = "";
        public string PatMInitial { get; set; } = "";
        public string PatLName { get; set; } = "";
        public string PatSex { get; set; } = "";
        public DateTime? PatBirthdate { get; set; }
        public int? LocationNumber { get; set; }
        public DateTime? Date { get; set; }
        public TimeSpan? Time { get; set; }
        public string ApptType { get; set; } = "";
        public string Reason { get; set; } = "";
        public string PatAddress1 { get; set; } = "";
        public string PatAddress2 { get; set; } = "";
        public string PatCity { get; set; } = "";
        public string PatState { get; set; } = "";
        public string PatZip5 { get; set; } = "";
        public string HomePhone { get; set; } = "";
        public string WorkPhone { get; set; } = "";
        public int? TicketNumber { get; set; }
        public string UserField1 { get; set; } = "";
        public string UserField2 { get; set; } = "";
        public string UserField3 { get; set; } = "";
        public string UserField4 { get; set; } = "";
        public string UserField5 { get; set; } = "";
        public int? RDrNumber { get; set; }
        public int? DictationID { get; set; }
        public string AdmitDate { get; set; } = "";
        public string DischargeDate { get; set; } = "";
        public DateTime? TwoDayAgo { get; set; }
        public DateTime? Tomorrow { get; set; }
        public DateTime? TwoDaysAgoDateOnly { get; set; }
        public DateTime? TomorrowDateOnly { get; set; }
        public string ResourceName { get; set; } = "";
        public string ProviderName { get; set; } = "";
        public string PrimaryInsuranceName { get; set; } = "";
        public string PrimaryInsSubscriberNo { get; set; } = "";
        public string PrimaryInsGroupNo { get; set; } = "";
        public string PrimaryInsCopays { get; set; } = "";
        public string PatientBalance { get; set; } = "";
        public string AccountBalance { get; set; } = "";
        public string BotDetails { get; set; } = "";
        public string BotName { get; set; } = "";
        public string Webhook { get; set; } = "";
        public DateTime? CreatedDate { get; set; }
        public string Status { get; set; } = "";
        public DateTime? ModifiedDate { get; set; }
    }

}