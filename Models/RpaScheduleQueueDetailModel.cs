using System;

namespace R26_DailyQueueWinForm.Models
{
    public class RpaScheduleQueueDetailModel
    {
        public int RpaScheduleQueueDetailUid { get; set; }
        public int RpaScheduleQueueUid { get; set; }
        public int CompanyUid { get; set; }
        public string CompanyName { get; set; }
        public int? RawReportUid { get; set; }
        public string RawReportName { get; set; }
        public string RawFilePath { get; set; }
        public string Status { get; set; }
        public DateTime? BotStartTime { get; set; }
        public DateTime? BotEndTime { get; set; }
        public string Duration { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
    }
}