using System;
namespace R26_DailyQueueWinForm.Models
{
    public class ReportProcessScheduleQueueModel
    {
        public int ReportScheduleQueueUid { get; set; }
        public int? CompanyUid { get; set; }
        public string CompanyName { get; set; }
        public int? ReportTypeUid { get; set; }
        public string ReportTypeName { get; set; } // NEW PROPERTY
        public int? FrequencyUid { get; set; }
        public DateTime? ScheduleDate { get; set; }
        public TimeSpan? ScheduleTime { get; set; }
        public string ExecutionDuration { get; set; }
        public string RawFilePath { get; set; }
        public string ProcessedFilePath { get; set; }
        public string Timezone { get; set; }
        public string Status { get; set; }
        public DateTime? ReportStartTime { get; set; }
        public DateTime? ReportEndTime { get; set; }
        public string SLAComplianceFlag { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
    }
}