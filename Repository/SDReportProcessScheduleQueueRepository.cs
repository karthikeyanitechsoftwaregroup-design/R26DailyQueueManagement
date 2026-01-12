using R26_DailyQueueWinForm.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace R26_DailyQueueWinForm.Data
{
    public class SDReportProcessScheduleQueueRepository
    {
        private readonly string _connectionString;

        public SDReportProcessScheduleQueueRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public List<SDReportProcessScheduleQueueModel> GetAllSDReportProcessScheduleQueue()
        {
            var queueList = new List<SDReportProcessScheduleQueueModel>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("kar_GetAllSDReportProcessScheduleQueue", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 120;

                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var queue = new SDReportProcessScheduleQueueModel
                            {
                                SdReportScheduleQueueUid = reader.GetInt32(reader.GetOrdinal("sdreportschedulequeueuid")),
                                CompanyUid = reader.IsDBNull(reader.GetOrdinal("companyuid")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("companyuid")),
                                CompanyName = reader.IsDBNull(reader.GetOrdinal("companyname")) ? null : reader.GetString(reader.GetOrdinal("companyname")),
                                SdRawReportUid = reader.IsDBNull(reader.GetOrdinal("sdrawreportuid")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("sdrawreportuid")),
                                ScheduleDate = reader.IsDBNull(reader.GetOrdinal("scheduledate")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("scheduledate")),
                                ScheduleTime = reader.IsDBNull(reader.GetOrdinal("scheduletime")) ? (TimeSpan?)null : reader.GetTimeSpan(reader.GetOrdinal("scheduletime")),
                                ExecutionDuration = reader.IsDBNull(reader.GetOrdinal("executionduration")) ? null : reader.GetString(reader.GetOrdinal("executionduration")),
                                RawFilePath = reader.IsDBNull(reader.GetOrdinal("rawfilepath")) ? null : reader.GetString(reader.GetOrdinal("rawfilepath")),
                                ProcessedFilePath = reader.IsDBNull(reader.GetOrdinal("processedfilepath")) ? null : reader.GetString(reader.GetOrdinal("processedfilepath")),
                                Status = reader.IsDBNull(reader.GetOrdinal("status")) ? null : reader.GetString(reader.GetOrdinal("status")),
                                ReportStartTime = reader.IsDBNull(reader.GetOrdinal("ReportStartTime")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("ReportStartTime")),
                                ReportEndTime = reader.IsDBNull(reader.GetOrdinal("ReportEndTime")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("ReportEndTime")),
                                SLAComplianceFlag = reader.IsDBNull(reader.GetOrdinal("SLAComplianceFlag")) ? null : reader.GetString(reader.GetOrdinal("SLAComplianceFlag")),
                                CreatedDate = reader.IsDBNull(reader.GetOrdinal("createddate")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("createddate")),
                                CreatedBy = reader.IsDBNull(reader.GetOrdinal("createdby")) ? null : reader.GetString(reader.GetOrdinal("createdby")),
                                ModifiedDate = reader.IsDBNull(reader.GetOrdinal("modifieddate")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("modifieddate")),
                                ModifiedBy = reader.IsDBNull(reader.GetOrdinal("modifiedby")) ? null : reader.GetString(reader.GetOrdinal("modifiedby"))
                            };

                            queueList.Add(queue);
                        }
                    }
                }
            }

            return queueList;
        }

        public int UpdateStatuses(Dictionary<int, string> updates, string systemName)
        {
            int updatedCount = 0;

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                foreach (var update in updates)
                {
                    using (SqlCommand cmd = new SqlCommand("kar_UpdateSDReportProcessScheduleQueue", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@sdreportschedulequeueuid", update.Key);
                        cmd.Parameters.AddWithValue("@status", update.Value);

                        cmd.ExecuteNonQuery();
                        updatedCount++;
                    }
                }
            }

            return updatedCount;
        }

        public List<string> GetAllStatuses()
        {
            var statuses = new List<string>();
            var customStatuses = new HashSet<string>();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    con.Open();

                    // Correct table name: SDreportprocesschedulequeue (no 'h' in schedule)
                    using (SqlCommand cmd = new SqlCommand(
                               "SELECT DISTINCT status FROM SDreportprocesschedulequeue WHERE status IS NOT NULL AND status != ''",
                               con))
                    {
                        using (SqlDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                string status = rdr["status"].ToString();
                                if (!string.IsNullOrWhiteSpace(status))
                                {
                                    customStatuses.Add(status);
                                }
                            }
                        }
                    }
                }

                // Add statuses ordered alphabetically
                statuses.AddRange(customStatuses.OrderBy(s => s));
            }
            catch (Exception ex)
            {
                // On error, just rethrow and let UI handle message
                throw new Exception($"Error retrieving statuses: {ex.Message}", ex);
            }

            return statuses;
        }
    }
}