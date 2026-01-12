using R26_DailyQueueWinForm.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace R26_DailyQueueWinForm.Data
{
    public class RpaScheduleQueueDetailRepository
    {
        private readonly string _connectionString;

        public RpaScheduleQueueDetailRepository(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(connectionString));
            }

            try
            {
                var builder = new SqlConnectionStringBuilder(connectionString);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid connection string format: {ex.Message}", nameof(connectionString), ex);
            }

            _connectionString = connectionString;
        }

        public List<RpaScheduleQueueDetailModel> GetAllRpaScheduleQueueDetail()
        {
            var list = new List<RpaScheduleQueueDetailModel>();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                using (SqlCommand cmd = new SqlCommand("kar_GetAllRpaScheduleQueueDetail", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 120;

                    con.Open();

                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(new RpaScheduleQueueDetailModel
                            {
                                RpaScheduleQueueDetailUid = SafeGetInt(rdr, "rpaschedulequeuedetailuid"),
                                RpaScheduleQueueUid = SafeGetInt(rdr, "rpaschedulequeueuid"),
                                CompanyUid = SafeGetInt(rdr, "companyuid"),
                                CompanyName = SafeGetString(rdr, "companyname") ?? "Company " + SafeGetInt(rdr, "companyuid"),
                                RawReportUid = SafeGetNullableInt(rdr, "rawreportuid"),
                                RawReportName = SafeGetString(rdr, "rawreportname"),
                                RawFilePath = SafeGetString(rdr, "rawfilepath"),
                                Status = SafeGetString(rdr, "status"),
                                BotStartTime = SafeGetNullableDateTime(rdr, "botstarttime"),
                                BotEndTime = SafeGetNullableDateTime(rdr, "botendtime"),
                                Duration = SafeGetString(rdr, "duration"),
                                CreatedDate = SafeGetNullableDateTime(rdr, "createddate"),
                                CreatedBy = SafeGetString(rdr, "createdby"),
                                ModifiedDate = SafeGetNullableDateTime(rdr, "modifieddate"),
                                ModifiedBy = SafeGetString(rdr, "modifiedby")
                            });
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                throw new Exception($"Database error while retrieving RPA queue: {sqlEx.Message}\nProcedure: kar_GetAllRpaScheduleQueueDetail", sqlEx);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving RPA queue: {ex.Message}", ex);
            }

            return list;
        }

        private int SafeGetInt(SqlDataReader rdr, string columnName)
        {
            try
            {
                return rdr[columnName] == DBNull.Value ? 0 : Convert.ToInt32(rdr[columnName]);
            }
            catch
            {
                return 0;
            }
        }

        private int? SafeGetNullableInt(SqlDataReader rdr, string columnName)
        {
            try
            {
                return rdr[columnName] == DBNull.Value ? (int?)null : Convert.ToInt32(rdr[columnName]);
            }
            catch
            {
                return null;
            }
        }

        private string SafeGetString(SqlDataReader rdr, string columnName)
        {
            try
            {
                return rdr[columnName] == DBNull.Value ? null : rdr[columnName].ToString();
            }
            catch
            {
                return null;
            }
        }

        private DateTime? SafeGetNullableDateTime(SqlDataReader rdr, string columnName)
        {
            try
            {
                return rdr[columnName] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(rdr[columnName]);
            }
            catch
            {
                return null;
            }
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

                    using (SqlCommand cmd = new SqlCommand(
                               "SELECT DISTINCT status FROM rpaschedulequeuedetail WHERE status IS NOT NULL AND status != ''",
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

                statuses.AddRange(customStatuses.OrderBy(s => s));
            }
            catch (Exception ex)
            {
                // On error, just rethrow and let UI handle message; no hardcoded statuses.
                throw new Exception($"Error retrieving statuses: {ex.Message}", ex);
            }

            return statuses;
        }


        public int UpdateStatuses(Dictionary<int, string> updates, string systemName)
        {
            int count = 0;
            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    try
                    {
                        foreach (var item in updates)
                        {
                            using (SqlCommand cmd = new SqlCommand("kar_UpdateRpaScheduleQueueDetail", con, transaction))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.AddWithValue("@rpaschedulequeuedetailuid", item.Key);
                                cmd.Parameters.AddWithValue("@status", item.Value);
                                count += cmd.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
            return count;
        }
    }
}