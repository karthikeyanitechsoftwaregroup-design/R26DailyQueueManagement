using R26_DailyQueueWinForm.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace R26_DailyQueueWinForm.Data
{
    public class R26QueueRepository
    {
        private readonly string _connectionString;

        public R26QueueRepository(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(connectionString));
            }

            // Validate the connection string format
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

        public List<R26QueueModel> GetAllR26Queue()
        {
            var list = new List<R26QueueModel>();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                using (SqlCommand cmd = new SqlCommand("GetAllR26DailyQueue", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    con.Open();

                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(new R26QueueModel
                            {
                                R26QueueUid = (int)rdr["r26queueuid"],
                                CompanyUid = (int)rdr["companyuid"],
                                CompanyName = rdr["companyname"] == DBNull.Value ? "Company " + rdr["companyuid"] : rdr["companyname"].ToString(),
                                PatNumber = rdr["PatNumber"] == DBNull.Value ? null : (int?)rdr["PatNumber"],
                                PatFName = rdr["PatFName"] == DBNull.Value ? null : rdr["PatFName"].ToString(),
                                PatMInitial = rdr["PatMInitial"] == DBNull.Value ? null : rdr["PatMInitial"].ToString(),
                                PatLName = rdr["PatLName"] == DBNull.Value ? null : rdr["PatLName"].ToString(),
                                PatSex = rdr["PatSex"] == DBNull.Value ? null : rdr["PatSex"].ToString(),
                                PatBirthdate = rdr["PatBirthdate"] == DBNull.Value ? null : (DateTime?)rdr["PatBirthdate"],
                                LocationNumber = rdr["LocationNumber"] == DBNull.Value ? null : (int?)rdr["LocationNumber"],
                                Date = rdr["Date"] == DBNull.Value ? null : (DateTime?)rdr["Date"],
                                Time = rdr["Time"] == DBNull.Value ? null : (TimeSpan?)rdr["Time"],
                                ApptType = rdr["ApptType"] == DBNull.Value ? null : rdr["ApptType"].ToString(),
                                Reason = rdr["Reason"] == DBNull.Value ? null : rdr["Reason"].ToString(),
                                PatAddress1 = rdr["PatAddress1"] == DBNull.Value ? null : rdr["PatAddress1"].ToString(),
                                PatCity = rdr["PatCity"] == DBNull.Value ? null : rdr["PatCity"].ToString(),
                                PatState = rdr["PatState"] == DBNull.Value ? null : rdr["PatState"].ToString(),
                                PatZip5 = rdr["PatZip5"] == DBNull.Value ? null : rdr["PatZip5"].ToString(),
                                HomePhone = rdr["HomePhone"] == DBNull.Value ? null : rdr["HomePhone"].ToString(),
                                WorkPhone = rdr["WorkPhone"] == DBNull.Value ? null : rdr["WorkPhone"].ToString(),
                                ResourceName = rdr["ResourceName"] == DBNull.Value ? null : rdr["ResourceName"].ToString(),
                                ProviderName = rdr["ProviderName"] == DBNull.Value ? null : rdr["ProviderName"].ToString(),
                                PrimaryInsuranceName = rdr["PrimaryInsuranceName"] == DBNull.Value ? null : rdr["PrimaryInsuranceName"].ToString(),
                                BotName = rdr["botname"] == DBNull.Value ? null : rdr["botname"].ToString(),
                                Status = rdr["status"] == DBNull.Value ? null : rdr["status"].ToString(),
                                CreatedDate = rdr["createdate"] == DBNull.Value ? null : (DateTime?)rdr["createdate"],
                                ModifiedDate = rdr["modifieddate"] == DBNull.Value ? null : (DateTime?)rdr["modifieddate"]
                            });
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                throw new Exception($"Database error while retrieving R26 queue: {sqlEx.Message}", sqlEx);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving R26 queue: {ex.Message}", ex);
            }

            return list;
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

                    using (SqlCommand cmd = new SqlCommand("SELECT DISTINCT status FROM R26_DailyQueue WHERE status IS NOT NULL AND status != ''", con))
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

                // If no statuses found in DB, provide defaults
                if (customStatuses.Count == 0)
                {
                    statuses.AddRange(new[] { "Pending", "Failed", "Processed", "No Data" });
                }
                else
                {
                    // Add all statuses from DB in alphabetical order
                    statuses.AddRange(customStatuses.OrderBy(s => s));
                }
            }
            catch (SqlException sqlEx)
            {
                // If error occurs, return default statuses
                statuses.Clear();
                statuses.AddRange(new[] { "Pending", "Failed", "Processed", "No Data" });
                throw new Exception($"SQL error retrieving statuses (using defaults): {sqlEx.Message}", sqlEx);
            }
            catch (Exception ex)
            {
                // If error occurs, return default statuses
                statuses.Clear();
                statuses.AddRange(new[] { "Pending", "Failed", "Processed", "No Data" });
                throw new Exception($"Error retrieving statuses (using defaults): {ex.Message}", ex);
            }

            return statuses;
        }

        public List<R26QueueModel> GetPagedR26Queue(int pageNumber, int pageSize, out int totalRecords)
        {
            var list = new List<R26QueueModel>();
            totalRecords = 0;

            using (SqlConnection con = new SqlConnection(_connectionString))
            using (SqlCommand cmd = new SqlCommand("GetR26DailyQueuePaginated", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PageNumber", pageNumber);
                cmd.Parameters.AddWithValue("@PageSize", pageSize);
                cmd.CommandTimeout = 60;

                con.Open();
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    // Result Set 1: Total Count
                    if (rdr.Read())
                    {
                        totalRecords = Convert.ToInt32(rdr["TotalRecords"]);
                    }

                    // Result Set 2: The Data
                    if (rdr.NextResult())
                    {
                        while (rdr.Read())
                        {
                            string rawStatus = rdr["status"] == DBNull.Value ? "" : rdr["status"].ToString();

                            list.Add(new R26QueueModel
                            {
                                R26QueueUid = (int)rdr["r26queueuid"],
                                CompanyUid = (int)rdr["companyuid"],
                                PatNumber = rdr["PatNumber"] == DBNull.Value ? null : (int?)rdr["PatNumber"],
                                PatFName = rdr["PatFName"] == DBNull.Value ? null : rdr["PatFName"].ToString(),
                                PatMInitial = rdr["PatMInitial"] == DBNull.Value ? null : rdr["PatMInitial"].ToString(),
                                PatLName = rdr["PatLName"] == DBNull.Value ? null : rdr["PatLName"].ToString(),
                                PatSex = rdr["PatSex"] == DBNull.Value ? null : rdr["PatSex"].ToString(),
                                PatBirthdate = rdr["PatBirthdate"] == DBNull.Value ? null : (DateTime?)rdr["PatBirthdate"],
                                LocationNumber = rdr["LocationNumber"] == DBNull.Value ? null : (int?)rdr["LocationNumber"],
                                Date = rdr["Date"] == DBNull.Value ? null : (DateTime?)rdr["Date"],
                                Time = rdr["Time"] == DBNull.Value ? null : (TimeSpan?)rdr["Time"],
                                ApptType = rdr["ApptType"] == DBNull.Value ? null : rdr["ApptType"].ToString(),
                                Reason = rdr["Reason"] == DBNull.Value ? null : rdr["Reason"].ToString(),
                                PatAddress1 = rdr["PatAddress1"] == DBNull.Value ? null : rdr["PatAddress1"].ToString(),
                                PatCity = rdr["PatCity"] == DBNull.Value ? null : rdr["PatCity"].ToString(),
                                PatState = rdr["PatState"] == DBNull.Value ? null : rdr["PatState"].ToString(),
                                PatZip5 = rdr["PatZip5"] == DBNull.Value ? null : rdr["PatZip5"].ToString(),
                                HomePhone = rdr["HomePhone"] == DBNull.Value ? null : rdr["HomePhone"].ToString(),
                                WorkPhone = rdr["WorkPhone"] == DBNull.Value ? null : rdr["WorkPhone"].ToString(),
                                ResourceName = rdr["ResourceName"] == DBNull.Value ? null : rdr["ResourceName"].ToString(),
                                ProviderName = rdr["ProviderName"] == DBNull.Value ? null : rdr["ProviderName"].ToString(),
                                PrimaryInsuranceName = rdr["PrimaryInsuranceName"] == DBNull.Value ? null : rdr["PrimaryInsuranceName"].ToString(),
                                BotName = rdr["botname"] == DBNull.Value ? null : rdr["botname"].ToString(),
                                Status = string.IsNullOrWhiteSpace(rawStatus) ? "Pending" : rawStatus,
                                CreatedDate = rdr["createdate"] == DBNull.Value ? null : (DateTime?)rdr["createdate"],
                                ModifiedDate = rdr["modifieddate"] == DBNull.Value ? null : (DateTime?)rdr["modifieddate"]
                            });
                        }
                    }
                }
            }
            return list;
        }

        public int UpdateStatuses(Dictionary<int, string> updates)
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
                            using (SqlCommand cmd = new SqlCommand("UpdateR26QueueStatus", con, transaction))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.AddWithValue("@QueueId", item.Key);
                                cmd.Parameters.AddWithValue("@Status", item.Value);
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