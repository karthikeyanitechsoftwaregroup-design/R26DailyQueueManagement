using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using R26_DailyQueueWinForm.Forms;
using R26_DailyQueueWinForm.Models;

namespace R26_DailyQueueWinForm
{
    static class Program
    {
        public static string DbConnectionString { get; private set; }

        // Machine-level entropy
        private static readonly byte[] _entropy =
            Encoding.UTF8.GetBytes("R26_DailyQueue_SecureKey_v1");

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                // Load configuration
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // Decrypt DB connection string
                string rawConnectionString = configuration["ConnectionStrings:DailyQueueConnection"];
                DbConnectionString = TryDecryptOrUsePlain(rawConnectionString);

                if (string.IsNullOrWhiteSpace(DbConnectionString))
                {
                    MessageBox.Show(
                        "Database connection string is missing or invalid.",
                        "Configuration Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Email configuration
                string rawPassword = configuration["SmtpSettings:Password"];

                var emailConfig = new EmailConfiguration
                {
                    SmtpSettings = new SmtpSettings
                    {
                        Host = configuration["SmtpSettings:Host"],
                        Port = int.Parse(configuration["SmtpSettings:Port"] ?? "587"),
                        EnableSsl = bool.Parse(configuration["SmtpSettings:EnableSsl"] ?? "true"),
                        UserName = configuration["SmtpSettings:UserName"],
                        Password = TryDecryptOrUsePlain(rawPassword),
                        FromEmail = configuration["SmtpSettings:FromEmail"],
                        FromName = configuration["SmtpSettings:FromName"]
                                   ?? "Report Management System"
                    },
                    ToMails = configuration["EmailRecipients:ToMails"] ?? "",
                    ErrorNotifications = new ErrorNotifications
                    {
                        CcEmails = configuration["EmailRecipients:CcEmails"] ?? ""
                    }
                };

                // START WITH SAM'S DELIVERY REPORT FORM
                Application.Run(
                    new SamsDeliveryReportForm(
                        DateTime.Today,
                        emailConfig,
                        DbConnectionString,
                        "SDgetSAMSReportStatus",
                        null
                    )
                );
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show(
                    $"Missing configuration file:\n{ex.FileName}",
                    "Startup Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Startup error:\n{ex.Message}",
                    "Startup Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // Try to decrypt
        private static string TryDecryptOrUsePlain(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            if (value.Length > 100 && IsBase64String(value))
            {
                try
                {
                    return DecryptString(value);
                }
                catch
                {
                    // If decryption fails, treat as plain text
                    return value;
                }
            }

            // Return as plain text
            return value;
        }

        // Helper to check if string is Base64
        private static bool IsBase64String(string s)
        {
            if (string.IsNullOrEmpty(s))
                return false;

            s = s.Trim();
            return (s.Length % 4 == 0) &&
                   System.Text.RegularExpressions.Regex.IsMatch(s, @"^[a-zA-Z0-9\+/]*={0,3}$",
                   System.Text.RegularExpressions.RegexOptions.None);
        }

        // Decrypt (Machine-scoped)
        private static string DecryptString(string encryptedString)
        {
            if (string.IsNullOrWhiteSpace(encryptedString))
                return string.Empty;

            byte[] encryptedData = Convert.FromBase64String(encryptedString);

            byte[] decryptedData = ProtectedData.Unprotect(
                encryptedData,
                _entropy,
                DataProtectionScope.LocalMachine
            );

            return Encoding.UTF8.GetString(decryptedData);
        }

        // Encrypt (Machine-scoped)
        public static string EncryptString(string plainText)
        {
            if (string.IsNullOrWhiteSpace(plainText))
                return string.Empty;

            byte[] data = Encoding.UTF8.GetBytes(plainText);

            byte[] encrypted = ProtectedData.Protect(
                data,
                _entropy,
                DataProtectionScope.LocalMachine
            );

            return Convert.ToBase64String(encrypted);
        }
    }
}