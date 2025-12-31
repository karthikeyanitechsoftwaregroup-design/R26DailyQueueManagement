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
        // ✅ Global decrypted DB connection string
        public static string DbConnectionString { get; private set; }

        // 🔐 Machine-level entropy (DO NOT CHANGE after release)
        private static readonly byte[] _entropy =
            Encoding.UTF8.GetBytes("R26_DailyQueue_SecureKey_v1");

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                // 📄 Load configuration
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // 🔓 Decrypt DB connection string
                DbConnectionString = DecryptString(
                    configuration["ConnectionStrings:DailyQueueConnection"]
                );

                if (string.IsNullOrWhiteSpace(DbConnectionString))
                {
                    MessageBox.Show(
                        "Database connection string is missing or invalid.",
                        "Configuration Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 📧 Email configuration
                var emailConfig = new EmailConfiguration
                {
                    SmtpSettings = new SmtpSettings
                    {
                        Host = configuration["SmtpSettings:Host"],
                        Port = int.Parse(configuration["SmtpSettings:Port"] ?? "587"),
                        EnableSsl = bool.Parse(configuration["SmtpSettings:EnableSsl"] ?? "true"),
                        UserName = configuration["SmtpSettings:UserName"],
                        Password = DecryptString(configuration["SmtpSettings:Password"]),
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
            catch (CryptographicException ex)
            {
                MessageBox.Show(
                    "Failed to decrypt configuration values.\n\n" +
                    "Re-encrypt values on this machine.\n\n" +
                    ex.Message,
                    "Decryption Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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

        // 🔓 Decrypt (Machine-scoped)
        private static string DecryptString(string encryptedString)
        {
            if (string.IsNullOrWhiteSpace(encryptedString))
                return string.Empty;

            try
            {
                byte[] encryptedData = Convert.FromBase64String(encryptedString);

                byte[] decryptedData = ProtectedData.Unprotect(
                    encryptedData,
                    _entropy,
                    DataProtectionScope.LocalMachine
                );

                return Encoding.UTF8.GetString(decryptedData);
            }
            catch
            {
                throw new CryptographicException(
                    "Invalid encrypted value. Re-encryption is required.");
            }
        }

        // 🔐 Encrypt (Machine-scoped) — keep for future re-encryption tools
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
