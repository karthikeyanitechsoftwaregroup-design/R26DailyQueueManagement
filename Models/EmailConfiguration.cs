using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace R26_DailyQueueWinForm.Models
{
    public class EmailConfiguration
    {
        public SmtpSettings SmtpSettings { get; set; } = new SmtpSettings();
        public ErrorNotifications ErrorNotifications { get; set; } = new ErrorNotifications();
        public string ToMails { get; set; } = "";
    }

    public class SmtpSettings
    {
        public string Host { get; set; } = "smtp.gmail.com";
        public int Port { get; set; } = 587;
        public bool EnableSsl { get; set; } = true;
        public string UserName { get; set; } = "";
        public string Password { get; set; } = "";
        public string FromEmail { get; set; } = "";
        public string FromName { get; set; } = "Report Management System";
    }

    public class ErrorNotifications
    {
        public string CcEmails { get; set; } = "";
    }
}