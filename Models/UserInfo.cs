using System;

namespace WeeklyReportApp.Models
{
    public class UserInfo
    {
        public string FullName { get; set; }
        public string Email { get; set; }
        public string TargetEmail { get; set; }
        public string TemplatePath { get; set; }
        public DateTime LastReportDate { get; set; }
        public string EmailPassword { get; set; }  
        
        public string SmtpServer { get; set; }
        public int SmtpPort { get; set; }
        public bool EnableSsl { get; set; }


    }
} 