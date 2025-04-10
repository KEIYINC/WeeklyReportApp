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
    }
} 