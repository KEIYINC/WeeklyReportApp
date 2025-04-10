using System;
using System.IO;
using System.Text.Json;
using WeeklyReportApp.Models;

namespace WeeklyReportApp.Utils
{
    public static class ConfigManager
    {
        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WeeklyReportApp",
            "config.json");

        public static bool IsFirstRun()
        {
            return !File.Exists(ConfigPath);
        }

        public static UserInfo LoadUserInfo()
        {
            if (!File.Exists(ConfigPath))
                return null;

            string json = File.ReadAllText(ConfigPath);
            return JsonSerializer.Deserialize<UserInfo>(json);
        }

        public static void SaveUserInfo(UserInfo userInfo)
        {
            string directory = Path.GetDirectoryName(ConfigPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = JsonSerializer.Serialize(userInfo);
            File.WriteAllText(ConfigPath, json);
        }
    }
} 