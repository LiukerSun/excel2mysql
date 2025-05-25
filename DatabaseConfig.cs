using System.Text.Json;

namespace WindowsFormsApp
{
    public class DatabaseConfig
    {
        public string Server { get; set; } = "localhost";
        public int Port { get; set; } = 3306;
        public string Database { get; set; } = "";
        public string Username { get; set; } = "root";
        public string Password { get; set; } = "";

        public string GetConnectionString()
        {
            return $"Server={Server};Port={Port};Database={Database};Uid={Username};Pwd={Password}";
        }

        public static DatabaseConfig Load()
        {
            string configPath = "dbconfig.json";
            if (File.Exists(configPath))
            {
                string jsonString = File.ReadAllText(configPath);
                return JsonSerializer.Deserialize<DatabaseConfig>(jsonString) ?? new DatabaseConfig();
            }
            return new DatabaseConfig();
        }

        public void Save()
        {
            string jsonString = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText("dbconfig.json", jsonString);
        }
    }
} 