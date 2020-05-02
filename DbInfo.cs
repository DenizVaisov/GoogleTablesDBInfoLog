using System;

namespace GoogleTablesDBInfoLog
{
    public class DbInfo
    {
        public string ServerName { get; set; }
        public string Server { get; set; }
        public double DiskSize { get; set; }
        public string ConnectionString { get; set; }
        public string DataBaseName { get; set; }
        public double DataBaseSize { get; set; }
        public string Date { get; set; }
    }
}