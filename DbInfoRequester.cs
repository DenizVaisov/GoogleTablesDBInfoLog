using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Configuration;
using Npgsql;

namespace GoogleTablesDBInfoLog
{
    /// <summary>
    ///  Класс DbInfoRequester
    /// </summary>
    public class DbInfoRequester
    {
        public static List<DbInfo> dbInfoList;

        /// <summary>
        ///  Конструктор класса DbInfoRequester
        /// </summary>
        /// <param name="_dbInfoList"> Список обобщенный классом DbInfo.</param>
        public DbInfoRequester(List<DbInfo> _dbInfoList)
        {
            dbInfoList = _dbInfoList;
        }
        
        /// <summary>
        ///  Используется для получения данных из SQL запроса
        /// </summary>
        public void SQLQuery()
        {
            try
            {
                foreach (var dbInfo in dbInfoList)
                {
                    using (var connection = new NpgsqlConnection(dbInfo.ConnectionString))
                    {
                        connection.Open();
                        string query = @"SELECT boot_val, current_database(), round(pg_database_size(current_database())/1024.0/1024/1024, 3) 
                                         from pg_settings where name = 'listen_addresses'";
                        var command = new NpgsqlCommand(query, connection);
                        NpgsqlDataReader dataReader = command.ExecuteReader();

                        if (dataReader.Read())
                        {
                            dbInfo.ServerName = dataReader.GetString(0);
                            dbInfo.DataBaseName = dataReader.GetString(1);
                            dbInfo.DataBaseSize = dataReader.GetDouble(2);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}