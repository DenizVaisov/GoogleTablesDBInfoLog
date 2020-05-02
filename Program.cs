using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Npgsql;

namespace GoogleTablesDBInfoLog
{
    class Program
    {
        private static IConfigurationRoot configuration;
        public static List<string> connectionStringsList;
        public static List<string> dataBaseNameList;
        public static List<string> dataBaseInfoList;

         static int Main(string[] args)
        {
            try
            {
                MainAsync(args).Wait();
                return 0;
            }
            catch
            {
                return 1;
            }
        }

        static async Task MainAsync(string[] args)
        {
            Task task = new Task(() =>
            {
                while (true)
                {
                    ServiceCollection serviceCollection = new ServiceCollection();
                    ConfigureServices(serviceCollection);
            
                    var connectionStrings = configuration.GetSection("ConnectionStrings").GetChildren().AsEnumerable();

                    connectionStringsList = new List<string>();

                    foreach (var item in connectionStrings)
                    {
                        connectionStringsList.Add(item.Value);
                    }
            
                    dataBaseNameList = new List<string>();
                    dataBaseInfoList = new List<string>();
            
                    try
                    {
                        GoogleDriveSpreadSheetAPI spreadSheetApi = new GoogleDriveSpreadSheetAPI(configuration);
                        SQLQuery(dataBaseNameList, dataBaseInfoList);
                        spreadSheetApi.CreateGoogleTableDoc();
                        spreadSheetApi.CreateGoogleTableSheet(dataBaseNameList);
                        spreadSheetApi.DeleteGoogleTableSheet();

                        foreach (var item in dataBaseInfoList)
                        {
                            var items = item.Split(' ');
                    
                            spreadSheetApi.WriteDataToGoogleTableSheets(items);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }

                    int inSeconds = 3600;
                    Console.WriteLine($"Обновление документа: {DateTime.Now}");
                    Thread.Sleep(inSeconds * 1000);
                }
            });
            task.Start();
            
            try
            {
                await task;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static void SQLQuery(List<string> dataBaseNameList, List<string> dataBaseInfoList)
        {
            try
            {
                foreach (var item in connectionStringsList)
                {
                    using (var connection = new NpgsqlConnection(item))
                    {
                        connection.Open();
                        string query = @"SELECT boot_val, current_database(), round(pg_database_size(current_database())/1024.0/1024/1024, 3) 
                                         from pg_settings where name = 'listen_addresses'";
                        var command = new NpgsqlCommand(query, connection);
                        NpgsqlDataReader dataReader = command.ExecuteReader();

                        while (dataReader.Read())
                        {
                            DbInfo dbInfo = new DbInfo
                            {
                                Server = dataReader.GetString(0), 
                                DataBaseName = dataReader.GetString(1), 
                                DataBaseSize = dataReader.GetDouble(2),
                                Date = DateTime.UtcNow.ToString("dd.MM.yyyy")
                            };
                            dataBaseNameList.Add(dbInfo.DataBaseName);
                            dataBaseInfoList.Add(dbInfo.Server + " " + dbInfo.DataBaseName + " " + dbInfo.DataBaseSize + " " + dbInfo.Date);
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
        
        private static void ConfigureServices(IServiceCollection serviceCollection)
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var path = Path.GetFullPath(Path.Combine(currentDirectory, @"..\..\..\"));
            
            configuration = new ConfigurationBuilder()
                .SetBasePath(path)
                .AddJsonFile("appsettings.json", false)
                .Build();

            serviceCollection.AddSingleton(configuration);
        }
    }
}