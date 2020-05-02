﻿using System;
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
        private static IConfiguration configuration;
        public static List<DbInfo> dbInfoList;
        
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
                ServiceCollection serviceCollection = new ServiceCollection();
                ConfigureServices(serviceCollection);
                int currentHour = 0;
                int workInHours = 8;

                while (true)
                {
                    var servers = configuration.GetSection("Servers").GetChildren().AsEnumerable();
                    dbInfoList = new List<DbInfo>();
                    foreach (var item in servers)
                    {
                        DbInfo dbInfo = new DbInfo
                        {
                            Server = item.Key,
                            ConnectionString = configuration.GetValue<string>($"Servers:{item.Key}:connectionString"),
                            DiskSize = configuration.GetValue<double>($"Servers:{item.Key}:diskSize"),
                            Date = DateTime.UtcNow.ToString("dd.MM.yyyy")
                        };

                        dbInfoList.Add(dbInfo);
                    }

                    DbInfoRequester dbInfoRequester = new DbInfoRequester(dbInfoList);
                    dbInfoRequester.SQLQuery();

                    GoogleDriveSpreadSheetAPI spreadSheetApi = new GoogleDriveSpreadSheetAPI(configuration);
                    spreadSheetApi.CreateGoogleTableDoc();
                    spreadSheetApi.CreateGoogleTableSheet(dbInfoList);
                    spreadSheetApi.DeleteGoogleTableSheet();
                    spreadSheetApi.WriteDataToGoogleTableSheets(dbInfoList, workInHours, currentHour);
                    
                    if (currentHour == workInHours)
                    {
                        currentHour = 0;
                    }
                    currentHour++;
                    int inSeconds = 10;
                    Console.WriteLine($"Updating document: {DateTime.Now}");
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