﻿using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Configuration;
using File = Google.Apis.Drive.v3.Data.File;

namespace GoogleTablesDBInfoLog
{
    public class GoogleDriveSpreadSheetAPI
    {
        private readonly IConfiguration config;
        public string FileID { get; set; }
        
        public GoogleDriveSpreadSheetAPI(IConfiguration config)
        {
            this.config = config;
        }
        
        public File CreateGoogleTableDoc()
        {
            var clientId = config.GetValue<string>("GoogleDrive:ClientId");
            var clientSecret = config.GetValue<string>("GoogleDrive:ClientSecret");
            var userName = config.GetValue<string>("GoogleDrive:UserName");

            string[] scopes =
            {
                DriveService.Scope.Drive,
                DriveService.Scope.DriveFile
            };
            
            var clientSecrets = new ClientSecrets()
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };
            
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            DriveService service = new DriveService(new BaseClientService.Initializer()  
            {  
                HttpClientInitializer = credential,  
                ApplicationName = "GoogleTablesDBInfoLog",  
            });  
            
            var parent = "";//ID of folder if you want to create spreadsheet in specific folder
            var fileName = "DbInfoLog";

            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.Q = $"name='{fileName}' and trashed=false";
            var files = listRequest.Execute();

            File file = null;

            if (files.Files.Count == 0)
            {
                var fileMetadata = new File()  
                {  
                    Name = fileName,  
                    MimeType = "application/vnd.google-apps.spreadsheet",  
                };  
                FilesResource.CreateRequest request = service.Files.Create(fileMetadata);
             
                
//                fileMetadata.Parents = new List<string> { parent }; // Parent folder id or TeamDriveID
                request.Fields = "id"; 
                file = request.Execute();
                Console.WriteLine("File ID: " + file.Id);
                FileID = file.Id;
            }
            else
            {
                foreach (var item in files.Files)
                {
                    FileID = item.Id;
                }
            }
            
            return file;  
        }

        public void WriteDataToGoogleTableSheets(string sheetName, string [] dataBaseInfo)
        {
            var clientId = config.GetValue<string>("GoogleDrive:ClientId");
            var clientSecret = config.GetValue<string>("GoogleDrive:ClientSecret");
            var userName = config.GetValue<string>("GoogleDrive:UserName");

            string[] scopes =
            {
                DriveService.Scope.Drive,
                DriveService.Scope.DriveFile
            };
            
            var clientSecrets = new ClientSecrets()
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };
            
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GoogleTablesDBInfoLog",
            });

            var dataBaseServer = dataBaseInfo[0];
            var dataBaseName = dataBaseInfo[1];
            var dataBaseSize = dataBaseInfo[2];
            var date = dataBaseInfo[3];
            
            var range = $"{dataBaseName}!A:D";
            var valueRange = new ValueRange();

            var objectList = new List<object>()
            {
               dataBaseServer,
               dataBaseName,
               dataBaseSize,
               date
            };
            
            valueRange.Values = new List<IList<object>> {objectList};

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, FileID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();
        }

        public void DeleteGoogleTableSheet()
        {
            var clientId = config.GetValue<string>("GoogleDrive:ClientId");
            var clientSecret = config.GetValue<string>("GoogleDrive:ClientSecret");
            var userName = config.GetValue<string>("GoogleDrive:UserName");

            string[] scopes =
            {
                DriveService.Scope.Drive,
                DriveService.Scope.DriveFile
            };
            
            var clientSecrets = new ClientSecrets()
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };
            
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GoogleTablesDBInfoLog",
            });

            var ssRequest = service.Spreadsheets.Get(FileID);
            Spreadsheet ss = ssRequest.Execute();
            List<string> sheetList = new List<string>();

            foreach(Sheet sheet in ss.Sheets)
            {
                sheetList.Add(sheet.Properties.Title);
            }

            if (sheetList.Contains("Лист1"))
            {
                var deleteSheetRequest = new DeleteSheetRequest();
                deleteSheetRequest.SheetId = 0;
            
                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();

                service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, FileID);

                batchUpdateSpreadsheetRequest.Requests.Add(new Request
                {
                    DeleteSheet = deleteSheetRequest,
                });

                var batchUpdateRequest =
                    service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, FileID);

                batchUpdateRequest.Execute();
            }
        }
        
        public void CreateGoogleTableSheet(List<string> dataBaseNameList)
        {
            var clientId = config.GetValue<string>("GoogleDrive:ClientId");
            var clientSecret = config.GetValue<string>("GoogleDrive:ClientSecret");
            var userName = config.GetValue<string>("GoogleDrive:UserName");

            string[] scopes =
            {
                DriveService.Scope.Drive,
                DriveService.Scope.DriveFile
            };
            
            var clientSecrets = new ClientSecrets()
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };
            
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GoogleTablesDBInfoLog",
            });
            try
            {
                var ssRequest = service.Spreadsheets.Get(FileID);
                Spreadsheet ss = ssRequest.Execute();
                List<string> sheetList = new List<string>();

                foreach(Sheet sheet in ss.Sheets)
                {
                    sheetList.Add(sheet.Properties.Title);
                }

                foreach (var item in dataBaseNameList)
                {
                    if (sheetList.Contains(item))
                    {
                        continue;
                    }
                    
                    string sheetName = $"{item}";
                    var addSheetRequest = new AddSheetRequest();
                    addSheetRequest.Properties = new SheetProperties();
                    addSheetRequest.Properties.Title = sheetName;
            
                    BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                    batchUpdateSpreadsheetRequest.Requests = new List<Request>();

                    service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, FileID);

                    batchUpdateSpreadsheetRequest.Requests.Add(new Request
                    {
                        AddSheet = addSheetRequest
                    });

                    var batchUpdateRequest =
                        service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, FileID);

                    batchUpdateRequest.Execute();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}