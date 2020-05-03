using System;
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
    /// <summary>
    ///  Класс GoogleDriveSpreadSheetAPI
    /// </summary>
    public class GoogleDriveSpreadSheetAPI
    {
        private readonly IConfiguration config;
        public string FileID { get; set; }
        private string[] defaultListNames = {"Лист1"};
        
        /// <summary>
        ///  Конструктор класса GoogleDriveSpreadSheetAPI
        /// </summary>
        /// <param name="config"> Ссылка на интерфейс IConfiguration.</param>
        public GoogleDriveSpreadSheetAPI(IConfiguration config)
        {
            this.config = config;
        }
        
        /// <summary>
        ///  Используется для создания документа GoogleTables
        /// </summary>
        
        public File CreateGoogleTableDoc()
        {
            // Считывание конфигурации
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
            // Авторизация в Google сервисе

            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            DriveService service = new DriveService(new BaseClientService.Initializer()  
            {  
                HttpClientInitializer = credential,  
                ApplicationName = "GoogleTablesDBInfoLog",  
            });  
            
            // Запрос на поиск документа по имени в GoogleDrive
            var fileName = "DbInfoLog";
            FilesResource.ListRequest listRequest = service.Files.List();
            // Не искать в корзине GoogleDrive
            listRequest.Q = $"name='{fileName}' and trashed=false";
            // Выполнение запроса
            var files = listRequest.Execute();

            File file = null;

            // Проверка на создание документа
            if (files.Files.Count == 0)
            {
                var fileMetadata = new File()  
                {  
                    Name = fileName,  
                    MimeType = "application/vnd.google-apps.spreadsheet",  
                };  
                FilesResource.CreateRequest request = service.Files.Create(fileMetadata);
                request.Fields = "id"; 
                file = request.Execute();
                Console.WriteLine($"Document {fileName} created");
                Console.WriteLine($"ID of created document: {file.Id}");
                FileID = file.Id;
            }
            else
            {
                //Если документ уже создан, то получаем его ID
                foreach (var item in files.Files)
                {
                    FileID = item.Id;
                }
            }
            
            return file;  
        }
        
        /// <summary>
        ///  Используется для записи данных в документ GoogleTable 
        /// </summary>
        /// <param name="dbInfoList"> Ссылка на список обобщенный классом DbInfo</param>
        /// <param name="workInHours"> Целочисленное значение для записи результа за всё время</param>
        /// <param name="currentHour"> Целочисленное значение сколько часов работает приложение</param>
        
        public void WriteDataToGoogleTableSheets(List<DbInfo> dbInfoList, int workInHours, int currentHour)
        {
            // Считывание конфигурации
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
            // Авторизация в Google сервисе
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GoogleTablesDBInfoLog",
            });

            string serverName, server, dataBaseName;
            double dataBaseSize, freeSpace;
            string date;

            foreach (var item in dbInfoList)
            {
                server = item.Server;
                serverName = item.ServerName;
                dataBaseName = item.DataBaseName;
                dataBaseSize = item.DataBaseSize;
                date = item.Date;
                freeSpace = item.DiskSize - item.DataBaseSize;
                
                // Диапазоны для работы с ячейками документа
                var range = $"{server}!A:D";
                var headerRange = $"{server}!A1:D1";
                var totalRange = $"{server}!A:D"; 
                
                var valueRange = new ValueRange();
                var headerValueRange = new ValueRange();
                var totalValueRange = new ValueRange();
                // Данные для записи в документ
                var objectList = new List<object>()
                {
                    server,
                    dataBaseName,
                    dataBaseSize,
                    date
                };
                
                var headerList = new List<object>()
                {
                    "Сервер", "База данных", "Размер в ГБ", "Дата обновления"
                };
                
                var totalList = new List<object>()
                {
                    server, "Свободно", freeSpace, date
                };
                
                // Запросы на запись данных в документ
                headerValueRange.Values = new List<IList<object>> {headerList};
                
                var request = service.Spreadsheets.Values.Update(headerValueRange, FileID, headerRange);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                request.Execute();
                
                valueRange.Values = new List<IList<object>> {objectList};

                var appendRequest = service.Spreadsheets.Values.Append(valueRange, FileID, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                appendRequest.Execute();
                
                // Запись данных в документ за всё время работы
                if (currentHour == workInHours)
                {
                    totalValueRange.Values = new List<IList<object>> {totalList};
                
                    var _request = service.Spreadsheets.Values.Append(totalValueRange, FileID, totalRange);
                    _request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    _request.Execute();
                }
            }
        }
        
        /// <summary>
        ///  Используется для удаления автоматически созданного листа документа GoogleTable 
        /// </summary>
        public void DeleteGoogleTableSheet()
        {
            // Считывание конфигурации
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
            
            // Авторизация в Google сервисе
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, scopes, userName, CancellationToken.None).Result;
            
            // Сервис для работы с GoogleTables
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GoogleTablesDBInfoLog",
            });

            // Получение именования листов документа
            var ssRequest = service.Spreadsheets.Get(FileID);
            Spreadsheet ss = ssRequest.Execute();
            List<string> sheetList = new List<string>();

            // Получение именования листов документа
            foreach(Sheet sheet in ss.Sheets)
            {
                sheetList.Add(sheet.Properties.Title);
            }

            // Цикл по массиву с именованиями автоматически созданных листов
            foreach (var listName in defaultListNames)
            {
                // Если в документе есть автоматически созданный лист, то удалить его
                if (sheetList.Contains(listName))
                {
                    var deleteSheetRequest = new DeleteSheetRequest {SheetId = 0};

                    BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest {Requests = new List<Request>()};

                    service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, FileID);

                    batchUpdateSpreadsheetRequest.Requests.Add(new Request
                    {
                        DeleteSheet = deleteSheetRequest,
                    });

                    var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, FileID);

                    batchUpdateRequest.Execute();
                }
            }
        }
        
        /// <summary>
        ///  Используется для создания листов в документе GoogleTable 
        /// </summary>
        /// <param name="dataBaseNameList"> Список обобщенный типом DbInfo</param>
        public void CreateGoogleTableSheet(List<DbInfo> dataBaseNameList)
        {
            // Считывание конфигурации
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
            
            // Сервис для работы с GoogleTables
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

                // Получение именования листов документа
                foreach(Sheet sheet in ss.Sheets)
                {
                    sheetList.Add(sheet.Properties.Title);
                }

                foreach (var item in dataBaseNameList)
                {
                    // Если лист с таким именованием существует, то перейти к следующей итерации цикла
                    if (sheetList.Contains(item.Server))
                    {
                        continue;
                    }
                    
                    // Запрос на создание листа в документе
                    string sheetName = $"{item.Server}";
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