using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

//"1pkPWcTLgkyEXoCOtn9cmniUPybBvQ71YM6WGOcYpDC4"

namespace ArduinoSerial
{
    public class SheetsHelper
    {
        readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        readonly string ApplicationName = "SerialClient";
        private readonly string SpreadsheetId;
        //private static readonly string sheet = "T";
        private readonly SheetsService _service;
        
        public SheetsHelper(string spreadsheetId)
        {
            SpreadsheetId = spreadsheetId;
            GoogleCredential credential;
            
            using (var stream =
                   new FileStream("/home/williamb/Downloads/key.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            
            _service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
        
        void ReadEntries()
        {
            var range = "A1:B4";
            var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
            
            var response = request.Execute();
            var values = response.Values;
            
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[1]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
        
        static void WriteEntry()
        {
            var range = "A1:B4";
            var valueRange = new ValueRange();
            
            var oblist = new List<object>() { "1", "2", "3", "4" };
            valueRange.Values = new List<IList<object>> { oblist };
            
            var appendRequest = _service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            
            var appendReponse = appendRequest.Execute();
        }
    }
}