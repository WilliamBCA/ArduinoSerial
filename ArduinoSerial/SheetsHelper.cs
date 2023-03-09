namespace ArduinoSerial;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

public class SheetsHelper
{
    readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    readonly string ApplicationName = "SpreadsheetClient";

    private readonly string SpreadsheetId;

    //private static readonly string sheet = "T";
    private readonly SheetsService _service;

    public SheetsHelper(string spreadsheetId)
    {
        SpreadsheetId = spreadsheetId;
        GoogleCredential credential;

        using (var stream =
               new FileStream("/Users/oscarstromsborg/Library/Mobile Documents/com~apple~CloudDocs/Downloads/key.json", FileMode.Open, FileAccess.Read))
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


    public async Task AddRowToSheetAsync(string[] values)
    {
        // Get the current data in the sheet to determine the last row
        string range = "Sheet1!A:Z";
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = await request.ExecuteAsync();
        var numRows = response.Values != null ? response.Values.Count : 0;

        // Append the new row to the end o  f the sheet
        range = $"Sheet1!A{numRows + 1}:Z{numRows + 1}";
        var valueRange = new ValueRange { Values = new List<IList<object>> { values } };
        var appendRequest = _service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
        appendRequest.ValueInputOption =
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        var appendResponse = await appendRequest.ExecuteAsync();

        Console.WriteLine($"Added row to sheet: {string.Join(", ", values)}");
    }


    public IList<IList<string>> GetAllRowsFromSheet()
    {
        string range = "Sheet1!A:Z";
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();
        var numRows = response.Values != null ? response.Values.Count : 0;

        var stringValues = new List<IList<string>>();
        foreach (var row in response.Values)
        {
            var stringRow = new List<string>();
            foreach (var value in row)
            {
                stringRow.Add(value.ToString());
            }

            stringValues.Add(stringRow);
        }

        return stringValues;
    }
}
    