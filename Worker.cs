using System.Text.Json;
using System.Globalization;
using ClosedXML.Excel;

namespace ApiToExcelService;

public class Worker(ILogger<Worker> logger, HttpClient httpClient, IConfiguration configuration) : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var intervalSeconds = configuration.GetValue<int>("ApiSettings:IntervalSeconds", 60);
        var postUrl = configuration.GetValue<string>("ApiSettings:PostUrl") ?? "https://zapp.zoifintech.com/trackerapi/api/rest/attendance/MonthWiseAttendanceReport";
        var outputFolder = configuration.GetValue<string>("ApiSettings:OutputFolderPath") ?? @"C:\ExcelDownloads";
        var apiKey = configuration.GetValue<string>("ApiSettings:ApiKey");
        var officeStartTimeStr = configuration.GetValue<string>("ApiSettings:OfficeStartTime") ?? "09:30:00";
        var userIds = configuration.GetSection("ApiSettings:UserIds").Get<string[]>() ?? [];

        if (!TimeSpan.TryParse(officeStartTimeStr, out var officeStartTime))
        {
            officeStartTime = new TimeSpan(9, 30, 0);
        }

        if (!string.IsNullOrEmpty(apiKey))
        {
            httpClient.DefaultRequestHeaders.Add("apikey", apiKey);
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
        }

        var formDataConfig = configuration.GetSection("ApiSettings:FormData").GetChildren();
        var baseFormData = new Dictionary<string, string>();
        foreach (var item in formDataConfig)
        {
            if (item.Key != null && item.Value != null)
                baseFormData.Add(item.Key, item.Value);
        }

        bool isFirstRun = true;
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                var now = DateTime.Now;
                var nextRun = new DateTime(now.Year, now.Month, now.Day, 10, 0, 0);
                
                // If it's already past 10:00 AM today, schedule for 10:00 AM tomorrow
                if (now >= nextRun)
                {
                    nextRun = nextRun.AddDays(1);
                }
                
                var delay = nextRun - now;
                
                if (isFirstRun)
                {
                    logger.LogInformation("First run of the service. Executing report generation immediately.");
                    isFirstRun = false;
                }
                else
                {
                    logger.LogInformation("Next report generation scheduled at {NextRun} (waiting {Hours}h {Minutes}m)", nextRun, delay.Hours, delay.Minutes);
                    // Wait until 10:00 AM
                    await Task.Delay(delay, stoppingToken);
                }

                // Time to run the report!
                logger.LogInformation("Worker triggering daily report at: {time}", DateTimeOffset.Now);
                
                var today = DateTime.Now;
                var yesterday = today.AddDays(-1);

                baseFormData["fromDate"] = yesterday.ToString("yyyy-MM-dd");
                baseFormData["toDate"] = today.ToString("yyyy-MM-dd");

                var aggregatedRecords = new List<EmployeeRecord>();

                foreach (var userId in userIds)
                {
                    logger.LogInformation("Fetching data for userId {UserId}", userId);
                    
                    var dict = new Dictionary<string, string>(baseFormData)
                    {
                        { "userId", userId }
                    };

                    using var requestContent = new FormUrlEncodedContent(dict);
                    
                    var requestString = await requestContent.ReadAsStringAsync(stoppingToken);
                    LogToFile(outputFolder, $"REQUEST to {postUrl} - User: {userId} - Data: {requestString}");
                    
                    var response = await httpClient.PostAsync(postUrl, requestContent, stoppingToken);
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        var errorBody = await response.Content.ReadAsStringAsync(stoppingToken);
                        LogToFile(outputFolder, $"ERROR RESPONSE for {userId} - Status: {response.StatusCode} - Body: {errorBody}");
                        logger.LogWarning("Failed to fetch data for {UserId}. Status: {Status}", userId, response.StatusCode);
                        continue;
                    }

                    var jsonString = await response.Content.ReadAsStringAsync(stoppingToken);
                    LogToFile(outputFolder, $"SUCCESS RESPONSE for {userId} - JSON Length: {jsonString.Length} - JSON: {jsonString}");
                    var record = ParseEmployeeRecord(jsonString, today, yesterday);
                    if (record != null)
                    {
                        if (CheckIfLate(record.InTimeToday, officeStartTime))
                        {
                            var lateReason = await FetchExceptionOrPermissionAsync(httpClient, userId, today, stoppingToken);
                            if (!string.IsNullOrEmpty(lateReason))
                            {
                                record.IfLateReason = lateReason;
                            }
                        }
                        aggregatedRecords.Add(record);
                    }
                }

                // Create Excel
                if (aggregatedRecords.Any())
                {
                    CreateCustomExcelFile(aggregatedRecords, outputFolder, today);
                }
                else
                {
                    logger.LogWarning("No data found to generate Excel report.");
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error occurred while gathering data or creating Excel.");
            }
        }
    }

    private bool CheckIfLate(string inTimeToday, TimeSpan threshold)
    {
        if (string.IsNullOrWhiteSpace(inTimeToday) || inTimeToday == "NA" || inTimeToday.Equals("On Leave", StringComparison.OrdinalIgnoreCase) || inTimeToday.Equals("Holiday", StringComparison.OrdinalIgnoreCase))
            return false;

        if (DateTime.TryParseExact(inTimeToday, "hh:mm tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedTime))
        {
            return parsedTime.TimeOfDay > threshold;
        }
        
        if (DateTime.TryParse(inTimeToday, out var generalParsed))
        {
            return generalParsed.TimeOfDay > threshold;
        }

        return false;
    }

    private async Task<string> FetchExceptionOrPermissionAsync(HttpClient client, string userId, DateTime today, CancellationToken stoppingToken)
    {
        try 
        {
            var firstOfMonth = new DateTime(today.Year, today.Month, 1).ToString("yyyy-MM-dd");
            var lastOfMonth = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month)).ToString("yyyy-MM-dd");

            var payloadObj = new
            {
                flag = "G",
                userId = int.Parse(userId),
                fromDate = firstOfMonth,
                toDate = lastOfMonth
            };

            var jsonPayload = JsonSerializer.Serialize(payloadObj);

            using var content1 = new StringContent(jsonPayload, System.Text.Encoding.UTF8, "application/json");
            using var content2 = new StringContent(jsonPayload, System.Text.Encoding.UTF8, "application/json");

            var exTask = client.PostAsync("https://zapp.zoifintech.com/trackerapi/api/rest/attendance/CallExceptionUserList", content1, stoppingToken);
            var permTask = client.PostAsync("https://zapp.zoifintech.com/trackerapi/api/rest/attendance/CallPermissionUserList", content2, stoppingToken);

            await Task.WhenAll(exTask, permTask);

            string finalReason = string.Empty;

            if (exTask.Result.IsSuccessStatusCode)
            {
                var json = await exTask.Result.Content.ReadAsStringAsync(stoppingToken);
                LogToFile(Path.GetDirectoryName(Path.Combine(@"C:\ExcelDownloads", "dummy"))!, $"EXCEPTION API RESPONSE for {userId}: {json}");
                var reason = ParseExceptionReason(json, today);
                if (!string.IsNullOrEmpty(reason)) 
                {
                    finalReason = $"Exception: {reason}";
                }
            }

            if (permTask.Result.IsSuccessStatusCode)
            {
                var json = await permTask.Result.Content.ReadAsStringAsync(stoppingToken);
                LogToFile(Path.GetDirectoryName(Path.Combine(@"C:\ExcelDownloads", "dummy"))!, $"PERMISSION API RESPONSE for {userId}: {json}");
                var reason = ParsePermissionReason(json, today);
                if (!string.IsNullOrEmpty(reason)) 
                {
                    if (!string.IsNullOrEmpty(finalReason)) finalReason += " | ";
                    finalReason += $"Permission: {reason}";
                }
            }

            LogToFile(Path.GetDirectoryName(Path.Combine(@"C:\ExcelDownloads", "dummy"))!, $"FINAL REASON for {userId}: {finalReason}");
            return finalReason;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to fetch exception/permission for {UserId}", userId);
        }
        return string.Empty;
    }

    private string? ParsePermissionReason(string json, DateTime today)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (root.TryGetProperty("data", out var data) && 
                data.ValueKind == JsonValueKind.Object &&
                data.TryGetProperty("permissionData", out var arr) && 
                arr.ValueKind == JsonValueKind.Array)
            {
                var todayStr = today.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
                foreach (var item in arr.EnumerateArray())
                {
                    if (item.TryGetProperty("permissionDate", out var pDate) && pDate.GetString() == todayStr)
                    {
                        return item.TryGetProperty("reason", out var r) ? r.GetString() : "Approved";
                    }
                }
            }
        } catch {}
        return null;
    }

    private string? ParseExceptionReason(string json, DateTime today)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (root.TryGetProperty("data", out var data) && 
                data.ValueKind == JsonValueKind.Object &&
                data.TryGetProperty("userList", out var arr) && 
                arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    var status = item.TryGetProperty("Status", out var s) ? s.GetString() : null;
                    if (status != "Approved") continue;

                    var fromDateStr = item.TryGetProperty("fromDate", out var fd) ? fd.GetString() : null;
                    var toDateStr = item.TryGetProperty("toDate", out var td) ? td.GetString() : null;
                    
                    if (DateTime.TryParseExact(fromDateStr, "dd-MMM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var fromDate) &&
                        DateTime.TryParseExact(toDateStr, "dd-MMM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var toDate))
                    {
                        if (today.Date >= fromDate.Date && today.Date <= toDate.Date)
                        {
                            return item.TryGetProperty("reason", out var r) ? r.GetString() : "Approved";
                        }
                    }
                }
            }
        } catch {}
        return null;
    }

    private void LogToFile(string outputFolder, string message)
    {
        try
        {
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            var logPath = Path.Combine(outputFolder, $"ApiLogs_{DateTime.Now:yyyyMMdd}.txt");
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            File.AppendAllText(logPath, $"[{timestamp}] {message}{Environment.NewLine}");
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to write to text log.");
        }
    }

    private EmployeeRecord? ParseEmployeeRecord(string jsonString, DateTime today, DateTime yesterday)
    {
        try
        {
            using var document = JsonDocument.Parse(jsonString);
            var root = document.RootElement;

            if (root.ValueKind == JsonValueKind.Object &&
                root.TryGetProperty("data", out var dataObj) && 
                dataObj.ValueKind == JsonValueKind.Object &&
                dataObj.TryGetProperty("monthAttendanceList", out var listArr) && 
                listArr.ValueKind == JsonValueKind.Array)
            {
                var todayStr = today.ToString("dd-MMM-yyyy");
                var yesterdayStr = yesterday.ToString("dd-MMM-yyyy");

                string? employeeName = null;
                string inTimeToday = "NA";
                string outTimeYesterday = "NA";
                string leavePresent = "NA";

                foreach (var item in listArr.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object) continue;

                    var dateVal = item.TryGetProperty("attendanceDate", out var d) ? d.GetString() : null;
                    var inTime = item.TryGetProperty("inTime", out var i) ? i.GetString() : null;
                    var outTime = item.TryGetProperty("outTime", out var o) ? o.GetString() : null;
                    
                    if (employeeName == null && item.TryGetProperty("userName", out var u))
                    {
                        employeeName = u.GetString() ?? "Unknown";
                        // Remove " ( ZFS58 )"
                        var braceIndex = employeeName.IndexOf('(');
                        if (braceIndex > 0)
                        {
                            employeeName = employeeName.Substring(0, braceIndex).Trim();
                        }
                    }

                    if (dateVal == todayStr)
                    {
                        string rawInTime = inTime ?? "";
                        if (rawInTime.Equals("On Leave", StringComparison.OrdinalIgnoreCase) || rawInTime.Equals("Holiday", StringComparison.OrdinalIgnoreCase))
                        {
                            inTimeToday = "NA";
                            leavePresent = rawInTime; // "On Leave" or "Holiday"
                        }
                        else if (!string.IsNullOrWhiteSpace(rawInTime) && rawInTime != "--:--")
                        {
                            inTimeToday = rawInTime;
                            leavePresent = "Present";
                        }
                    }
                    else if (dateVal == yesterdayStr)
                    {
                        if (!string.IsNullOrWhiteSpace(outTime) && outTime != "--:--" && !outTime.Equals("On Leave", StringComparison.OrdinalIgnoreCase))
                        {
                            outTimeYesterday = outTime;
                        }
                    }
                }

                if (employeeName != null)
                {
                    if (leavePresent == "On Leave") leavePresent = "Leave";
                    
                    return new EmployeeRecord
                    {
                        EmployeeName = employeeName,
                        InTimeToday = inTimeToday,
                        OutTimeYesterday = outTimeYesterday,
                        LeaveOrPresent = leavePresent,
                        CurrentProject = "" // Missing from API
                    };
                }
            }
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to parse individual employee record.");
        }

        return null;
    }

    private void CreateCustomExcelFile(List<EmployeeRecord> records, string outputFolder, DateTime today)
    {
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Attendance");

        // Top Row: Date
        worksheet.Cell(1, 1).Value = today.ToString("dd-MMM-yy");

        // Headers
        worksheet.Cell(2, 1).Value = "Employee";
        worksheet.Cell(2, 2).Value = "In Time (Today)";
        worksheet.Cell(2, 3).Value = "If Late (Reason and did they get permission or not)";
        worksheet.Cell(2, 4).Value = "Out Time (Yesterday)";
        worksheet.Cell(2, 5).Value = "Leave/Present";
        worksheet.Cell(2, 6).Value = "Current Project";

        int row = 3;
        foreach (var r in records)
        {
            worksheet.Cell(row, 1).Value = r.EmployeeName;
            worksheet.Cell(row, 2).Value = r.InTimeToday;
            worksheet.Cell(row, 3).Value = r.IfLateReason;
            worksheet.Cell(row, 4).Value = r.OutTimeYesterday;
            worksheet.Cell(row, 5).Value = r.LeaveOrPresent;
            worksheet.Cell(row, 6).Value = r.CurrentProject;
            row++;
        }

        // Add sorting/filtering for headers
        var range = worksheet.Range(2, 1, row - 1, 6);
        range.CreateTable();

        worksheet.Columns().AdjustToContents();

        var fileName = $"DailyReport_{today:yyyyMMdd_HHmmss}.xlsx";
        var filePath = Path.Combine(outputFolder, fileName);
        workbook.SaveAs(filePath);

        logger.LogInformation("Successfully created daily Excel file at: {Path}", filePath);
    }

    private class EmployeeRecord
    {
        public string EmployeeName { get; set; } = "";
        public string InTimeToday { get; set; } = "NA";
        public string IfLateReason { get; set; } = "";
        public string OutTimeYesterday { get; set; } = "NA";
        public string LeaveOrPresent { get; set; } = "NA";
        public string CurrentProject { get; set; } = "";
    }
}
