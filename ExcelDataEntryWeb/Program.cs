using System.Text.Json;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.Services;
using ExcelDataEntryWeb;
using ExcelDataEntryWeb.Dtos;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddSingleton<AppSession>();
builder.Services.Configure<Microsoft.AspNetCore.Http.Json.JsonOptions>(o =>
{
    o.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
    o.SerializerOptions.PropertyNameCaseInsensitive = true;
    o.SerializerOptions.WriteIndented = false;
});

var app = builder.Build();
app.UseDefaultFiles();
app.UseStaticFiles();

// ─── WORKBOOK ────────────────────────────────────────────────────────────────


app.MapPost("/api/workbook/open-path", (OpenWorkbookPathRequest req, AppSession session) =>
{
    var path = req.Path?.Trim('"', ' ', '\'');
    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        return Results.BadRequest(new { error = $"File không tồn tại: '{path}'" });
    try
    {
        session.OpenWorkbook(path);
        return Results.Ok(BuildWorkbookInfo(session));
    }
    catch (IOException ex) when (ex.Message.Contains("used by another process", StringComparison.OrdinalIgnoreCase) ||
                                ex.Message.Contains("locked", StringComparison.OrdinalIgnoreCase) ||
                                ex.HResult == unchecked((int)0x80070020) || ex.HResult == unchecked((int)0x80070021))
    {
        return Results.BadRequest(new { error = "File đang được mở hoặc bị khóa bởi ứng dụng khác (như Microsoft Excel). Vui lòng đóng file ở ứng dụng đó rồi thử lại.", isFileLocked = true });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});

app.MapPost("/api/workbook/open-upload", async (HttpRequest request, AppSession session) =>
{
    if (!request.HasFormContentType) return Results.BadRequest("Cần multipart/form-data");
    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null) return Results.BadRequest("Thiếu file");

    var tmpPath = Path.Combine(Path.GetTempPath(), $"excel_upload_{Guid.NewGuid():N}.xlsx");
    await using (var fs = File.Create(tmpPath))
        await file.CopyToAsync(fs);

    try { session.OpenWorkbook(tmpPath); return Results.Ok(BuildWorkbookInfo(session)); }
    catch (Exception ex) { File.Delete(tmpPath); return Results.BadRequest(ex.Message); }
});

app.MapPost("/api/workbook/close", (AppSession session) =>
{
    session.CloseWorkbook();
    return Results.Ok(new { message = "Đã đóng file." });
});

app.MapGet("/api/workbook/info", (AppSession session) => Results.Ok(BuildWorkbookInfo(session)));

app.MapGet("/api/workbook/sheets", (AppSession session, bool includeHidden = false) =>
    Results.Ok(session.GetWorksheetInfos(includeHidden)));

app.MapPost("/api/workbook/select-sheet", (SelectSheetRequest req, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    try
    {
        session.SelectSheet(req.SheetName, req.HeaderRow, req.NameColumn, req.SampleRow,
            req.AutoSkipBlank, req.SuggestionsDisabled, req.AutoSave);
        return Results.Ok(BuildWorkbookInfo(session));
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
});

app.MapGet("/api/workbook/headers", (AppSession session) =>
    Results.Ok(session.HeaderColumns.Select(h => new HeaderDto(h.ColumnIndex, h.Name, h.IsVisible, h.SheetName, h.ToggleLabel))));

app.MapPost("/api/workbook/headers/{columnIndex:int}/visibility",
    (int columnIndex, SetVisibilityRequest req, AppSession session, string sheetName = "") =>
    {
        session.SetHeaderVisibility(columnIndex, sheetName, req.IsVisible);
        return Results.Ok();
    });

app.MapPost("/api/workbook/headers/show-all", (ShowAllRequest req, AppSession session) =>
{
    session.ShowAllHeaders(req.Show);
    return Results.Ok();
});

app.MapPost("/api/workbook/save", (AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest(new { error = "Chưa có file nào được mở để lưu." });
    session.FlushAndSave();
    if (session.WorkbookService.TrySave(out var err))
        return Results.Ok(new { message = session.StatusMessage });
    return Results.BadRequest(new { error = err, isFileLocked = true });
});

app.MapPost("/api/workbook/save-as", (SaveAsRequest req, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa có file nào được mở.");
    if (string.IsNullOrWhiteSpace(req.Path)) return Results.BadRequest("Thiếu đường dẫn");
    session.TrySaveAs(req.Path);
    return Results.Ok(new { message = $"Đã lưu thành {req.Path}" });
});

app.MapGet("/api/workbook/download", async (AppSession session) =>
{
    if (!session.IsWorkbookLoaded || string.IsNullOrWhiteSpace(session.FilePath) || !File.Exists(session.FilePath))
        return Results.NotFound("Chưa có file nào được mở.");

    session.FlushAndSave();
    session.WorkbookService.TrySave(out _);

    var bytes = await File.ReadAllBytesAsync(session.FilePath);
    var fileName = Path.GetFileName(session.FilePath);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
});

app.MapGet("/api/workbook/candidate-header-rows", (AppSession session) =>
    Results.Ok(session.WorkbookService.GetCandidateHeaderRows()));

// ─── RECORDS ─────────────────────────────────────────────────────────────────

app.MapGet("/api/records", (AppSession session, string? q) =>
{
    if (!session.IsWorkbookLoaded || string.IsNullOrEmpty(session.SelectedSheet))
        return Results.Ok(Enumerable.Empty<RecordDto>());
    var records = session.Records.AsEnumerable();
    if (!string.IsNullOrWhiteSpace(q))
        records = records.Where(r => ExcelDataEntryApp.Infrastructure.VietnameseTextHelper.ContainsNormalized(r.KeyDisplay, q));
    return Results.Ok(records.Select(r => new RecordDto(r.RowIndex, r.LogicalIndex, r.KeyDisplay, r.TooltipPreview)));
});

app.MapPost("/api/records", (AddRecordRequest req, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest(new { error = "Chưa mở file." });
    try
    {
        var rowIndex = session.AddRecord(req.Name ?? string.Empty);
        return Results.Ok(new { rowIndex, message = session.StatusMessage });
    }
    catch (Exception ex) { return Results.BadRequest(new { error = ex.Message }); }
});

app.MapDelete("/api/records/{rowIndex:int}", (int rowIndex, AppSession session, int logicalIndex = -1) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest(new { error = "Chưa mở file." });
    session.DeleteRecord(rowIndex, logicalIndex);
    return Results.Ok(new { message = session.StatusMessage });
});

app.MapPut("/api/records/{rowIndex:int}/rename", (int rowIndex, RenameRecordRequest req, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest(new { error = "Chưa mở file." });
    session.RenameRecord(rowIndex, req.LogicalIndex, req.NewName);
    return Results.Ok();
});

app.MapPost("/api/records/{rowIndex:int}/select", (int rowIndex, AppSession session, int logicalIndex = -1) =>
{
    if (!session.IsWorkbookLoaded || string.IsNullOrEmpty(session.SelectedSheet))
        return Results.BadRequest(new { error = "Chưa mở file hoặc chưa chọn sheet." });
    try
    {
        session.SelectRecord(rowIndex, logicalIndex);
        return Results.Ok(new { message = session.StatusMessage });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});

app.MapGet("/api/records/{rowIndex:int}/fields", (int rowIndex, AppSession session, int logicalIndex = -1) =>
{
    if (!session.IsWorkbookLoaded || string.IsNullOrEmpty(session.SelectedSheet))
        return Results.Ok(new List<EditableFieldDto>());
    try
    {
        session.SelectRecord(rowIndex, logicalIndex);
        return Results.Ok(session.EditableFields.Select(f => EditableFieldDto.From(f)).ToList());
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});

app.MapPost("/api/records/cell", (UpdateCellRequest req, AppSession session) =>
{
    if (session.SelectedRecord is null) return Results.BadRequest("Chưa chọn record.");
    session.UpdateCell(req.SheetName ?? string.Empty, req.RowIndex, req.ColumnIndex, req.Value ?? string.Empty);
    return Results.Ok(new { message = session.StatusMessage });
});

app.MapPost("/api/records/toggle-all-headers", (AppSession session) =>
{
    var label = session.ToggleAllHeaders();
    return Results.Ok(new { label, fields = session.EditableFields.Select(f => EditableFieldDto.From(f)).ToList() });
});

// ─── DROPDOWN ────────────────────────────────────────────────────────────────

app.MapGet("/api/dropdown/{columnIndex:int}", (int columnIndex, AppSession session,
    int rowIndex = 0, string sheetName = "") =>
    Results.Ok(session.GetDropdownOptions(columnIndex, rowIndex, sheetName)));

app.MapGet("/api/dropdown/{columnIndex:int}/search", (int columnIndex, AppSession session,
    string q = "", int rowIndex = 0, string sheetName = "") =>
    Results.Ok(session.SearchDropdownOptions(sheetName, rowIndex, columnIndex, q)));

app.MapPost("/api/dropdown/{columnIndex:int}/dependent", (int columnIndex, DependentDropdownRequest req, AppSession session) =>
{
    var overrides = req.Overrides?.ToDictionary(
        kv => new CellKey(kv.SheetName, kv.RowIndex, kv.ColumnIndex), kv => kv.Value)
        ?? [];
    return Results.Ok(session.GetDependentDropdownOptions(req.SheetName ?? string.Empty, req.RowIndex, columnIndex, overrides));
});

// ─── MULTI-SHEET ─────────────────────────────────────────────────────────────

app.MapGet("/api/multisheet/config", (AppSession session) =>
    Results.Ok(session.TryGetSavedMultiSheetConfig()));

app.MapPost("/api/multisheet/apply", (MultiSheetImportSession sessionData, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    try
    {
        session.ApplyMultiSheetImport(sessionData);
        return Results.Ok(BuildWorkbookInfo(session));
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
});

app.MapDelete("/api/multisheet", (AppSession session) =>
{
    session.ExitMultiSheetMode();
    return Results.Ok(BuildWorkbookInfo(session));
});

app.MapGet("/api/multisheet/worksheets", (AppSession session, bool includeHidden = false) =>
    Results.Ok(session.GetWorksheetInfos(includeHidden)));

app.MapGet("/api/multisheet/headers", (AppSession session, string sheetName, int fromRow = 1, int toRow = 3) =>
{
    var headers = session.WorkbookService.ReadHeaderBlock(sheetName, fromRow, toRow);
    return Results.Ok(headers);
});

// ─── IMPORT LIST (Data Link) ──────────────────────────────────────────────────

app.MapGet("/api/import-list/sheets", (AppSession session, string filePath) =>
{
    if (!File.Exists(filePath)) return Results.BadRequest("File không tồn tại.");
    try
    {
        var linkReader = new ExcelLookupReader();
        linkReader.Open(filePath, string.Empty);
        var sheets = linkReader.GetWorksheetNames();
        linkReader.Dispose();
        return Results.Ok(sheets);
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
});

app.MapPost("/api/import-list/link", (DataLinkProfile profile, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    var error = session.RunDataLink(profile);
    if (error is not null) return Results.BadRequest(error);
    return Results.Ok(new { message = session.StatusMessage });
});

app.MapGet("/api/import-list/profiles", (AppSession session) =>
{
    if (string.IsNullOrWhiteSpace(session.ProfileKey)) return Results.Ok(Array.Empty<string>());
    return Results.Ok(session.DataLinkProfileStore.GetSavedProfileNames(session.ProfileKey));
});

app.MapGet("/api/import-list/profiles/{name}", (string name, AppSession session) =>
{
    if (string.IsNullOrWhiteSpace(session.ProfileKey)) return Results.NotFound();
    var profile = session.DataLinkProfileStore.TryLoadNamed(session.ProfileKey, name);
    return profile is not null ? Results.Ok(profile) : Results.NotFound();
});

app.MapPost("/api/import-list/profiles/{name}", (string name, DataLinkProfile profile, AppSession session) =>
{
    if (string.IsNullOrWhiteSpace(session.ProfileKey)) return Results.BadRequest("Chưa có profile key.");
    session.DataLinkProfileStore.SaveNamed(session.ProfileKey, name, profile);
    session.DataLinkProfileStore.SaveLastNamedProfileName(session.ProfileKey, name);
    return Results.Ok();
});

app.MapDelete("/api/import-list/profiles/{name}", (string name, AppSession session) =>
{
    if (string.IsNullOrWhiteSpace(session.ProfileKey)) return Results.BadRequest("Chưa có profile key.");
    session.DataLinkProfileStore.DeleteNamed(session.ProfileKey, name);
    return Results.Ok();
});

app.MapGet("/api/import-list/columns", (AppSession session, string filePath, string sheetName, int headerRow) =>
{
    if (!File.Exists(filePath)) return Results.BadRequest("File không tồn tại.");
    try
    {
        using var lookup = new ExcelLookupReader();
        lookup.Open(filePath, sheetName);
        var headers = lookup.ReadHeaderRow(headerRow);
        return Results.Ok(headers);
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
});

// ─── EXPORT ───────────────────────────────────────────────────────────────────

app.MapGet("/api/export/hssk/profile", (AppSession session) =>
    Results.Ok(session.GetHsskExportProfile()));

app.MapPost("/api/export/hssk/profile", (HsskExportProfile profile, AppSession session) =>
{
    session.SaveHsskExportProfile(profile);
    return Results.Ok();
});

app.MapPost("/api/export/hssk/run", async (HsskExportRequest req, AppSession session, HttpContext ctx) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    if (!File.Exists(req.TemplateDocxPath)) return Results.BadRequest("Không tìm thấy file template.");
    try
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"hssk_{Guid.NewGuid():N}.docx");
        var profile = session.GetHsskExportProfile();
        // Build row value dictionaries for each visible record
        var allRecordValues = new List<IReadOnlyDictionary<int, string>>();
        foreach (var record in session.Records)
        {
            session.SelectRecord(record.RowIndex, record.LogicalIndex);
            var rowValues = session.EditableFields.ToDictionary(f => f.ColumnIndex, f => f.Value);
            allRecordValues.Add(rowValues);
        }
        if (allRecordValues.Count == 0) return Results.BadRequest("Áp dụng cho 0 bản ghi.");
        WordTemplateExportService.ExportManyRecordsToSingleFile(req.TemplateDocxPath, outputPath, allRecordValues);
        var bytes = await File.ReadAllBytesAsync(outputPath);
        File.Delete(outputPath);
        return Results.File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hssk_export.docx");
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
});

app.MapGet("/api/export/import-file/profile", (AppSession session) =>
    Results.Ok(session.GetImportFileExportProfile()));

app.MapPost("/api/export/import-file/profile", (ImportFileExportProfile profile, AppSession session) =>
{
    session.SaveImportFileExportProfile(profile);
    return Results.Ok();
});

app.MapPost("/api/export/import-file/run", async (ImportFileExportRequest req, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    try
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"import_{Guid.NewGuid():N}.xlsx");
        var profile = session.GetImportFileExportProfile();
        // Build sheet plans
        var plans = new List<ImportFileSheetPlan>();
        if (session.IsMultiSheetMode && session.MultiSheetSession is not null)
        {
            foreach (var sheetCfg in session.MultiSheetSession.Sheets)
            {
                var sheetHeaders = session.HeaderColumns
                    .Where(h => string.Equals(h.SheetName, sheetCfg.SheetName, StringComparison.Ordinal))
                    .Select(h => h.ColumnIndex).ToList();
                var lastRow = session.WorkbookService.GetEndRow(sheetCfg.SheetName);
                plans.Add(new ImportFileSheetPlan
                {
                    SheetName = sheetCfg.SheetName,
                    FirstDataRow = sheetCfg.FirstDataRow,
                    LastDataRow = lastRow,
                    SampleRow = sheetCfg.SampleRow > 0 ? sheetCfg.SampleRow : sheetCfg.HeaderLastRow,
                    ColumnIndexes = sheetHeaders,
                    GenderColumnIndex = profile.GenderColumnIndex,
                    SkipSampleColumnIndexes = profile.SkipSampleColumnIndexes ?? []
                });
            }
        }
        else
        {
            var lastRow = session.WorkbookService.GetEndRow();
            plans.Add(new ImportFileSheetPlan
            {
                SheetName = session.SelectedSheet,
                FirstDataRow = session.HeaderRowNumber + 1,
                LastDataRow = lastRow,
                SampleRow = session.SampleRowThreshold > 0 ? session.SampleRowThreshold : session.HeaderRowNumber,
                ColumnIndexes = session.HeaderColumns.Select(h => h.ColumnIndex).ToList(),
                GenderColumnIndex = profile.GenderColumnIndex,
                SkipSampleColumnIndexes = profile.SkipSampleColumnIndexes ?? []
            });
        }
        var result = ImportFileExportService.Export(session.WorkbookService.WorkbookPath, outputPath, plans);
        var bytes = await File.ReadAllBytesAsync(outputPath);
        File.Delete(outputPath);
        return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "import_export.xlsx");
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
});

// ─── SETTINGS ────────────────────────────────────────────────────────────────

app.MapGet("/api/settings/session", (AppSession session) =>
    Results.Ok(new
    {
        session.HeaderRowNumber,
        session.NameColumnIndex,
        session.SampleRowThreshold,
        session.AutoSkipBlankHeaders,
        session.IsAutoSaveEnabled,
        session.SuggestionsDisabled,
        session.IsMultiSheetMode,
        session.StatusMessage,
        session.FilePath,
        session.SelectedSheet,
        isLoaded = session.IsWorkbookLoaded
    }));

app.MapPost("/api/settings/sort-values", (AppSession session, int columnIndex) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    var firstDataRow = session.HeaderRowNumber + 1;
    var excludedRow = session.SampleRowThreshold;
    var values = session.WorkbookService.GetDistinctColumnValues(firstDataRow, columnIndex, excludedRow);
    return Results.Ok(values);
});

app.MapPost("/api/settings/sort", (SortRequest req, AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    session.SortRecords(req.ColumnIndex, req.Mode, req.CustomValues);
    return Results.Ok(new { message = session.StatusMessage });
});

app.MapGet("/api/settings/skip-columns/scan", (AppSession session) =>
{
    if (!session.IsWorkbookLoaded) return Results.BadRequest("Chưa mở file.");
    var sheetName = session.WorkbookService.ActiveSheetName ?? session.SelectedSheet;
    var markerCols = session.WorkbookService.FindColumnsWithMarkerInRows(sheetName, 1, 2,
        HsskExportGenderHelper.FemaleColumnMarker);
    return Results.Ok(markerCols);
});

// ─── CONFIG BACKUP ────────────────────────────────────────────────────────────

app.MapGet("/api/config/backup-info", (AppSession session) =>
    Results.Ok(new
    {
        backupPath = session.AppConfigBackupStore.BackupFilePath,
        hsskProfilesPath = session.HsskExportProfileStore.ProfileFilePath,
        importFileProfilesPath = session.ImportFileExportProfileStore.ProfileFilePath
    }));

app.MapPost("/api/config/export-backup", async (AppSession session) =>
{
    var path = session.AppConfigBackupStore.BackupFilePath;
    if (!File.Exists(path)) return Results.BadRequest("Chưa có backup.");
    var bytes = await File.ReadAllBytesAsync(path);
    return Results.File(bytes, "application/json", "config-backup.json");
});

app.MapPost("/api/config/import-backup", async (HttpRequest request, AppSession session) =>
{
    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null) return Results.BadRequest("Thiếu file");
    var tmpPath = Path.GetTempFileName();
    await using (var fs = File.Create(tmpPath))
        await file.CopyToAsync(fs);

    try
    {
        // Copy to backup location and restore
        File.Copy(tmpPath, session.AppConfigBackupStore.BackupFilePath, overwrite: true);
        session.AppConfigBackupStore.TryRestoreIfPresent(
            session.SessionSettingsStore, session.MultiSheetConfigStore,
            session.HsskExportProfileStore, session.ImportFileExportProfileStore);
        return Results.Ok(new { message = "Đã khôi phục backup." });
    }
    catch (Exception ex) { return Results.BadRequest(ex.Message); }
    finally { File.Delete(tmpPath); }
});

// ─── STATUS ───────────────────────────────────────────────────────────────────
app.MapGet("/api/status", () => Results.Ok(new { ok = true, version = "1.0" }));

app.Run();

// ─── HELPERS ─────────────────────────────────────────────────────────────────

static WorkbookInfoDto BuildWorkbookInfo(AppSession session) => new(
    session.IsWorkbookLoaded,
    session.FilePath,
    session.SelectedSheet,
    session.StatusMessage,
    session.IsMultiSheetMode,
    session.GetSheets().ToList(),
    session.HeaderRowNumber,
    session.NameColumnIndex,
    session.SampleRowThreshold,
    session.IsAutoSaveEnabled,
    session.SuggestionsDisabled,
    session.AutoSkipBlankHeaders
);
