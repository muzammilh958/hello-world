using ClosedXML.Excel;
using System.Globalization;

var builder = WebApplication.CreateBuilder(args);
var port = Environment.GetEnvironmentVariable("PORT") ?? "8080";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");
builder.Services.AddSingleton<ReconciliationStore>();
var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/health", () => Results.Ok(new { status = "ok", service = "ReconFlow" }));

app.MapPost("/api/reconcile", async (HttpRequest request, ReconciliationStore store) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Use multipart/form-data." });

    var form = await request.ReadFormAsync();
    var invoiceFile = form.Files.GetFile("invoiceFile");
    var paymentFile = form.Files.GetFile("paymentFile");

    if (invoiceFile is null || paymentFile is null)
        return Results.BadRequest(new { error = "Both invoiceFile and paymentFile are required." });

    try
    {
        var invoices = SpreadsheetReader.Read(invoiceFile, "invoice");
        var payments = SpreadsheetReader.Read(paymentFile, "payment");
        var result = ReconciliationEngine.Run(invoices, payments);
        store.Save(result);
        return Results.Ok(new
        {
            result.RunId,
            result.CreatedAt,
            summary = new
            {
                invoiceCount = invoices.Count,
                paymentCount = payments.Count,
                matched = result.Matches.Count(x => x.Status == "Matched"),
                unmatchedInvoices = result.Matches.Count(x => x.Status == "Unmatched Invoice"),
                unmatchedPayments = result.Matches.Count(x => x.Status == "Unmatched Payment"),
                duplicates = result.Matches.Count(x => x.Status == "Duplicate"),
                variance = result.Matches.Sum(x => x.Variance)
            },
            rows = result.Matches.Take(1000)
        });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
}).DisableAntiforgery();

app.MapGet("/api/export/{runId:guid}", (Guid runId, ReconciliationStore store) =>
{
    var result = store.Get(runId);
    if (result is null) return Results.NotFound(new { error = "Reconciliation run not found." });

    using var workbook = new XLWorkbook();
    var ws = workbook.Worksheets.Add("Reconciliation");
    var headers = new[] { "Status", "Invoice Reference", "Payment Reference", "Invoice Amount", "Payment Amount", "Variance", "Invoice Date", "Payment Date" };
    for (var i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];

    var row = 2;
    foreach (var item in result.Matches)
    {
        ws.Cell(row, 1).Value = item.Status;
        ws.Cell(row, 2).Value = item.InvoiceReference ?? "";
        ws.Cell(row, 3).Value = item.PaymentReference ?? "";
        ws.Cell(row, 4).Value = item.InvoiceAmount;
        ws.Cell(row, 5).Value = item.PaymentAmount;
        ws.Cell(row, 6).Value = item.Variance;
        ws.Cell(row, 7).Value = item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "";
        ws.Cell(row, 8).Value = item.PaymentDate?.ToString("yyyy-MM-dd") ?? "";
        row++;
    }
    ws.Columns().AdjustToContents();
    using var stream = new MemoryStream();
    workbook.SaveAs(stream);
    return Results.File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"reconciliation-{runId}.xlsx");
});

app.Run();

record TransactionRow(string Reference, decimal Amount, DateTime? Date, string Source);
record ReconciliationRow(string Status, string? InvoiceReference, string? PaymentReference, decimal InvoiceAmount, decimal PaymentAmount, decimal Variance, DateTime? InvoiceDate, DateTime? PaymentDate);
record ReconciliationResult(Guid RunId, DateTime CreatedAt, List<ReconciliationRow> Matches);

sealed class ReconciliationStore
{
    private readonly Dictionary<Guid, ReconciliationResult> _runs = new();
    private readonly object _lock = new();
    public void Save(ReconciliationResult result) { lock (_lock) _runs[result.RunId] = result; }
    public ReconciliationResult? Get(Guid id) { lock (_lock) return _runs.TryGetValue(id, out var r) ? r : null; }
}

static class SpreadsheetReader
{
    public static List<TransactionRow> Read(IFormFile file, string source)
    {
        var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
        using var stream = file.OpenReadStream();
        return extension switch
        {
            ".xlsx" => ReadXlsx(stream, source),
            ".csv" => ReadCsv(stream, source),
            _ => throw new InvalidOperationException("Only .xlsx and .csv files are supported.")
        };
    }

    private static List<TransactionRow> ReadXlsx(Stream stream, string source)
    {
        using var workbook = new XLWorkbook(stream);
        var ws = workbook.Worksheet(1);
        var used = ws.RangeUsed() ?? throw new InvalidOperationException("Spreadsheet is empty.");
        var firstRow = used.FirstRow();
        var headers = firstRow.Cells().Select((c, i) => new Header(Normalize(c.GetString()), i + 1)).ToList();
        var refCol = Find(headers, "reference", "ref", "orderid", "invoiceid", "transactionid", "id");
        var amountCol = Find(headers, "amount", "transactionamount", "total", "value");
        var dateCol = FindOptional(headers, "date", "transactiondate", "businessday", "invoicedate", "paymentdate");

        var rows = new List<TransactionRow>();
        foreach (var row in used.RowsUsed().Skip(1))
        {
            var reference = row.Cell(refCol).GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(reference)) continue;
            if (!TryDecimal(row.Cell(amountCol).GetFormattedString(), out var amount)) continue;
            DateTime? date = null;
            if (dateCol is not null)
            {
                if (row.Cell(dateCol.Value).TryGetValue<DateTime>(out var dt)) date = dt;
                else if (DateTime.TryParse(row.Cell(dateCol.Value).GetFormattedString(), out dt)) date = dt;
            }
            rows.Add(new TransactionRow(reference, amount, date, source));
        }
        if (rows.Count == 0) throw new InvalidOperationException($"No usable {source} rows found. Required columns: reference and amount.");
        return rows;
    }

    private static List<TransactionRow> ReadCsv(Stream stream, string source)
    {
        using var reader = new StreamReader(stream);
        var headerLine = reader.ReadLine() ?? throw new InvalidOperationException("CSV is empty.");
        var headersRaw = SplitCsv(headerLine);
        var headers = headersRaw.Select((h, i) => new Header(Normalize(h), i)).ToList();
        var refCol = FindZero(headers, "reference", "ref", "orderid", "invoiceid", "transactionid", "id");
        var amountCol = FindZero(headers, "amount", "transactionamount", "total", "value");
        var dateCol = FindOptionalZero(headers, "date", "transactiondate", "businessday", "invoicedate", "paymentdate");
        var rows = new List<TransactionRow>();
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            var cells = SplitCsv(line);
            if (cells.Count <= Math.Max(refCol, amountCol)) continue;
            var reference = cells[refCol].Trim();
            if (string.IsNullOrWhiteSpace(reference) || !TryDecimal(cells[amountCol], out var amount)) continue;
            DateTime? date = null;
            if (dateCol is not null && cells.Count > dateCol.Value && DateTime.TryParse(cells[dateCol.Value], out var dt)) date = dt;
            rows.Add(new TransactionRow(reference, amount, date, source));
        }
        if (rows.Count == 0) throw new InvalidOperationException($"No usable {source} rows found. Required columns: reference and amount.");
        return rows;
    }

    private static bool TryDecimal(string value, out decimal result) => decimal.TryParse(value.Replace(",", "").Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out result) || decimal.TryParse(value, out result);
    private static string Normalize(string value) => new(value.ToLowerInvariant().Where(char.IsLetterOrDigit).ToArray());
    private static int Find(List<Header> headers, params string[] names)
    {
        var header = headers.FirstOrDefault(h => names.Contains(h.Name));
        return header is null ? throw new InvalidOperationException($"Missing required column. Expected one of: {string.Join(", ", names)}") : header.Index;
    }
    private static int? FindOptional(List<Header> headers, params string[] names) => headers.FirstOrDefault(h => names.Contains(h.Name))?.Index;
    private static int FindZero(List<Header> headers, params string[] names) => Find(headers, names);
    private static int? FindOptionalZero(List<Header> headers, params string[] names) => FindOptional(headers, names);
    private sealed record Header(string Name, int Index);

    private static List<string> SplitCsv(string line)
    {
        var result = new List<string>();
        var current = new System.Text.StringBuilder();
        var quoted = false;
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (quoted && i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
                else quoted = !quoted;
            }
            else if (ch == ',' && !quoted) { result.Add(current.ToString()); current.Clear(); }
            else current.Append(ch);
        }
        result.Add(current.ToString());
        return result;
    }
}

static class ReconciliationEngine
{
    public static ReconciliationResult Run(List<TransactionRow> invoices, List<TransactionRow> payments)
    {
        var rows = new List<ReconciliationRow>();
        var paymentUsed = new HashSet<int>();
        var duplicateInvoiceRefs = invoices.GroupBy(x => x.Reference, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1).Select(g => g.Key).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var duplicatePaymentRefs = payments.GroupBy(x => x.Reference, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1).Select(g => g.Key).ToHashSet(StringComparer.OrdinalIgnoreCase);

        for (var invoiceIndex = 0; invoiceIndex < invoices.Count; invoiceIndex++)
        {
            var invoice = invoices[invoiceIndex];
            if (duplicateInvoiceRefs.Contains(invoice.Reference))
            {
                rows.Add(new ReconciliationRow("Duplicate", invoice.Reference, null, invoice.Amount, 0, invoice.Amount, invoice.Date, null));
                continue;
            }

            var matchIndex = -1;
            for (var i = 0; i < payments.Count; i++)
            {
                var payment = payments[i];
                if (paymentUsed.Contains(i) || duplicatePaymentRefs.Contains(payment.Reference)) continue;
                if (payment.Reference.Equals(invoice.Reference, StringComparison.OrdinalIgnoreCase) && payment.Amount == invoice.Amount)
                {
                    matchIndex = i;
                    break;
                }
            }

            if (matchIndex < 0)
            {
                for (var i = 0; i < payments.Count; i++)
                {
                    var payment = payments[i];
                    if (paymentUsed.Contains(i) || duplicatePaymentRefs.Contains(payment.Reference)) continue;
                    if (payment.Amount == invoice.Amount && DatesClose(invoice.Date, payment.Date))
                    {
                        matchIndex = i;
                        break;
                    }
                }
            }

            if (matchIndex >= 0)
            {
                var payment = payments[matchIndex];
                paymentUsed.Add(matchIndex);
                rows.Add(new ReconciliationRow("Matched", invoice.Reference, payment.Reference, invoice.Amount, payment.Amount, invoice.Amount - payment.Amount, invoice.Date, payment.Date));
            }
            else rows.Add(new ReconciliationRow("Unmatched Invoice", invoice.Reference, null, invoice.Amount, 0, invoice.Amount, invoice.Date, null));
        }

        for (var i = 0; i < payments.Count; i++)
        {
            if (paymentUsed.Contains(i)) continue;
            var p = payments[i];
            var status = duplicatePaymentRefs.Contains(p.Reference) ? "Duplicate" : "Unmatched Payment";
            rows.Add(new ReconciliationRow(status, null, p.Reference, 0, p.Amount, -p.Amount, null, p.Date));
        }
        return new ReconciliationResult(Guid.NewGuid(), DateTime.UtcNow, rows);
    }

    private static bool DatesClose(DateTime? a, DateTime? b) => a is null || b is null || Math.Abs((a.Value.Date - b.Value.Date).TotalDays) <= 1;
}