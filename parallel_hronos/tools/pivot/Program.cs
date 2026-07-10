using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

const string RegistrySheet = "Сводный_Реестр";
const string RegistryTable = "Таблица1";
const string VolumeCheckSheet = "Проверка_объемов";
const string FreshnessSheet = "Комплектность_исходников";
const string FilesSheet = "Файлы";
const string ExpectedPivotSheet = "Ожидаемая_сводная";
const string CheckPivotSheet = "Контроль_сводной";
const string RegistryDateFormat = @"[$-419]d\ mmm;@";
const string MotoristColumn1 = "Ф.И.О. моториста п/п 1";
const string MotoristColumn2 = "Ф.И.О. моториста п/п 2";

var opts = Options.Parse(args);
if (opts is null)
{
    Console.Error.WriteLine("usage: HronosPivotBuilder <book.xlsx> --out <final.xlsx> [--date yyyy-mm-dd] [--template pivot_template.xlsx]");
    return 2;
}

try
{
    var input = new FileInfo(opts.InputPath);
    var output = new FileInfo(opts.OutputPath);
    var template = new FileInfo(opts.TemplatePath);

    if (!input.Exists)
    {
        Console.Error.WriteLine($"[ERR] нет файла: {input.FullName}");
        return 3;
    }
    if (!template.Exists)
    {
        Console.Error.WriteLine($"[ERR] нет шаблона сводной: {template.FullName}");
        return 5;
    }

    object[,] data;
    Dictionary<string, object[,]> auxiliarySheets;
    using (var sourcePackage = new ExcelPackage(input))
    {
        var sourceRegistry = sourcePackage.Workbook.Worksheets[RegistrySheet];
        if (sourceRegistry?.Dimension is null)
        {
            Console.Error.WriteLine($"[ERR] нет листа «{RegistrySheet}» или он пуст");
            return 4;
        }

        data = sourceRegistry.Cells[sourceRegistry.Dimension.Address].Value as object[,]
               ?? throw new InvalidOperationException("не удалось прочитать реестр");
        data = DropRegistryColumns(data, MotoristColumn1, MotoristColumn2);
        data = MoveReportDateRowsFirst(data, opts.ReportDate);
        auxiliarySheets = ReadAuxiliarySheets(sourcePackage, VolumeCheckSheet, FreshnessSheet, FilesSheet, ExpectedPivotSheet);
    }

    output.Directory?.Create();
    using (var package = new ExcelPackage(template))
    {
        var book = package.Workbook;
        var registry = book.Worksheets[RegistrySheet]
                       ?? throw new InvalidOperationException($"в шаблоне нет листа «{RegistrySheet}»");

        var table = RebuildRegistryTableFromData(registry, data);
        var headers = ReadHeaders(registry);
        var cols = RequiredColumns.Resolve(headers);
        var filteredRows = FilterRows(data, cols, opts.ReportDate, opts.Peredels);

        BuildControlCrosstab(book, filteredRows);
        RebuildAuxiliarySheets(book, auxiliarySheets);
        registry.View.FreezePanes(2, 1);
        ApplyBasicWidths(registry);

        package.SaveAs(output);
        Console.WriteLine($"[OK] «{RegistryTable}»={table.Address.Address}; контроль за {opts.ReportDate:yyyy-MM-dd}: {filteredRows.Count} строк; передел: {string.Join(", ", opts.Peredels)}");
    }

    PostprocessWorkbookXml(output.FullName, data, RequiredColumns.Resolve(ReadHeadersFromData(data)), opts.ReportDate, opts.Peredels);
    Console.WriteLine($"[OK] источник сводной: {RegistrySheet}!{RegistryTable}");
    return 0;
}
catch (Exception exc)
{
    Console.Error.WriteLine($"[ERR] {exc.GetType().Name}: {exc.Message}");
    return 1;
}

static ExcelTable RebuildRegistryTableFromData(ExcelWorksheet ws, object[,] data)
{
    var rowCount = data.GetLength(0);
    var colCount = data.GetLength(1);
    if (rowCount < 2)
        throw new InvalidOperationException("в реестре нет строк данных");

    if (ws.Tables[RegistryTable] is not null)
        ws.Tables.Delete(RegistryTable, false);

    var oldRows = ws.Dimension?.Rows ?? 0;
    if (oldRows > 0)
        ws.DeleteRow(1, oldRows);

    var rows = new List<object[]>(rowCount);
    for (var r = 0; r < rowCount; r++)
    {
        var row = new object[colCount];
        for (var c = 0; c < colCount; c++)
            row[c] = data[r, c];
        rows.Add(row);
    }

    ws.Cells["A1"].LoadFromArrays(rows);
    var table = ws.Tables.Add(ws.Cells[1, 1, rowCount, colCount], RegistryTable);
    table.TableStyle = TableStyles.Medium7;
    table.ShowFilter = true;

    using (var registryRange = ws.Cells[1, 1, rowCount, colCount])
    {
        registryRange.Style.Font.Name = "Arial";
        registryRange.Style.Font.Size = 9;
    }
    ws.DefaultRowHeight = 15;
    ws.Row(1).Height = 42;
    StyleHeader(ws, colCount);
    FormatDateColumns(ws, ReadHeaders(ws), "Дата. Факт", "Дата выдачи наряд-задания");
    return table;
}

static Dictionary<string, object[,]> ReadAuxiliarySheets(ExcelPackage package, params string[] sheetNames)
{
    var result = new Dictionary<string, object[,]>(StringComparer.Ordinal);
    foreach (var name in sheetNames)
    {
        var ws = package.Workbook.Worksheets[name];
        if (ws?.Dimension is null)
            continue;

        if (ws.Cells[ws.Dimension.Address].Value is object[,] values)
            result[name] = values;
        else
            result[name] = new object[1, 1] { { ws.Cells[ws.Dimension.Address].Value } };
    }
    return result;
}

static object[,] MoveReportDateRowsFirst(object[,] data, DateTime reportDate)
{
    var headers = ReadHeadersFromData(data);
    var cols = RequiredColumns.Resolve(headers);
    var rowCount = data.GetLength(0);
    var colCount = data.GetLength(1);
    if (rowCount <= 2)
        return data;

    var rows = Enumerable.Range(1, rowCount - 1)
        .Select(index => new
        {
            Index = index,
            DateIssue = TryCoerceDate(data[index, cols.DateIssue - 1], out var issueDate) ? issueDate.Date : (DateTime?)null,
            TimeMinutes = TryCoerceTime(data[index, cols.Time - 1], out var minutes) ? minutes : int.MaxValue,
            Unit = Norm(data[index, cols.Unit - 1])
        })
        .OrderByDescending(row => row.DateIssue == reportDate.Date)
        .ThenBy(row => row.DateIssue ?? DateTime.MaxValue)
        .ThenBy(row => row.TimeMinutes)
        .ThenBy(row => row.Unit, StringComparer.Ordinal)
        .ThenBy(row => row.Index)
        .Select(row => row.Index)
        .ToList();

    var result = new object[rowCount, colCount];
    for (var c = 0; c < colCount; c++)
        result[0, c] = data[0, c];
    for (var r = 0; r < rows.Count; r++)
    {
        var sourceRow = rows[r];
        for (var c = 0; c < colCount; c++)
            result[r + 1, c] = data[sourceRow, c];
    }
    return result;
}

static void RebuildAuxiliarySheets(ExcelWorkbook book, Dictionary<string, object[,]> sheets)
{
    foreach (var (name, data) in sheets)
    {
        DeleteSheetIfExists(book, name);
        var ws = book.Worksheets.Add(name);
        var rows = new List<object[]>(data.GetLength(0));
        for (var r = 0; r < data.GetLength(0); r++)
        {
            var row = new object[data.GetLength(1)];
            for (var c = 0; c < data.GetLength(1); c++)
                row[c] = data[r, c];
            rows.Add(row);
        }
        ws.Cells["A1"].LoadFromArrays(rows);
        StyleHeader(ws, data.GetLength(1));
        ApplyBasicWidths(ws);
        ws.View.FreezePanes(2, 1);
    }
}

static void PostprocessWorkbookXml(
    string xlsxPath,
    object[,] data,
    RequiredColumns cols,
    DateTime reportDate,
    IReadOnlyList<string> peredels)
{
    var files = ReadZipEntries(xlsxPath);
    var tableRef = DeduplicateRegistryTable(files);
    EnsurePivotCachesUseRegistryTable(files);
    RemovePivotFields(files, MotoristColumn1, MotoristColumn2);
    ApplyPivotPageFilters(files, data, cols, reportDate, peredels);
    FixFilterDefinedName(files, tableRef);

    var tmp = Path.Combine(
        Path.GetDirectoryName(xlsxPath) ?? ".",
        Path.GetFileNameWithoutExtension(xlsxPath) + ".tmp.xlsx"
    );
    WriteZipEntries(tmp, files);
    File.Move(tmp, xlsxPath, overwrite: true);
}

static Dictionary<string, byte[]> ReadZipEntries(string path)
{
    using var archive = ZipFile.OpenRead(path);
    var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
    foreach (var entry in archive.Entries)
    {
        using var input = entry.Open();
        using var memory = new MemoryStream();
        input.CopyTo(memory);
        files[entry.FullName] = memory.ToArray();
    }
    return files;
}

static void WriteZipEntries(string path, Dictionary<string, byte[]> files)
{
    if (File.Exists(path))
        File.Delete(path);

    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    foreach (var (name, bytes) in files.OrderBy(kv => kv.Key, StringComparer.Ordinal))
    {
        var entry = archive.CreateEntry(name, System.IO.Compression.CompressionLevel.Optimal);
        using var output = entry.Open();
        output.Write(bytes);
    }
}

static string DeduplicateRegistryTable(Dictionary<string, byte[]> files)
{
    var tableParts = files.Keys
        .Where(name => Regex.IsMatch(name, @"^xl/tables/.*\.xml$", RegexOptions.CultureInvariant))
        .Where(name => TableName(GetXml(files, name)) == RegistryTable)
        .OrderBy(name => name, StringComparer.Ordinal)
        .ToList();
    if (tableParts.Count == 0)
        throw new InvalidOperationException($"не найден table-part для «{RegistryTable}»");

    var tablePartSet = tableParts.ToHashSet(StringComparer.Ordinal);
    var sheetInfo = FindWorksheetWithRegistryTable(files, tablePartSet);
    var sheetXml = GetXml(files, sheetInfo.SheetPart);
    var tableRef = DimensionRef(sheetXml);

    var sourcePart = tableParts
        .FirstOrDefault(part => TableRef(GetXml(files, part)) == tableRef)
        ?? tableParts.First();
    var canonicalPart = tablePartSet.Contains("xl/tables/table1.xml")
        ? "xl/tables/table1.xml"
        : sourcePart;

    var canonicalTarget = "../tables/" + Path.GetFileName(canonicalPart);
    var canonicalRel = sheetInfo.Relationships.FirstOrDefault(r => AbsolutePart(r.Target) == canonicalPart);
    var keepRid = canonicalRel?.Id ?? sheetInfo.Relationships.First().Id;

    var tableXml = PatchTableXml(GetXml(files, sourcePart), tableRef);
    files[canonicalPart] = Utf8(tableXml);

    var registryTableParts = sheetInfo.Relationships
        .Select(r => AbsolutePart(r.Target))
        .Where(tablePartSet.Contains)
        .ToHashSet(StringComparer.Ordinal);
    foreach (var part in tableParts)
    {
        if (part != canonicalPart)
            files.Remove(part);
    }

    var relationshipIdsToDrop = sheetInfo.Relationships
        .Where(r => r.Id != keepRid && registryTableParts.Contains(AbsolutePart(r.Target)))
        .Select(r => r.Id)
        .ToHashSet(StringComparer.Ordinal);

    var relsXml = GetXml(files, sheetInfo.RelsPart);
    relsXml = Regex.Replace(
        relsXml,
        @"<Relationship\b[^>]*/>",
        m =>
        {
            var id = XmlAttr(m.Value, "Id");
            if (id is null || !registryTableParts.Contains(AbsolutePart(XmlAttr(m.Value, "Target") ?? "")))
                return m.Value;
            if (relationshipIdsToDrop.Contains(id))
                return "";
            if (id == keepRid)
                return SetXmlAttr(m.Value, "Target", canonicalTarget);
            return m.Value;
        },
        RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(1)
    );
    files[sheetInfo.RelsPart] = Utf8(relsXml);

    sheetXml = Regex.IsMatch(sheetXml, @"<tableParts\b", RegexOptions.CultureInvariant)
        ? Regex.Replace(
            sheetXml,
            @"<tableParts\b[^>]*>.*?</tableParts>",
            $@"<tableParts count=""1""><tablePart r:id=""{keepRid}""/></tableParts>",
            RegexOptions.Singleline | RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(1)
        )
        : sheetXml.Replace("</worksheet>", $@"<tableParts count=""1""><tablePart r:id=""{keepRid}""/></tableParts></worksheet>", StringComparison.Ordinal);
    files[sheetInfo.SheetPart] = Utf8(sheetXml);

    FixContentTypes(files, canonicalPart, tableParts.Where(part => part != canonicalPart));
    return tableRef;
}

static WorksheetTableInfo FindWorksheetWithRegistryTable(Dictionary<string, byte[]> files, IReadOnlySet<string> tableParts)
{
    var sheetParts = files.Keys
        .Where(name => Regex.IsMatch(name, @"^xl/worksheets/sheet\d+\.xml$", RegexOptions.CultureInvariant))
        .OrderBy(name => name, StringComparer.Ordinal);

    foreach (var sheetPart in sheetParts)
    {
        var relsPart = $"xl/worksheets/_rels/{Path.GetFileName(sheetPart)}.rels";
        if (!files.ContainsKey(relsPart))
            continue;

        var relationships = TableRelationships(GetXml(files, relsPart))
            .Where(rel => tableParts.Contains(AbsolutePart(rel.Target)))
            .ToList();
        if (relationships.Count > 0)
            return new WorksheetTableInfo(sheetPart, relsPart, relationships);
    }

    throw new InvalidOperationException($"не найден лист, связанный с «{RegistryTable}»");
}

static List<TableRelationship> TableRelationships(string relsXml)
{
    return Regex.Matches(relsXml, @"<Relationship\b[^>]*/>", RegexOptions.CultureInvariant)
        .Select(m => new TableRelationship(XmlAttr(m.Value, "Id") ?? "", XmlAttr(m.Value, "Target") ?? ""))
        .Where(r => r.Id.Length > 0 && r.Target.Contains("tables/", StringComparison.Ordinal))
        .ToList();
}

static string PatchTableXml(string tableXml, string tableRef)
{
    tableXml = SetXmlElementAttr(tableXml, "table", "id", "1");
    tableXml = SetXmlElementAttr(tableXml, "table", "name", RegistryTable);
    tableXml = SetXmlElementAttr(tableXml, "table", "displayName", RegistryTable);
    tableXml = SetXmlElementAttr(tableXml, "table", "ref", tableRef);
    tableXml = Regex.Replace(
        tableXml,
        @"<autoFilter\b[^>]*/>",
        m => SetXmlAttr(m.Value, "ref", tableRef),
        RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(1)
    );
    return tableXml;
}

static void FixContentTypes(Dictionary<string, byte[]> files, string canonicalPart, IEnumerable<string> droppedParts)
{
    var xml = GetXml(files, "[Content_Types].xml");
    foreach (var part in droppedParts)
    {
        xml = Regex.Replace(
            xml,
            $@"<Override\b[^>]*PartName=""/{Regex.Escape(part)}""[^>]*/>",
            "",
            RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(1)
        );
    }

    if (!xml.Contains($@"PartName=""/{canonicalPart}""", StringComparison.Ordinal))
    {
        const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        xml = xml.Replace(
            "</Types>",
            $@"<Override PartName=""/{canonicalPart}"" ContentType=""{contentType}""/></Types>",
            StringComparison.Ordinal
        );
    }
    files["[Content_Types].xml"] = Utf8(xml);
}

static void EnsurePivotCachesUseRegistryTable(Dictionary<string, byte[]> files)
{
    var cacheParts = files.Keys
        .Where(name => Regex.IsMatch(name, @"^xl/pivotCache/pivotCacheDefinition\d+\.xml$", RegexOptions.CultureInvariant))
        .ToList();

    foreach (var part in cacheParts)
    {
        var xml = GetXml(files, part);
        xml = Regex.Replace(
            xml,
            @"<worksheetSource\b[^>]*/>",
            $@"<worksheetSource name=""{RegistryTable}""/>",
            RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(1)
        );
        xml = SetXmlElementAttr(xml, "pivotCacheDefinition", "refreshOnLoad", "1");
        xml = SetXmlElementAttr(xml, "pivotCacheDefinition", "invalid", "1");
        files[part] = Utf8(xml);
    }
}

static void RemovePivotFields(Dictionary<string, byte[]> files, params string[] fieldNames)
{
    var removedIndexes = new SortedSet<int>();

    foreach (var part in files.Keys.Where(name => Regex.IsMatch(name, @"^xl/pivotCache/pivotCacheDefinition\d+\.xml$", RegexOptions.CultureInvariant)).ToList())
    {
        var doc = XDocument.Parse(GetXml(files, part), LoadOptions.PreserveWhitespace);
        var root = doc.Root ?? throw new InvalidOperationException($"битый XML: {part}");
        var ns = root.Name.Namespace;
        var cacheFieldsRoot = root.Element(ns + "cacheFields");
        var cacheFields = cacheFieldsRoot?.Elements(ns + "cacheField").ToList()
                          ?? throw new InvalidOperationException($"нет cacheFields: {part}");

        for (var i = cacheFields.Count - 1; i >= 0; i--)
        {
            var fieldName = cacheFields[i].Attribute("name")?.Value ?? "";
            if (!fieldNames.Any(name => HeaderEquals(fieldName, name)))
                continue;
            cacheFields[i].Remove();
            removedIndexes.Add(i);
        }

        cacheFieldsRoot.SetAttributeValue("count", cacheFieldsRoot.Elements(ns + "cacheField").Count().ToString(CultureInfo.InvariantCulture));
        files[part] = Utf8(XmlToString(doc));
    }

    if (removedIndexes.Count == 0)
        return;

    var descending = removedIndexes.OrderByDescending(i => i).ToList();
    foreach (var part in files.Keys.Where(name => Regex.IsMatch(name, @"^xl/pivotCache/pivotCacheRecords\d+\.xml$", RegexOptions.CultureInvariant)).ToList())
    {
        var doc = XDocument.Parse(GetXml(files, part), LoadOptions.PreserveWhitespace);
        var root = doc.Root ?? throw new InvalidOperationException($"битый XML: {part}");
        var ns = root.Name.Namespace;
        foreach (var record in root.Elements(ns + "r"))
        {
            var items = record.Elements().ToList();
            foreach (var index in descending)
            {
                if (index < items.Count)
                    items[index].Remove();
            }
        }
        files[part] = Utf8(XmlToString(doc));
    }

    foreach (var part in files.Keys.Where(name => Regex.IsMatch(name, @"^xl/pivotTables/pivotTable\d+\.xml$", RegexOptions.CultureInvariant)).ToList())
    {
        var doc = XDocument.Parse(GetXml(files, part), LoadOptions.PreserveWhitespace);
        var root = doc.Root ?? throw new InvalidOperationException($"битый XML: {part}");
        var ns = root.Name.Namespace;
        var pivotFieldsRoot = root.Element(ns + "pivotFields");
        var pivotFields = pivotFieldsRoot?.Elements(ns + "pivotField").ToList()
                          ?? throw new InvalidOperationException($"нет pivotFields: {part}");
        foreach (var index in descending)
        {
            if (index < pivotFields.Count)
                pivotFields[index].Remove();
        }
        pivotFieldsRoot.SetAttributeValue("count", pivotFieldsRoot.Elements(ns + "pivotField").Count().ToString(CultureInfo.InvariantCulture));
        AdjustPivotFieldReferences(root, removedIndexes.ToList());
        files[part] = Utf8(XmlToString(doc));
    }
}

static void AdjustPivotFieldReferences(XElement pivotRoot, List<int> removedIndexes)
{
    foreach (var element in pivotRoot.Descendants().ToList())
    {
        var local = element.Name.LocalName;
        if (local is "pageField" or "dataField" or "filter" or "pivotFilter")
        {
            if (AdjustFieldIndexAttribute(element, "fld", removedIndexes) < 0)
                element.Remove();
        }
        else if (local == "field" && element.Parent?.Name.LocalName is "rowFields" or "colFields")
        {
            if (AdjustFieldIndexAttribute(element, "x", removedIndexes) < 0)
                element.Remove();
        }
    }

    foreach (var containerName in new[] { "pageFields", "rowFields", "colFields", "dataFields", "filters" })
    {
        var container = pivotRoot.Element(pivotRoot.Name.Namespace + containerName);
        if (container is not null)
            container.SetAttributeValue("count", container.Elements().Count().ToString(CultureInfo.InvariantCulture));
    }
}

static int AdjustFieldIndexAttribute(XElement element, string attributeName, List<int> removedIndexes)
{
    var attr = element.Attribute(attributeName);
    if (attr is null)
        return int.MaxValue;
    if (!int.TryParse(attr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var oldIndex))
        return int.MaxValue;
    if (removedIndexes.Contains(oldIndex))
        return -1;
    var newIndex = oldIndex - removedIndexes.Count(index => index < oldIndex);
    attr.Value = newIndex.ToString(CultureInfo.InvariantCulture);
    return newIndex;
}

static void ApplyPivotPageFilters(
    Dictionary<string, byte[]> files,
    object[,] data,
    RequiredColumns cols,
    DateTime reportDate,
    IReadOnlyList<string> selectedPeredels)
{
    var selectedSet = selectedPeredels.Select(Norm).Where(s => s.Length > 0).ToHashSet(StringComparer.Ordinal);
    PivotFilterIndexes? filterIndexes = null;

    foreach (var part in files.Keys.Where(name => Regex.IsMatch(name, @"^xl/pivotCache/pivotCacheDefinition\d+\.xml$", RegexOptions.CultureInvariant)).ToList())
    {
        var doc = XDocument.Parse(GetXml(files, part), LoadOptions.PreserveWhitespace);
        var root = doc.Root ?? throw new InvalidOperationException($"битый XML: {part}");
        var ns = root.Name.Namespace;
        var cacheFields = root.Element(ns + "cacheFields")?.Elements(ns + "cacheField").ToList()
                          ?? throw new InvalidOperationException($"нет cacheFields: {part}");

        var dateFilter = RebuildDateCacheItems(cacheFields[cols.DateIssue - 1], data, cols.DateIssue, reportDate);
        RemapPivotCacheRecordIndexes(files, cols.DateIssue - 1, dateFilter.OldToNewIndex);
        var peredelFilter = RebuildPeredelCacheItems(cacheFields[cols.Peredel - 1], DistinctTexts(data, cols.Peredel), selectedSet);
        RemapPivotCacheRecordIndexes(files, cols.Peredel - 1, peredelFilter.OldToNewIndex);
        var timeSort = SortTimeCacheItems(cacheFields[cols.Time - 1]);
        RemapPivotCacheRecordIndexes(files, cols.Time - 1, timeSort.OldToNewIndex);
        var dateFactSort = SortDateCacheItems(cacheFields[cols.DateFact - 1]);
        RemapPivotCacheRecordIndexes(files, cols.DateFact - 1, dateFactSort.OldToNewIndex);

        filterIndexes ??= new PivotFilterIndexes(
            dateFilter.ItemIndex,
            dateFilter.ItemCount,
            peredelFilter.SelectedIndexes,
            peredelFilter.ItemCount,
            timeSort.ItemCount,
            dateFactSort.ItemCount
        );
        root.SetAttributeValue("refreshOnLoad", "1");
        root.SetAttributeValue("invalid", "1");
        files[part] = Utf8(XmlToString(doc));
    }

    if (filterIndexes is null)
        throw new InvalidOperationException("нет pivotCacheDefinition для сводной");

    foreach (var part in files.Keys.Where(name => Regex.IsMatch(name, @"^xl/pivotTables/pivotTable\d+\.xml$", RegexOptions.CultureInvariant)).ToList())
    {
        var doc = XDocument.Parse(GetXml(files, part), LoadOptions.PreserveWhitespace);
        var root = doc.Root ?? throw new InvalidOperationException($"битый XML: {part}");
        var ns = root.Name.Namespace;
        var pivotFields = root.Element(ns + "pivotFields")?.Elements(ns + "pivotField").ToList()
                          ?? throw new InvalidOperationException($"нет pivotFields: {part}");

        ApplyPivotFieldItems(
            pivotFields[cols.DateIssue - 1],
            cols.DateIssue - 1,
            filterIndexes.DateItemCount,
            new HashSet<int> { filterIndexes.DateItemIndex },
            multipleSelection: false
        );
        ApplyPivotFieldItems(
            pivotFields[cols.Peredel - 1],
            cols.Peredel - 1,
            filterIndexes.PeredelItemCount,
            filterIndexes.SelectedPeredelIndexes,
            multipleSelection: true
        );
        ApplyAllPivotItems(pivotFields[cols.Time - 1], filterIndexes.TimeItemCount);
        ApplyAllPivotItems(pivotFields[cols.DateFact - 1], filterIndexes.DateFactItemCount);
        ReplacePageFields(root, ns, cols.DateIssue - 1, filterIndexes.DateItemIndex, cols.Peredel - 1);
        files[part] = Utf8(XmlToString(doc));
    }
}

static List<string> DistinctTexts(object[,] data, int column)
{
    var seen = new HashSet<string>(StringComparer.Ordinal);
    var values = new List<string>();
    for (var r = 1; r < data.GetLength(0); r++)
    {
        var text = Norm(data[r, column - 1]);
        if (text.Length > 0 && seen.Add(text))
            values.Add(text);
    }
    return values;
}

static DateFilterIndex RebuildDateCacheItems(XElement cacheField, object[,] data, int column, DateTime reportDate)
{
    var ns = cacheField.Name.Namespace;
    var oldShared = cacheField.Element(ns + "sharedItems");
    var oldItems = oldShared?.Elements().ToList() ?? [];

    var dates = new SortedSet<DateTime>();
    var containsBlank = oldShared?.Attribute("containsBlank")?.Value == "1";
    for (var r = 1; r < data.GetLength(0); r++)
    {
        if (TryCoerceDate(data[r, column - 1], out var value))
            dates.Add(value.Date);
        else
            containsBlank = true;
    }
    dates.Add(reportDate.Date);

    var orderedDates = dates.ToList();                              // хронологический порядок (SortedSet)
    var reportIndex = orderedDates.FindIndex(d => d.Date == reportDate.Date);

    var oldToNew = new Dictionary<int, int>();
    for (var i = 0; i < oldItems.Count; i++)
    {
        var item = oldItems[i];
        if (item.Name.LocalName == "m")
        {
            containsBlank = true;
            oldToNew[i] = orderedDates.Count;
            continue;
        }

        var oldDate = ParseExcelDate(item.Attribute("v")?.Value ?? "");
        if (oldDate is null)
            continue;
        var newIndex = orderedDates.FindIndex(d => d.Date == oldDate.Value.Date);
        if (newIndex >= 0)
            oldToNew[i] = newIndex;
    }

    var shared = new XElement(ns + "sharedItems",
        new XAttribute("containsNonDate", "0"),
        new XAttribute("containsDate", "1"),
        new XAttribute("containsString", "0"),
        new XAttribute("count", (orderedDates.Count + (containsBlank ? 1 : 0)).ToString(CultureInfo.InvariantCulture)),
        new XAttribute("minDate", ExcelDate(orderedDates.Min())),
        new XAttribute("maxDate", ExcelDate(orderedDates.Max()))
    );
    if (containsBlank)
        shared.SetAttributeValue("containsBlank", "1");
    foreach (var value in orderedDates)
        shared.Add(new XElement(ns + "d", new XAttribute("v", ExcelDate(value))));
    if (containsBlank)
        shared.Add(new XElement(ns + "m"));

    ReplaceSharedItems(cacheField, shared);
    return new DateFilterIndex(reportIndex, orderedDates.Count + (containsBlank ? 1 : 0), oldToNew);
}

static PeredelFilterIndex RebuildPeredelCacheItems(
    XElement cacheField,
    List<string> sourceValues,
    HashSet<string> selectedSet)
{
    var ns = cacheField.Name.Namespace;
    var oldShared = cacheField.Element(ns + "sharedItems");
    var oldItems = oldShared?.Elements().ToList() ?? [];
    var containsBlank = oldShared?.Attribute("containsBlank")?.Value == "1";

    var canonicalByKey = sourceValues
        .Where(s => s.Length > 0)
        .GroupBy(CacheTextKey)
        .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);

    var newValues = new List<string>();
    var newIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
    int AddValue(string value)
    {
        var text = Norm(value);
        if (text.Length == 0)
            text = value;
        var key = CacheTextKey(text);
        if (newIndexes.TryGetValue(key, out var existing))
            return existing;
        newIndexes[key] = newValues.Count;
        newValues.Add(text);
        return newValues.Count - 1;
    }

    foreach (var value in sourceValues)
        AddValue(value);
    foreach (var value in selectedSet)
        AddValue(value);

    var oldToNew = new Dictionary<int, int>();
    var oldBlankIndexes = new List<int>();
    for (var i = 0; i < oldItems.Count; i++)
    {
        if (oldItems[i].Name.LocalName == "m")
        {
            containsBlank = true;
            oldBlankIndexes.Add(i);
            continue;
        }

        var oldValue = oldItems[i].Attribute("v")?.Value ?? "";
        var key = CacheTextKey(oldValue);
        var canonical = canonicalByKey.GetValueOrDefault(key) ?? Norm(oldValue);
        oldToNew[i] = AddValue(canonical);
    }
    foreach (var oldBlankIndex in oldBlankIndexes)
        oldToNew[oldBlankIndex] = newValues.Count;

    ReplaceSharedItems(cacheField, BuildTextSharedItems(ns, newValues, containsBlank));
    var selectedIndexes = newValues
        .Select((value, index) => (value, index))
        .Where(pair => selectedSet.Contains(Norm(pair.value)))
        .Select(pair => pair.index)
        .ToHashSet();
    if (selectedIndexes.Count == 0)
        throw new InvalidOperationException("нет выбранных переделов для фильтра сводной");

    return new PeredelFilterIndex(selectedIndexes, newValues.Count + (containsBlank ? 1 : 0), oldToNew);
}

static TimeSortIndex SortTimeCacheItems(XElement cacheField)
{
    var ns = cacheField.Name.Namespace;
    var oldShared = cacheField.Element(ns + "sharedItems");
    if (oldShared is null)
        return new TimeSortIndex(0, []);

    var oldItems = oldShared.Elements().ToList();
    var timedItems = new List<(int OldIndex, XElement Item, int Minutes, string TieBreaker)>();
    var blankItems = new List<(int OldIndex, XElement Item)>();
    var otherItems = new List<(int OldIndex, XElement Item, string TieBreaker)>();

    for (var i = 0; i < oldItems.Count; i++)
    {
        var item = oldItems[i];
        if (item.Name.LocalName == "m")
        {
            blankItems.Add((i, new XElement(item)));
            continue;
        }

        if (TryParseCacheTime(item, out var minutes))
            timedItems.Add((i, new XElement(item), minutes, CacheItemText(item)));
        else
            otherItems.Add((i, new XElement(item), CacheItemText(item)));
    }

    var ordered = new List<(int OldIndex, XElement Item)>();
    ordered.AddRange(
        timedItems
            .OrderBy(item => item.Minutes)
            .ThenBy(item => item.TieBreaker, StringComparer.Ordinal)
            .Select(item => (item.OldIndex, item.Item))
    );
    ordered.AddRange(
        otherItems
            .OrderBy(item => item.TieBreaker, StringComparer.Ordinal)
            .Select(item => (item.OldIndex, item.Item))
    );
    ordered.AddRange(blankItems);

    var oldToNew = new Dictionary<int, int>();
    var newShared = new XElement(oldShared.Name, oldShared.Attributes());
    newShared.RemoveNodes();
    for (var i = 0; i < ordered.Count; i++)
    {
        oldToNew[ordered[i].OldIndex] = i;
        newShared.Add(ordered[i].Item);
    }
    newShared.SetAttributeValue("count", ordered.Count.ToString(CultureInfo.InvariantCulture));
    ReplaceSharedItems(cacheField, newShared);
    return new TimeSortIndex(ordered.Count, oldToNew);
}

static TimeSortIndex SortDateCacheItems(XElement cacheField)
{
    var ns = cacheField.Name.Namespace;
    var oldShared = cacheField.Element(ns + "sharedItems");
    if (oldShared is null)
        return new TimeSortIndex(0, []);

    var oldItems = oldShared.Elements().ToList();
    var datedItems = new List<(int OldIndex, XElement Item, DateTime Date)>();
    var blankItems = new List<(int OldIndex, XElement Item)>();
    var otherItems = new List<(int OldIndex, XElement Item, string TieBreaker)>();

    for (var i = 0; i < oldItems.Count; i++)
    {
        var item = oldItems[i];
        if (item.Name.LocalName == "m")
        {
            blankItems.Add((i, new XElement(item)));
            continue;
        }

        var date = ParseExcelDate(item.Attribute("v")?.Value ?? "");
        if (date is not null)
            datedItems.Add((i, new XElement(item), date.Value));
        else
            otherItems.Add((i, new XElement(item), CacheItemText(item)));
    }

    var ordered = new List<(int OldIndex, XElement Item)>();
    ordered.AddRange(
        datedItems
            .OrderBy(item => item.Date)
            .Select(item => (item.OldIndex, item.Item))
    );
    ordered.AddRange(
        otherItems
            .OrderBy(item => item.TieBreaker, StringComparer.Ordinal)
            .Select(item => (item.OldIndex, item.Item))
    );
    ordered.AddRange(blankItems);

    var oldToNew = new Dictionary<int, int>();
    var newShared = new XElement(oldShared.Name, oldShared.Attributes());
    newShared.RemoveNodes();
    for (var i = 0; i < ordered.Count; i++)
    {
        oldToNew[ordered[i].OldIndex] = i;
        newShared.Add(ordered[i].Item);
    }
    newShared.SetAttributeValue("count", ordered.Count.ToString(CultureInfo.InvariantCulture));
    ReplaceSharedItems(cacheField, newShared);
    return new TimeSortIndex(ordered.Count, oldToNew);
}

static bool TryParseCacheTime(XElement item, out int minutes)
{
    minutes = 0;
    var value = item.Attribute("v")?.Value ?? "";
    if (item.Name.LocalName == "d" && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsedDate))
    {
        minutes = parsedDate.Hour * 60 + parsedDate.Minute;
        return true;
    }

    if (TryParseTime(value, out minutes))
        return true;

    if (double.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var numeric)
        && numeric >= 0
        && numeric < 2)
    {
        var total = (int)Math.Round((numeric % 1) * 24 * 60);
        minutes = total % (24 * 60);
        return true;
    }
    return false;
}

static string CacheItemText(XElement item) => item.Attribute("v")?.Value ?? "";

static void RemapPivotCacheRecordIndexes(Dictionary<string, byte[]> files, int fieldIndex, Dictionary<int, int> oldToNew)
{
    if (oldToNew.Count == 0)
        return;

    foreach (var part in files.Keys.Where(name => Regex.IsMatch(name, @"^xl/pivotCache/pivotCacheRecords\d+\.xml$", RegexOptions.CultureInvariant)).ToList())
    {
        var doc = XDocument.Parse(GetXml(files, part), LoadOptions.PreserveWhitespace);
        var root = doc.Root ?? throw new InvalidOperationException($"битый XML: {part}");
        var ns = root.Name.Namespace;
        foreach (var record in root.Elements(ns + "r"))
        {
            var items = record.Elements().ToList();
            if (fieldIndex >= items.Count)
                continue;
            var item = items[fieldIndex];
            if (item.Name.LocalName != "x")
                continue;
            var attr = item.Attribute("v");
            if (attr is null || !int.TryParse(attr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var oldIndex))
                continue;
            if (oldToNew.TryGetValue(oldIndex, out var newIndex))
                attr.Value = newIndex.ToString(CultureInfo.InvariantCulture);
        }
        files[part] = Utf8(XmlToString(doc));
    }
}

static string CacheTextKey(string value) => Norm(value).ToUpperInvariant();

static DateTime? ParseExcelDate(string value)
{
    if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
        return parsed.Date;
    return null;
}

static bool TryCoerceDate(object? value, out DateTime date)
{
    date = default;
    switch (value)
    {
        case DateTime dt:
            date = dt.Date;
            return true;
        case double serial when serial > 0:
            date = DateTime.FromOADate(serial).Date;
            return true;
        case int serial when serial > 0:
            date = DateTime.FromOADate(serial).Date;
            return true;
        case string text:
            text = Norm(text);
            if (text.Length == 0)
                return false;
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var invariantDate))
            {
                date = invariantDate.Date;
                return true;
            }
            if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out var ruDate))
            {
                date = ruDate.Date;
                return true;
            }
            return false;
        default:
            return false;
    }
}

static bool TryCoerceTime(object? value, out int minutes)
{
    minutes = 0;
    switch (value)
    {
        case DateTime dt:
            minutes = dt.Hour * 60 + dt.Minute;
            return true;
        case double serial when serial >= 0:
            minutes = ((int)Math.Round((serial % 1) * 24 * 60)) % (24 * 60);
            return true;
        case string text:
            return TryParseTime(text, out minutes);
        default:
            return false;
    }
}

static XElement BuildTextSharedItems(XNamespace ns, List<string> values, bool containsBlank = false)
{
    var count = values.Count + (containsBlank ? 1 : 0);
    var shared = new XElement(ns + "sharedItems",
        new XAttribute("containsString", "1"),
        new XAttribute("count", count)
    );
    if (containsBlank)
        shared.SetAttributeValue("containsBlank", "1");
    foreach (var value in values)
        shared.Add(new XElement(ns + "s", new XAttribute("v", value)));
    if (containsBlank)
        shared.Add(new XElement(ns + "m"));
    return shared;
}

static void ReplaceSharedItems(XElement cacheField, XElement sharedItems)
{
    cacheField.Elements(cacheField.Name.Namespace + "sharedItems").Remove();
    cacheField.Add(sharedItems);
}

static void ApplyPivotFieldItems(
    XElement pivotField,
    int fieldIndex,
    int valueCount,
    HashSet<int> visibleIndexes,
    bool multipleSelection)
{
    var ns = pivotField.Name.Namespace;
    pivotField.SetAttributeValue("axis", "axisPage");
    pivotField.SetAttributeValue("showAll", "0");
    pivotField.SetAttributeValue("multipleItemSelectionAllowed", multipleSelection ? "1" : null);
    pivotField.Elements(ns + "items").Remove();
    pivotField.Add(BuildPivotItems(ns, valueCount, visibleIndexes));
}

static void ApplyAllPivotItems(XElement pivotField, int valueCount)
{
    var ns = pivotField.Name.Namespace;
    pivotField.SetAttributeValue("showAll", "0");
    pivotField.SetAttributeValue("sortType", "ascending");
    pivotField.Elements(ns + "items").Remove();
    pivotField.Add(BuildPivotItems(ns, valueCount, Enumerable.Range(0, valueCount).ToHashSet()));
}

static XElement BuildPivotItems(XNamespace ns, int valueCount, HashSet<int> visibleIndexes)
{
    var items = new XElement(ns + "items", new XAttribute("count", valueCount + 1));
    for (var i = 0; i < valueCount; i++)
    {
        var item = new XElement(ns + "item", new XAttribute("x", i));
        if (!visibleIndexes.Contains(i))
            item.SetAttributeValue("h", "1");
        items.Add(item);
    }
    items.Add(new XElement(ns + "item", new XAttribute("t", "default")));
    return items;
}

static void ReplacePageFields(XElement pivotRoot, XNamespace ns, int dateFieldIndex, int dateItemIndex, int peredelFieldIndex)
{
    var pageFields = pivotRoot.Element(ns + "pageFields");
    if (pageFields is null)
    {
        pageFields = new XElement(ns + "pageFields");
        var dataFields = pivotRoot.Element(ns + "dataFields");
        if (dataFields is not null)
            dataFields.AddBeforeSelf(pageFields);
        else
            pivotRoot.Add(pageFields);
    }
    pageFields.RemoveNodes();
    pageFields.SetAttributeValue("count", "2");
    pageFields.Add(
        new XElement(ns + "pageField",
            new XAttribute("fld", dateFieldIndex),
            new XAttribute("hier", "-1"),
            new XAttribute("item", dateItemIndex)
        ),
        new XElement(ns + "pageField",
            new XAttribute("fld", peredelFieldIndex),
            new XAttribute("hier", "-1")
        )
    );
}

static string ExcelDate(DateTime value) => value.Date.ToString("yyyy-MM-dd'T'00:00:00", CultureInfo.InvariantCulture);

static string XmlToString(XDocument doc)
{
    doc.Declaration ??= new XDeclaration("1.0", "UTF-8", "yes");
    return doc.Declaration + doc.ToString(SaveOptions.DisableFormatting);
}

static void FixFilterDefinedName(Dictionary<string, byte[]> files, string tableRef)
{
    const string workbookPart = "xl/workbook.xml";
    if (!files.ContainsKey(workbookPart))
        return;

    var xml = GetXml(files, workbookPart);
    var absoluteRef = ToAbsoluteRef(tableRef);
    xml = Regex.Replace(
        xml,
        $@"{Regex.Escape(RegistrySheet)}!\$A\$1:\$[A-Z]+\$\d+",
        $@"{RegistrySheet}!{absoluteRef}",
        RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(1)
    );
    files[workbookPart] = Utf8(xml);
}

static string DimensionRef(string sheetXml)
{
    var match = Regex.Match(sheetXml, @"<dimension\b[^>]*\bref=""([^""]+)""", RegexOptions.CultureInvariant);
    if (!match.Success)
        throw new InvalidOperationException("у листа-реестра нет dimension");

    var dim = match.Groups[1].Value;
    return dim.Contains(':', StringComparison.Ordinal) ? dim : $"A1:{dim}";
}

static string TableName(string tableXml) => XmlAttr(Regex.Match(tableXml, @"<table\b[^>]*>", RegexOptions.CultureInvariant).Value, "name") ?? "";

static string TableRef(string tableXml) => XmlAttr(Regex.Match(tableXml, @"<table\b[^>]*>", RegexOptions.CultureInvariant).Value, "ref") ?? "";

static string AbsolutePart(string target)
{
    var clean = target.Replace('\\', '/').TrimStart('/');
    while (clean.StartsWith("../", StringComparison.Ordinal))
        clean = clean[3..];
    return clean.StartsWith("xl/", StringComparison.Ordinal) ? clean : "xl/" + clean;
}

static string SetXmlAttr(string tagOrElement, string attr, string value)
{
    var match = Regex.Match(tagOrElement, $@"(?<space>\s){Regex.Escape(attr)}=""[^""]*""", RegexOptions.CultureInvariant);
    if (!match.Success)
        return tagOrElement.Insert(tagOrElement.LastIndexOf('>'), $@" {attr}=""{EscapeAttr(value)}""");

    var replacement = $@"{match.Groups["space"].Value}{attr}=""{EscapeAttr(value)}""";
    return tagOrElement[..match.Index] + replacement + tagOrElement[(match.Index + match.Length)..];
}

static string SetXmlElementAttr(string xml, string element, string attr, string value)
{
    var match = Regex.Match(xml, $@"<{Regex.Escape(element)}\b[^>]*>", RegexOptions.CultureInvariant);
    if (!match.Success)
        return xml;

    var updated = SetXmlAttr(match.Value, attr, value);
    return xml[..match.Index] + updated + xml[(match.Index + match.Length)..];
}

static string? XmlAttr(string tag, string attr)
{
    var match = Regex.Match(tag, $@"\b{Regex.Escape(attr)}=""([^""]*)""", RegexOptions.CultureInvariant);
    return match.Success ? match.Groups[1].Value : null;
}

static string ToAbsoluteRef(string tableRef)
{
    var parts = tableRef.Split(':', 2);
    if (parts.Length != 2)
        return tableRef;
    return $"{AbsoluteCell(parts[0])}:{AbsoluteCell(parts[1])}";
}

static string AbsoluteCell(string cell)
{
    var match = Regex.Match(cell, @"^([A-Z]+)(\d+)$", RegexOptions.CultureInvariant);
    return match.Success ? $"${match.Groups[1].Value}${match.Groups[2].Value}" : cell;
}

static string EscapeAttr(string value)
{
    return value
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal);
}

static string GetXml(Dictionary<string, byte[]> files, string name) => Encoding.UTF8.GetString(files[name]).TrimStart('\uFEFF');

static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);

static List<object?[]> FilterRows(object[,] data, RequiredColumns cols, DateTime reportDate, IReadOnlyList<string> peredels)
{
    var rows = new List<object?[]>();
    var rowCount = data.GetLength(0);
    var colCount = data.GetLength(1);
    for (var r = 1; r < rowCount; r++)
    {
        var issueDate = AsDate(data[r, cols.DateIssue - 1]);
        var peredel = Norm(data[r, cols.Peredel - 1]);
        if (issueDate?.Date != reportDate.Date || !peredels.Contains(peredel))
            continue;

        var row = new object?[colCount];
        for (var c = 0; c < colCount; c++)
            row[c] = data[r, c];
        rows.Add(row);
    }
    return rows;
}

static void BuildControlCrosstab(ExcelWorkbook book, List<object?[]> rows)
{
    DeleteSheetIfExists(book, CheckPivotSheet);
    var ws = book.Worksheets.Add(CheckPivotSheet);
    ws.View.ShowGridLines = false;
    ws.Cells["A1"].Value = "Контроль значений сводной";
    ws.Cells["A1"].Style.Font.Bold = true;
    ws.Cells["A1"].Style.Font.Size = 14;

    var registry = book.Worksheets[RegistrySheet];
    var headers = ReadHeaders(registry);
    var cols = RequiredColumns.Resolve(headers);
    var aggregate = PivotAggregate.FromRows(rows, cols);
    var units = aggregate.Units;

    ws.Cells[3, 1].Value = "Дата";
    ws.Cells[3, 2].Value = "Время";
    for (var i = 0; i < units.Count; i++)
        ws.Cells[3, i + 3].Value = units[i];

    var rowNum = 4;
    foreach (var dateEntry in aggregate.ByDate)
    {
        foreach (var timeEntry in dateEntry.Value.OrderBy(kv => TimeSortKey(kv.Key)))
        {
            ws.Cells[rowNum, 1].Value = dateEntry.Key;
            ws.Cells[rowNum, 1].Style.Numberformat.Format = RegistryDateFormat;
            ws.Cells[rowNum, 2].Value = DisplayTime(timeEntry.Key);
            for (var i = 0; i < units.Count; i++)
                ws.Cells[rowNum, i + 3].Value = RoundOrBlank(timeEntry.Value.GetValueOrDefault(units[i]));
            rowNum++;
        }
    }

    if (rowNum > 4)
        ws.Cells[4, 3, rowNum - 1, Math.Max(3, units.Count + 2)].Style.Numberformat.Format = "#,##0.###";
    StyleHeader(ws, Math.Max(2, units.Count + 2), 3);
    ApplyBasicWidths(ws);
    ws.View.FreezePanes(4, 3);
}

static string[] ReadHeaders(ExcelWorksheet ws)
{
    var lastCol = ws.Dimension.End.Column;
    var headers = new string[lastCol];
    for (var c = 1; c <= lastCol; c++)
        headers[c - 1] = Norm(ws.Cells[1, c].Value);
    return headers;
}

static string[] ReadHeadersFromData(object[,] data)
{
    var colCount = data.GetLength(1);
    var headers = new string[colCount];
    for (var c = 0; c < colCount; c++)
        headers[c] = Norm(data[0, c]);
    return headers;
}

static object[,] DropRegistryColumns(object[,] data, params string[] columnNames)
{
    var headers = ReadHeadersFromData(data);
    var dropIndexes = headers
        .Select((header, index) => (header, index))
        .Where(pair => columnNames.Any(name => HeaderEquals(pair.header, name)))
        .Select(pair => pair.index)
        .ToHashSet();

    if (dropIndexes.Count == 0)
        return data;

    var rows = data.GetLength(0);
    var cols = data.GetLength(1) - dropIndexes.Count;
    var result = new object[rows, cols];
    for (var r = 0; r < rows; r++)
    {
        var targetCol = 0;
        for (var c = 0; c < data.GetLength(1); c++)
        {
            if (dropIndexes.Contains(c))
                continue;
            result[r, targetCol++] = data[r, c];
        }
    }
    return result;
}

static void DeleteSheetIfExists(ExcelWorkbook book, string name)
{
    var old = book.Worksheets[name];
    if (old is not null)
        book.Worksheets.Delete(old);
}

static void StyleHeader(ExcelWorksheet ws, int lastCol, int row = 1)
{
    using var range = ws.Cells[row, 1, row, lastCol];
    range.Style.Font.Bold = true;
    range.Style.Font.Color.SetColor(0, 255, 255, 255);
    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    range.Style.Fill.BackgroundColor.SetColor(0, 48, 84, 150);
    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    range.Style.WrapText = true;
}

static void ApplyBasicWidths(ExcelWorksheet ws)
{
    if (ws.Dimension is null)
        return;

    ws.Cells[ws.Dimension.Address].Style.Font.Name = "Arial";
    ws.Cells[ws.Dimension.Address].Style.Font.Size = 9;
    for (var c = 1; c <= ws.Dimension.End.Column; c++)
    {
        var header = Norm(ws.Cells[1, c].Value);
        var width = CompactColumnWidth(header);
        ws.Column(c).Width = width;
    }
    ws.Row(1).Height = Math.Max(ws.Row(1).Height, 36);
}

static double CompactColumnWidth(string header)
{
    if (HeaderEquals(header, "Подразделение")) return 8.5;
    if (HeaderEquals(header, "Дата. Факт")) return 8;
    if (HeaderEquals(header, "Дата выдачи наряд-задания")) return 10;
    if (HeaderEquals(header, "Смена")) return 5;
    if (HeaderEquals(header, "Ф.И.О. Ответственного")) return 12;
    if (HeaderEquals(header, "Ф.И.О. Машиниста экскватора")) return 12;
    if (HeaderEquals(header, "Ф.И.О. водителя самосвала")) return 13;
    if (HeaderEquals(header, "Время")) return 7;
    if (HeaderEquals(header, "Блок")) return 7;
    if (HeaderEquals(header, "Марка погрузочной единицы")) return 9.5;
    if (HeaderEquals(header, "Инв. № погрузочной единицы")) return 9;
    if (HeaderEquals(header, "Марка транспортировочной единицы")) return 10.5;
    if (HeaderEquals(header, "Инв. № транспортировочной единицы")) return 9.5;
    if (HeaderEquals(header, "Количство машин, шт")) return 7.5;
    if (HeaderEquals(header, "Обьем работ, м3")) return 7.5;
    if (HeaderEquals(header, "Обьем кузова,м3")) return 7.5;
    if (HeaderEquals(header, "Марка промывочного прибора")) return 10.5;
    if (HeaderEquals(header, "Инв. № промывочного прибора")) return 9.5;
    if (HeaderEquals(header, "Передел")) return 12;
    if (HeaderEquals(header, "Откатка, м")) return 7;
    if (HeaderEquals(header, "Примечание")) return 12;
    return Math.Min(16, Math.Max(7, header.Length / 2.4 + 3));
}

static void FormatDateColumns(ExcelWorksheet ws, string[] headers, params string[] names)
{
    foreach (var name in names)
    {
        var idx = Array.FindIndex(headers, h => HeaderEquals(h, name)) + 1;
        if (idx <= 0 || ws.Dimension.End.Row < 2)
            continue;
        ws.Cells[2, idx, ws.Dimension.End.Row, idx].Style.Numberformat.Format = RegistryDateFormat;
    }
}

static string Norm(object? value)
{
    if (value is null)
        return "";
    return string.Join(" ", Convert.ToString(value, CultureInfo.InvariantCulture)!.Replace('\u00a0', ' ').Split(
        new[] { ' ', '\t', '\n', '\r' },
        StringSplitOptions.RemoveEmptyEntries
    ));
}

static bool HeaderEquals(string actual, string expected)
{
    static string Clean(string s) => Norm(s).Replace(" ", "", StringComparison.Ordinal).Replace(".", "", StringComparison.Ordinal).ToUpperInvariant();
    return Clean(actual) == Clean(expected);
}

static DateTime? AsDate(object? value)
{
    if (value is null)
        return null;
    if (value is DateTime dt)
        return dt;
    if (value is double d)
        return DateTime.FromOADate(d);

    var text = Norm(value);
    if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out var parsed))
        return parsed;
    if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
        return parsed;
    return null;
}

static object RoundOrBlank(double value) => Math.Round(value, 0);

static string DisplayTime(string value)
{
    if (TryParseTime(value, out var minutes))
        return $"{minutes / 60 % 24:00}:{minutes % 60:00}:00";
    return value;
}

static int TimeSortKey(string value)
{
    if (!TryParseTime(value, out var minutes))
        return int.MaxValue;
    return minutes;
}

static bool TryParseTime(string value, out int minutes)
{
    minutes = 0;
    var text = Norm(value);
    if (text.Contains(' ', StringComparison.Ordinal))
        text = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? text;
    var parts = text.Split(':');
    if (parts.Length < 2)
        return false;
    if (!int.TryParse(parts[0], out var hour) || !int.TryParse(parts[1], out var minute))
        return false;
    minutes = (hour % 24) * 60 + minute;
    return true;
}

sealed class PivotAggregate
{
    public List<string> Units { get; }
    public SortedDictionary<DateTime, Dictionary<string, Dictionary<string, double>>> ByDate { get; }

    PivotAggregate(
        List<string> units,
        SortedDictionary<DateTime, Dictionary<string, Dictionary<string, double>>> byDate)
    {
        Units = units;
        ByDate = byDate;
    }

    public static PivotAggregate FromRows(List<object?[]> rows, RequiredColumns cols)
    {
        var units = rows
            .Select(r => NormLocal(r[cols.Unit - 1]))
            .Where(s => s.Length > 0)
            .Distinct()
            .OrderBy(s => s, StringComparer.Create(CultureInfo.GetCultureInfo("ru-RU"), ignoreCase: false))
            .ToList();

        var byDate = new SortedDictionary<DateTime, Dictionary<string, Dictionary<string, double>>>();
        foreach (var row in rows)
        {
            var fact = AsDateLocal(row[cols.DateFact - 1]);
            var time = NormLocal(row[cols.Time - 1]);
            var unit = NormLocal(row[cols.Unit - 1]);
            var volume = AsDoubleLocal(row[cols.Volume - 1]);
            if (fact is null || time.Length == 0 || unit.Length == 0)
                continue;

            var date = fact.Value.Date;
            if (!byDate.TryGetValue(date, out var byTime))
            {
                byTime = new Dictionary<string, Dictionary<string, double>>();
                byDate[date] = byTime;
            }
            if (!byTime.TryGetValue(time, out var byUnit))
            {
                byUnit = new Dictionary<string, double>();
                byTime[time] = byUnit;
            }
            byUnit[unit] = byUnit.GetValueOrDefault(unit) + volume;
        }
        return new PivotAggregate(units, byDate);
    }

    static string NormLocal(object? value)
    {
        if (value is null)
            return "";
        return string.Join(" ", Convert.ToString(value, CultureInfo.InvariantCulture)!.Replace('\u00a0', ' ').Split(
            new[] { ' ', '\t', '\n', '\r' },
            StringSplitOptions.RemoveEmptyEntries
        ));
    }

    static DateTime? AsDateLocal(object? value)
    {
        if (value is null)
            return null;
        if (value is DateTime dt)
            return dt;
        if (value is double d)
            return DateTime.FromOADate(d);

        var text = NormLocal(value);
        if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out var parsed))
            return parsed;
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
            return parsed;
        return null;
    }

    static double AsDoubleLocal(object? value)
    {
        if (value is null)
            return 0.0;
        if (value is double d)
            return d;
        if (value is int i)
            return i;
        if (double.TryParse(NormLocal(value).Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        return 0.0;
    }
}

sealed record Options(string InputPath, string OutputPath, DateTime ReportDate, IReadOnlyList<string> Peredels, string TemplatePath)
{
    static readonly string[] DefaultPeredels = ["Транспортировка песков", "Подача песков"];

    public static Options? Parse(string[] args)
    {
        if (args.Length == 0)
            return null;

        var input = args[0];
        string? output = null;
        string? template = null;
        DateTime? reportDate = null;
        var peredels = new List<string>();

        for (var i = 1; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--out":
                    output = args[++i];
                    break;
                case "--date":
                    reportDate = DateTime.ParseExact(args[++i], "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    break;
                case "--template":
                    template = args[++i];
                    break;
                case "--peredel":
                    peredels.Add(NormOption(args[++i]));
                    break;
                default:
                    throw new ArgumentException($"неизвестный аргумент: {args[i]}");
            }
        }

        output ??= Path.Combine(
            Path.GetDirectoryName(input) ?? ".",
            Path.GetFileNameWithoutExtension(input) + "_final.xlsx"
        );
        if (peredels.Count == 0)
            peredels.AddRange(DefaultPeredels);

        return new Options(input, output, reportDate ?? DateTime.Today.AddDays(-1), peredels, template ?? DefaultTemplatePath());
    }

    static string DefaultTemplatePath()
    {
        var candidates = new[]
        {
            Path.Combine(Environment.CurrentDirectory, "../tools/pivot/pivot_template.xlsx"),
            Path.Combine(Environment.CurrentDirectory, "tools/pivot/pivot_template.xlsx"),
            Path.Combine(AppContext.BaseDirectory, "../../../../../../tools/pivot/pivot_template.xlsx"),
        };
        return candidates.Select(Path.GetFullPath).FirstOrDefault(File.Exists) ?? Path.GetFullPath(candidates[0]);
    }

    static string NormOption(string value)
    {
        return string.Join(" ", (value ?? "").Replace('\u00a0', ' ').Split(
            new[] { ' ', '\t', '\n', '\r' },
            StringSplitOptions.RemoveEmptyEntries
        ));
    }
}

sealed record RequiredColumns(
    int Unit,
    int DateFact,
    int DateIssue,
    int Time,
    int Peredel,
    int Volume,
    string UnitName,
    string DateFactName,
    string DateIssueName,
    string TimeName,
    string PeredelName,
    string VolumeName)
{
    public static RequiredColumns Resolve(string[] headers)
    {
        var unit = Find(headers, "Подразделение", "Подразделения");
        var dateFact = Find(headers, "Дата. Факт", "Дата", "Дата факт");
        var dateIssue = Find(headers, "Дата выдачи наряд-задания", "Дата выдачи наряд-заданий");
        var time = Find(headers, "Время", "Час");
        var peredel = Find(headers, "Передел", "Вид работ");
        var volume = Find(headers, "Обьем работ, м3", "Объем работ, м3", "Объем");
        return new RequiredColumns(
            unit,
            dateFact,
            dateIssue,
            time,
            peredel,
            volume,
            headers[unit - 1],
            headers[dateFact - 1],
            headers[dateIssue - 1],
            headers[time - 1],
            headers[peredel - 1],
            headers[volume - 1]
        );
    }

    static int Find(string[] headers, params string[] names)
    {
        foreach (var name in names)
        {
            var idx = Array.FindIndex(headers, h => HeaderEqualsLocal(h, name));
            if (idx >= 0)
                return idx + 1;
        }
        throw new InvalidOperationException($"нет обязательной колонки: {string.Join(" / ", names)}");
    }

    static bool HeaderEqualsLocal(string actual, string expected)
    {
        static string Clean(string s) => string.Join("", (s ?? "").Replace('\u00a0', ' ').Split(
            new[] { ' ', '\t', '\n', '\r', '.' },
            StringSplitOptions.RemoveEmptyEntries
        )).ToUpperInvariant();
        return Clean(actual) == Clean(expected);
    }
}

sealed record TableRelationship(string Id, string Target);

sealed record WorksheetTableInfo(string SheetPart, string RelsPart, List<TableRelationship> Relationships);

sealed record DateFilterIndex(int ItemIndex, int ItemCount, Dictionary<int, int> OldToNewIndex);

sealed record PeredelFilterIndex(HashSet<int> SelectedIndexes, int ItemCount, Dictionary<int, int> OldToNewIndex);

sealed record TimeSortIndex(int ItemCount, Dictionary<int, int> OldToNewIndex);

sealed record PivotFilterIndexes(
    int DateItemIndex,
    int DateItemCount,
    HashSet<int> SelectedPeredelIndexes,
    int PeredelItemCount,
    int TimeItemCount,
    int DateFactItemCount);
