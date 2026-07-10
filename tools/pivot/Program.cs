using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// BookBuilder — берёт шаблон с нативной сводной и подменяет данные реестра свежими.
//   • «Сводный_Реестр»  = ПОЛНАЯ консолидация (ВСЕ переделы) + умная «Таблица1» —
//     это то, что смотрит пользователь И это ЕДИНСТВЕННЫЙ источник сводной.
//   • сводная «Сводная_пески» строится ИЗ «Таблица1» (worksheetSource name="Таблица1").
//     Её кэш (sharedItems + pivotCacheRecords) и фильтр «Передел»=«Транспортировка
//     песков» МАТЕРИАЛИЗУЕТ Python-пост-шаг (pivot_postprocess.py) прямо из данных
//     листа, после чего ставит refreshOnLoad="0" — Excel НЕ перестраивает кэш при
//     открытии, индексный выбор фильтра не «съезжает».
//   EPPlus 4.5.3.3 не умеет пересчитывать pivotCacheRecords из источника — поэтому
//   BookBuilder только пишет данные реестра; всю сводную формирует Python-пост-шаг.
//   BookBuilder при round-trip плодит дубль table-файла «Таблица1» — его тоже уберёт
//   Python-пост-шаг.
//   Использование: BookBuilder <data.xlsx> <template.xlsx> [out.xlsx=data.xlsx]
class Program
{
    const string SHEET = "Сводный_Реестр";
    const string TABLE = "Таблица1";
    const string PERDEL = "Передел";
    const string DATE_FMT = "[$-419]d mmm";          // «18 май»

    static string Norm(object o)
        => string.Join(" ", (Convert.ToString(o) ?? "")
            .Split(new[] { ' ', '\t', '\n', '\r', ' ' }, StringSplitOptions.RemoveEmptyEntries));

    static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("usage: BookBuilder <data.xlsx> <template.xlsx> [out.xlsx]");
            return 2;
        }
        string dataPath = args[0];
        string tplPath = args[1];
        string outPath = args.Length > 2 ? args[2] : dataPath;

        try
        {
            // 1) свежие данные реестра из data-книги (типы уже верные: даты, числа, текст)
            object[,] data;
            using (var dp = new ExcelPackage(new FileInfo(dataPath)))
            {
                var src = dp.Workbook.Worksheets[SHEET];
                if (src?.Dimension == null)
                { Console.Error.WriteLine($"[ERR] нет данных на «{SHEET}» в {dataPath}"); return 3; }
                data = src.Cells[src.Dimension.Address].Value as object[,];
            }
            int nRows = data.GetLength(0), nCols = data.GetLength(1);
            if (nRows < 2) { Console.Error.WriteLine("[ERR] в реестре нет строк"); return 4; }

            // индекс колонки «Передел» по заголовку (для sanity-лога)
            int perCol = -1;
            for (int c = 0; c < nCols; c++)
                if (Norm(data[0, c]) == PERDEL) { perCol = c; break; }
            if (perCol < 0) { Console.Error.WriteLine($"[ERR] нет колонки «{PERDEL}»"); return 6; }

            // ПОЛНЫЙ реестр (все строки, с заголовком)
            var rows = new List<object[]>(nRows);
            for (int r = 0; r < nRows; r++)
            {
                var arr = new object[nCols];
                for (int c = 0; c < nCols; c++) arr[c] = data[r, c];
                rows.Add(arr);
            }
            int nWritten = rows.Count;

            // 2) шаблон со сводной → подменяем данные листа-реестра + «Таблица1»
            using (var tpl = new ExcelPackage(new FileInfo(tplPath)))
            {
                var wb = tpl.Workbook;
                var ws = wb.Worksheets[SHEET];
                if (ws == null) { Console.Error.WriteLine($"[ERR] в шаблоне нет «{SHEET}»"); return 5; }

                // ---- «Сводный_Реестр»: ПОЛНЫЕ данные + «Таблица1» ----
                // LoadFromArrays не перезапишет ячейки существующей умной таблицы — сначала
                // снимаем таблицу, чистим лист, пишем, пересоздаём. EPPlus 4.5.3.3 при этом
                // плодит дубль table-файла — его уберёт Python-пост-шаг.
                if (ws.Tables[TABLE] != null) ws.Tables.Delete(TABLE, false);
                int oldRows = ws.Dimension?.Rows ?? 0;
                if (oldRows > 0) ws.DeleteRow(1, oldRows);
                ws.Cells["A1"].LoadFromArrays(rows);
                var tbl = ws.Tables.Add(ws.Cells[1, 1, nWritten, nCols], TABLE);
                tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

                // формат дат «18 май» (по заголовкам)
                for (int c = 1; c <= nCols; c++)
                {
                    var h = Convert.ToString(ws.GetValue(1, c));
                    if (h == "Дата. Факт" || h == "Дата выдачи наряд-задания")
                        ws.Cells[2, c, nWritten, c].Style.Numberformat.Format = DATE_FMT;
                }

                // Сводную (кэш, записи, фильтр «Передел») НЕ трогаем здесь — её целиком
                // материализует Python-пост-шаг из данных «Таблица1». Источник в шаблоне
                // уже worksheetSource name="Таблица1" — оставляем как есть.

                tpl.SaveAs(new FileInfo(outPath));
            }
            Console.WriteLine($"[OK] книга собрана: реестр {nWritten - 1} строк, {nCols} колонок; сводную материализует пост-шаг");
            return 0;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine($"[ERR] {e.GetType().Name}: {e.Message}");
            return 1;
        }
    }
}
