using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Dapper;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Npgsql;

namespace ExportPostgresqlTableSchema
{
    class Program
    {
        private static string ConnectionString { get; set; }
        private static ILogger<Program> _logger { get; set; }

        static async Task Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                            .SetBasePath(Directory.GetCurrentDirectory())
                            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            var Configuration = builder.Build();
            ConnectionString = Configuration.GetConnectionString("TargetDb");

            //setup our DI
            var services = new ServiceCollection()
                            .AddLogging(configure => configure.AddConsole());

            var serviceProvider = services.BuildServiceProvider();

            //configure console logging
            serviceProvider.GetService<ILoggerFactory>();
            serviceProvider.GetService<IConfiguration>();

            _logger = serviceProvider.GetService<ILoggerFactory>()
                .CreateLogger<Program>();
            _logger.LogInformation("Starting application");

            //do the actual work here
            var (ms, attachmentName) = WriteCSV();
            ms.Seek(0, SeekOrigin.Begin);
            using (FileStream file = new(attachmentName, FileMode.Create, FileAccess.ReadWrite))
            {
                byte[] bytes = new byte[ms.Length];
                ms.Read(bytes, 0, (int)ms.Length);
                file.Write(bytes, 0, bytes.Length);
                ms.Close();
            }

            _logger.LogInformation("All done!");
        }

        private static (MemoryStream, string) WriteCSV()
        {
            var filename = $"Table Schemas.xlsx";
            MemoryStream ms = new();
            try
            {
                var data = QueryTableSchemaModel()
                        .GroupBy(g => new { g.table_name, g.table_description })
                        .Select(s => new
                        {
                            tableName = s.Key.table_name,
                            tableDescription = s.Key.table_description,
                            rows = s.Select(ss => new
                            {
                                ss.column_name,
                                ss.column_description,
                                ss.ordinal_position,
                                ss.data_type,
                                ss.character_maximum_length,
                                ss.is_nullable,
                                ss.default_value,
                                ss.constraint_type,
                                ss.foreign_table_name,
                                ss.foreign_column_name
                            }).ToList()
                        }).ToList();
                _logger.LogInformation($"讀取資料 成功");

                // Create a spreadsheet document by supplying the filepath.
                // By default, AutoSave = true, Editable = true, and Type = xlsx.
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    // Add a WorkbookPart to the document.
                    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    // Adding Style
                    var sp = workbookpart.AddNewPart<WorkbookStylesPart>();
                    sp.Stylesheet = AddStyles();
                    sp.Stylesheet.Save();

                    // Add Sheets to the Workbook.
                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                        AppendChild<Sheets>(new Sheets());

                    for (int i = 0; i < data.Count; i++)
                    {
                        // Add a WorksheetPart to the WorkbookPart.
                        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        Worksheet worksheet = new();
                        SheetData sheetData = new();

                        // Append a new worksheet and associate it with the workbook.
                        Sheet sheet = new()
                        {
                            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                            SheetId = (uint)(i + 1),
                            Name = data[i].tableName
                        };
                        sheets.AppendChild(sheet);

                        var row = new Row() { RowIndex = 2 };
                        row.Append(
                            new Cell()
                            {
                                CellValue = new CellValue("中文table名稱"),
                                CellReference = $"B{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell() { CellReference = $"C{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"D{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"E{row.RowIndex}", StyleIndex = 1 },
                            new Cell()
                            {
                                CellValue = new CellValue("table name"),
                                CellReference = $"F{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell() { CellReference = $"G{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"H{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"I{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"J{row.RowIndex}", StyleIndex = 1 }
                        );
                        sheetData.AppendChild(row);

                        row = new Row() { RowIndex = 3 };
                        row.Append(
                            new Cell()
                            {
                                CellValue = new CellValue(data[i].tableDescription),
                                CellReference = $"B{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell() { CellReference = $"C{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"D{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"E{row.RowIndex}", StyleIndex = 1 },
                            new Cell()
                            {
                                CellValue = new CellValue(data[i].tableName),
                                CellReference = $"F{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell() { CellReference = $"G{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"H{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"I{row.RowIndex}", StyleIndex = 1 },
                            new Cell() { CellReference = $"J{row.RowIndex}", StyleIndex = 1 }
                        );
                        sheetData.AppendChild(row);

                        row = new Row() { RowIndex = 4 };
                        row.Append(
                            new Cell()
                            {
                                CellValue = new CellValue("項次"),
                                CellReference = $"B{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("欄位名稱"),
                                CellReference = $"C{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("欄位中文名稱"),
                                CellReference = $"D{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("型態"),
                                CellReference = $"E{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("LENGTH"),
                                CellReference = $"F{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("NULL?"),
                                CellReference = $"G{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("預設值"),
                                CellReference = $"H{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("CONSTRANT TYPE"),
                                CellReference = $"I{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            },
                            new Cell()
                            {
                                CellValue = new CellValue("備註說明"),
                                CellReference = $"J{row.RowIndex}",
                                DataType = CellValues.String,
                                StyleIndex = 1
                            }
                        );
                        sheetData.AppendChild(row);

                        for (int j = 0; j < data[i].rows.Count; j++)
                        {
                            row = new Row() { RowIndex = (uint)(5 + j) };
                            row.Append(
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].ordinal_position.GetValueOrDefault()),
                                    CellReference = $"B{row.RowIndex}",
                                    DataType = CellValues.Number,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].column_name),
                                    CellReference = $"C{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].column_description),
                                    CellReference = $"D{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].data_type),
                                    CellReference = $"E{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].character_maximum_length.GetValueOrDefault()),
                                    CellReference = $"F{row.RowIndex}",
                                    DataType = CellValues.Number,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].is_nullable),
                                    CellReference = $"G{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].default_value),
                                    CellReference = $"H{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].constraint_type),
                                    CellReference = $"I{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                },
                                new Cell()
                                {
                                    CellValue = new CellValue(data[i].rows[j].column_description),
                                    CellReference = $"J{row.RowIndex}",
                                    DataType = CellValues.String,
                                    StyleIndex = 1
                                }
                            );
                            sheetData.AppendChild(row);
                        }

                        worksheet.Append(sheetData);

                        //create a MergeCells class to hold each MergeCell
                        MergeCells mergeCells = new();

                        //append a MergeCell to the mergeCells for each set of merged cells
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B2:E2") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("F2:J2") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B3:E3") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("F3:J3") });

                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());

                        worksheetPart.Worksheet = worksheet;
                        worksheetPart.Worksheet.Save();
                    }
                    workbookpart.Workbook.Save();
                }

                _logger.LogInformation($"產生Excel 成功");
            }
            catch (Exception ex)
            {
                _logger.LogError($"錯誤 : {ex}");
            }
            return (ms, filename);
        }

        public static Stylesheet AddStyles()
        {
            Fonts fonts = new(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }
                ),
                new Font( // Index 1 - Title
                    new FontSize() { Val = 12 },
                    new Bold()
                    ));

            Fills fills = new(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "f6d8d8" } })
                    { PatternType = PatternValues.Solid }), // Index 1 - 壞掉的點點
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "ebf1de" } })
                    { PatternType = PatternValues.Solid }), // Index 2 - value
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "dce6f1" } })
                    { PatternType = PatternValues.Solid }), // Index 3 - formula
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "f6d8d8" } })
                    { PatternType = PatternValues.Solid }) // Index 4 - total
                );

            Borders borders = new(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            Alignment alignment = new()
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            };

            Alignment alignmentWrapText = new()
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center,
                WrapText = true
            };

            Alignment alignmentTitle = new()
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Top
            };

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, Alignment = (Alignment)alignment.CloneNode(true), ApplyBorder = true, ApplyAlignment = true }, // 1 - normal
                    new CellFormat { FontId = 0, FillId = 4, BorderId = 1, NumberFormatId = 3, ApplyNumberFormat = true, Alignment = (Alignment)alignment.CloneNode(true), ApplyBorder = true, ApplyFill = true, ApplyAlignment = true }, // 2 - total
                    new CellFormat { FontId = 0, FillId = 2, BorderId = 1, NumberFormatId = 3, ApplyNumberFormat = true, Alignment = (Alignment)alignment.CloneNode(true), ApplyBorder = true, ApplyFill = true, ApplyAlignment = true }, // 3 - value
                    new CellFormat { FontId = 0, FillId = 3, BorderId = 1, NumberFormatId = 3, ApplyNumberFormat = true, Alignment = (Alignment)alignment.CloneNode(true), ApplyBorder = true, ApplyFill = true, ApplyAlignment = true }, // 4 - formula(numbers)
                    new CellFormat { FontId = 0, FillId = 3, BorderId = 1, NumberFormatId = 10, ApplyNumberFormat = true, Alignment = (Alignment)alignment.CloneNode(true), ApplyBorder = true, ApplyFill = true, ApplyAlignment = true }, // 5 - formula(percentage)
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, Alignment = (Alignment)alignmentWrapText.CloneNode(true), ApplyBorder = true, ApplyFill = true, ApplyAlignment = true },   // 6 - cell of print date
                    new CellFormat { FontId = 1, FillId = 0, BorderId = 0, Alignment = (Alignment)alignmentTitle.CloneNode(true), ApplyFont = true, ApplyAlignment = true }, // 7 - title
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 3, ApplyNumberFormat = true, Alignment = (Alignment)alignmentTitle.CloneNode(true), ApplyFont = true, ApplyAlignment = true } // 8 - numbers with comma
                );

            return new Stylesheet(fonts, fills, borders, cellFormats);
        }

        private static IEnumerable<TableSchemaModel> QueryTableSchemaModel()
        {

            using (IDbConnection connection = new NpgsqlConnection(ConnectionString))
            {
                return connection.Query<TableSchemaModel>($@"
                    SET enable_nestloop=0;
                    SELECT 
                    -- t.table_catalog,
                    -- t.table_schema,
                    t.table_name,
                    obj_description((t.table_schema||'.'||quote_ident(t.table_name))::regclass) AS table_description,
                    c.column_name,
                    col_description((t.table_schema||'.'||quote_ident(t.table_name))::regclass::oid, c.ordinal_position) AS column_description,
                    c.ordinal_position,
                    c.data_type,
                    c.character_maximum_length,
                    n.constraint_type,
                    c.column_default AS default_value,
                    c.is_nullable,
                    k2.table_name AS foreign_table_name,
                    k2.column_name AS foreign_column_name
                    FROM information_schema.tables t NATURAL
                    LEFT JOIN information_schema.columns c LEFT JOIN(information_schema.key_column_usage k NATURAL
                    JOIN information_schema.table_constraints n NATURAL
                    LEFT JOIN information_schema.referential_constraints r)ON c.table_catalog=k.table_catalog
                            AND c.table_schema=k.table_schema
                            AND c.table_name=k.table_name
                            AND c.column_name=k.column_name
                    LEFT JOIN information_schema.key_column_usage k2
                        ON k.position_in_unique_constraint=k2.ordinal_position
                            AND r.unique_constraint_catalog=k2.constraint_catalog
                            AND r.unique_constraint_schema=k2.constraint_schema
                            AND r.unique_constraint_name=k2.constraint_name
                    WHERE t.TABLE_TYPE='BASE TABLE'
                            AND t.table_schema NOT IN('information_schema','pg_catalog')
                    ORDER BY t.table_name, c.ordinal_position;");
            }
        }
    }
}
