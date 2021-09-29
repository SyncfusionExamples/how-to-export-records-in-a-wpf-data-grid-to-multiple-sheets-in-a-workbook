using Syncfusion.Data;
using Syncfusion.UI.Xaml.Grid;
using Syncfusion.UI.Xaml.Grid.Converter;
using Syncfusion.XlsIO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SfDataGrid_MVVM
{
    public class CustomSfDataGridToExcelConverter : SfDataGridToExcelConverter
    {
        int NumberOfRecordsPerSheet { get; set; } = 1000;


        public override ExcelEngine ExportToExcel(SfDataGrid grid, ICollectionViewAdv view, ExcelExportingOptions excelExportingOptions)
        {
            ExcelEngine engine = new ExcelEngine();
            IWorkbook workbook = engine.Excel.Workbooks.Create();
            IWorksheet sheet;
            if (view == null)
                return engine;
            if (grid.DetailsViewDefinition.Count > 0)
                excelExportingOptions.AllowOutlining = true;

            bool exportAllPages = excelExportingOptions.ExportAllPages;

            workbook.Version = excelExportingOptions.ExcelVersion;
            excelExportingOptions.GetType().GetProperty("GroupColumnDescriptionsCount", System.Reflection.BindingFlags.NonPublic
                | System.Reflection.BindingFlags.Instance).SetValue(excelExportingOptions, view.GroupDescriptions.Count);

            var columns = (from column in grid.Columns where !excelExportingOptions.ExcludeColumns.Contains(column.MappingName) select column.MappingName).ToList();

            excelExportingOptions.GetType().GetField("columns", System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Instance).SetValue(excelExportingOptions, columns);


            sheet = workbook.Worksheets[0];
            ExportToExcelWorksheet(grid, view, sheet, excelExportingOptions);

            return engine;
        }

        protected override void ExportRecordsToExcel(SfDataGrid grid, IWorksheet sheet, ExcelExportingOptions excelExportingOptions,
            IEnumerable records, IPropertyAccessProvider propertyAccessProvider, Group group)
        {
            bool hasrecords = false;
            foreach (var rec in records)
            {
                hasrecords = true;
                break;
            }
            if (!hasrecords)
                return;

            var gridColumns = excelExportingOptions.Columns;

            if (grid.DetailsViewDefinition.Count == 0)
            {
                ObservableCollection<object> splitedRecords = new ObservableCollection<object>();
                int sheetCount = 0;
                int recordCount = 0;

                int remainder = grid.View.Records.Count % NumberOfRecordsPerSheet;
                int numberOfSheets = (grid.View.Records.Count / NumberOfRecordsPerSheet);

                if (remainder > 0)
                    numberOfSheets++;

                foreach (var rec in records)
                {
                    recordCount++;
                    splitedRecords.Add(rec);

                    if (splitedRecords.Count >= NumberOfRecordsPerSheet)
                    {
                        if (sheet.Workbook.Worksheets.Count < numberOfSheets)
                            sheet.Workbook.Worksheets.Create();
                        base.ExportRecordsToExcel(grid, sheet.Workbook.Worksheets[sheetCount], excelExportingOptions, splitedRecords, propertyAccessProvider, group);

                        this.GetType().BaseType.GetField("excelRowIndex", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(this, 0);

                        sheetCount++;
                        splitedRecords.Clear();

                    }

                }

                if (splitedRecords.Count != 0)
                {
                    if (sheet.Workbook.Worksheets.Count < numberOfSheets)
                        sheet.Workbook.Worksheets.Create();
                    base.ExportRecordsToExcel(grid, sheet.Workbook.Worksheets[sheetCount], excelExportingOptions, splitedRecords, propertyAccessProvider, group);
                }
            }

        }
    }
}
