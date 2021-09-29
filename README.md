# How to export records in a WPF DataGrid (SfDataGrid) to multiple worksheets in a workbook?

# About the sample

This example illustrates how to export records in a WPF DataGrid (SfDataGrid) to multiple worksheets in a workbook.

[WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid) provides support to export the data to Excel by using the [ExportToExcel](https://help.syncfusion.com/cr/wpf/Syncfusion.UI.Xaml.Grid.Converter.GridExcelExportExtension.html#Syncfusion_UI_Xaml_Grid_Converter_GridExcelExportExtension_ExportToExcel_Syncfusion_UI_Xaml_Grid_SfDataGrid_Syncfusion_Data_ICollectionViewAdv_Syncfusion_UI_Xaml_Grid_Converter_ExcelExportingOptions_) method. By default, all the records will be exported to a single worksheet of the exported excel workbook. It is possible to split the total records and export them to different worksheets by creating custom `SfDataGridToExcelConverter`.

```C#

private void Button_Click(object sender, RoutedEventArgs e)
{
    var options = new ExcelExportingOptions();
    options.ExcelVersion = ExcelVersion.Excel2016;
    var excelEngine = new CustomSfDataGridToExcelConverter().ExportToExcel(this.dataGrid, dataGrid.View, options);
    var workBook = excelEngine.Excel.Workbooks[0];
    workBook.SaveAs("Sample.xlsx");
}


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

```

Take a moment to peruse the [documentation](https://help.syncfusion.com/wpf/datagrid/export-to-excel), where you can find about export to excel feature in SfDataGrid, with code examples.

## Requirements to run the demo

Visual Studio 2015 and above versions.

