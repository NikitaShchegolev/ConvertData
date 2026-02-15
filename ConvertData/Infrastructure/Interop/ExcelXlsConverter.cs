using System;

namespace ConvertData.Infrastructure.Interop;

internal static class ExcelXlsConverter
{
    public static void ConvertXlsToXlsxViaExcel(string xlsPath, string xlsxPath)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new PlatformNotSupportedException("Conversion from .xls requires Microsoft Excel installed (Excel.Application COM is not available).");

        object? excel = null;
        object? workbooks = null;
        object? workbook = null;

        try
        {
            excel = Activator.CreateInstance(excelType);
            excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });

            workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excel, Array.Empty<object>());
            var workbooksType = workbooks!.GetType();

            workbook = workbooksType.InvokeMember(
                "Open",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                workbooks,
                new object[] { xlsPath }
            );

            workbook!.GetType().InvokeMember(
                "SaveAs",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                workbook,
                new object[] { xlsxPath, 51 }
            );

            workbook.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, workbook, new object[] { false });
            workbook = null;

            excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excel, Array.Empty<object>());
        }
        finally
        {
            try
            {
                if (workbook != null)
                    workbook.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, workbook, new object[] { false });
            }
            catch { }

            try
            {
                if (excel != null)
                    excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excel, Array.Empty<object>());
            }
            catch { }

            if (workbook != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
            if (workbooks != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbooks);
            if (excel != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel);
        }
    }
}
