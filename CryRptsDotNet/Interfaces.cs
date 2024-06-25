using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RemObjects.Hydra.CrossPlatform;
using System.Runtime.InteropServices;


[Guid("48457700-7b7a-11ea-ab12-0800200c9a66")]
public interface IReportViewer : IHYCrossPlatformInterface
{
    void LinkDisplay();
}

    
[Guid("3258cfc0-7cfe-11ea-ab12-0800200c9a66")]
public interface IReportManager : IHYCrossPlatformInterface
{
    void LoadReport(String filepath);

    void CleanupReport();

    String GetTables();

    String[] GetFormulaFields();

    String GetUsedFields();

    void SetDatabaseLogon(String serverName, String databaseName, String userName, String password);
    void SetDatabaseLogon(String serverName, String databaseName, String userName, String password, String serverNameSQL, String databaseNameSQL, String userNameSQL, String passwordSQL);

    void SetRecordSelectionFormula(String selectionFormula, Boolean overwrite);
    void ClearParameter(String parameterName);
    void ClearAllParameters();
    void SetDiscreteValue(String parameterName, String value);
    void SetDiscreteValues(String parameterName, String values);
    void SetRangeValue(String parameterName, String lowVal, String hiVal, int bounds);

    int ExportReport(String destination, String exportName, String exportType, String delim = "\"", String sep = ",", Boolean XLSX = false);
    void ExportReportPdf(String destination, String exportName);
    void ExportReportExcel(String destination, String exportName, Boolean XLSX);
    void SetAreaGroupNumber(int areaGroupNumber);
    void SetAreaType(int areaType);
    void SetTabHasColHeadings(Boolean tabHasColHeadings);
    void SetUseConstantColWidth(Boolean useConstantColWidth);

    void PrintOut();
    void SelectPrinter(String printerName);
    void SetPrintSize(String size);
    void SetPrintOrientation(String orientation);

}
