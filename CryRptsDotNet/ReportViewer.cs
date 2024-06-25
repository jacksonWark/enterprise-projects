using System;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using RemObjects.Hydra;

namespace CryRptsDotNet
{
    [Plugin, VisualPlugin]
    public partial class ReportViewer : VisualPlugin, IReportViewer, IReportManager
    {
        public ReportDocument report;
        public Report rptObj;

        public ReportViewer()
        {
            InitializeComponent();
            rptObj = new Report();
            crystalReportViewer.ToolPanelView = ToolPanelViewType.None;
        }



        public void LoadReport(String filepath)
        {
            report = rptObj.LoadRpt(filepath);
        }


        public void CleanupReport()
        {
            report.Close();
            report.Dispose();
        }


        public String GetTables()
        {
            return rptObj.GetTbls(report);
        }


        public String[] GetFormulaFields()
        {
            return rptObj.GetFrmlaFlds(report);
        }


        public String GetUsedFields()
        {
            return rptObj.GetUsedFields(report);
        }


        public void SetDatabaseLogon(String serverName, String databaseName, String userName, String password)
        {
            rptObj.SetDBLogon(report, serverName, databaseName, userName, password);
        }

        public void SetDatabaseLogon(String serverName, String databaseName, String userName, String password,
                                     String SRServerName, String SRDatabaseName, String SRUserName, String SRPassword)
        {
            rptObj.SetDBLogon(report, serverName, databaseName, userName, password, SRServerName, SRDatabaseName, SRUserName, SRPassword);
        }


        /// <summary>
        /// Binds the ReportDocument object to the visual component
        /// </summary>
        public void LinkDisplay()
        {
            crystalReportViewer.ReportSource = report;
        }

        #region Export

        public int ExportReport(String destination, String exportName, String exportType, String delim = "\"", String sep = ",", Boolean XLSX = false)
        {
            return rptObj.ExpRpt(report, destination, exportName, exportType, delim, sep, XLSX);
        }

        public void ExportReportPdf(String destination, String exportName)
        {
            // TODO - Implement this better
            rptObj.ExpRptPdf(report, destination, exportName);
        }

        public void ExportReportExcel(String destination, String exportName, Boolean XLSX)
        {
            // TODO - Implement this better
            rptObj.ExpRptExcel(report, destination, exportName, XLSX);
        }

        public void SetAreaGroupNumber(int areaGroupNumber) { rptObj.SetArGrpNum((short)(areaGroupNumber)); }
        public void SetAreaType(int areaType) { rptObj.SetAreaTyp(areaType); }
        public void SetTabHasColHeadings(Boolean tabHasColHeadings) { rptObj.SetTabHsColHead(tabHasColHeadings); }
        public void SetUseConstantColWidth(Boolean useConstantColWidth) { rptObj.SetUseConstColW(useConstantColWidth); }

        #endregion

        #region Print

        public void PrintOut()
        {
            rptObj.PrintOut(report);
        }

        public void SelectPrinter(String printerName)
        {
            rptObj.SelPrinter(report, printerName);
        }

        public void SetPrintSize(String size)
        {
            rptObj.SetPrtSize(report, size);
        }

        public void SetPrintOrientation(String orientation)
        {
            rptObj.SetPrtOrient(report, orientation);
        }

        #endregion

        #region Parameters

        public void SetRecordSelectionFormula(String selectionForumla, Boolean overwrite)
        {
            rptObj.SetRecSelFormula(report, selectionForumla, overwrite);
        }

        public void ClearParameter(String parameterName)
        {
            rptObj.ClearParam(report, parameterName);
        }

        public void ClearAllParameters()
        {
            rptObj.ClearAllParams(report);
        }

        public void SetDiscreteValue(String parameterName, String value)
        {
            rptObj.SetDiscVal(report, parameterName, value);
        }

        public void SetDiscreteValues(String parameterName, String values)
        {
            rptObj.SetDiscVals(report, parameterName, values);
        }

        public void SetRangeValue(String parameterName, String lowVal, String hiVal, int bounds)
        {
            rptObj.SetRangeVal(report, parameterName, lowVal, hiVal, bounds);
        }

        #endregion

    }
}
