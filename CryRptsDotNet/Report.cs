using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using System.Windows.Forms;

namespace CryRptsDotNet
{
    public class Report
    {
        ExcelFormatOptions excelOpt;

        public Report()
        {
            excelOpt = new ExcelFormatOptions();
        }

        /// <summary>
        /// Loads a Crystal report from disk in the .rpt format. This is the only way to access reports. This must always be the first method called
        /// </summary>
        /// <param name="filename"></param> Fully qualified path with filename for the desired report
        /// <returns></returns> A newly created ReportDocument object, to pass to either type of plugin. This will be the object that is manipulated.
        public ReportDocument LoadRpt(String filename)
        {
            ReportDocument rpt = new ReportDocument();
            rpt.Load(filename);
            return rpt;
        }


        public String GetTbls(ReportDocument rpt)
        {
            String tableNames = "";
            for (int I = 0; I < rpt.Database.Tables.Count; I++)
            {
                tableNames += rpt.Database.Tables[I].Name;
                if (I != rpt.Database.Tables.Count - 1) { tableNames += ", "; }
            }

            if (rpt.Subreports.Count > 0)
            {
                for (int I = 0; I < rpt.Subreports.Count; I++)
                {
                    Tables subTables = rpt.Subreports[I].Database.Tables;
                    for (int J = 0; J < subTables.Count; J++)
                    {
                        if (tableNames.Contains(subTables[J].Name) == false)
                        {
                            if (J == 0 & tableNames != "") { tableNames += ", "; }

                            tableNames += subTables[J].Name;
                            if (J != subTables.Count - 1) { tableNames += ", "; }
                        }
                    }
                    subTables = null;
                }
            }

            return tableNames;
        }


        public String[] GetFrmlaFlds(ReportDocument rpt)
        {
            int length = rpt.DataDefinition.FormulaFields.Count * 2;
            if (rpt.Subreports.Count > 0)
            {
                for (int I = 0; I < rpt.Subreports.Count; I++)
                {
                    length += rpt.Subreports[I].DataDefinition.FormulaFields.Count * 2;
                }
            }

            String[] FrmlaFlds = new String[length];
            int J = 0;
            for (int I = 0; I < rpt.DataDefinition.FormulaFields.Count; I++)
            {
                FrmlaFlds[J] = rpt.DataDefinition.FormulaFields[I].FormulaName;
                J++;
                FrmlaFlds[J] = rpt.DataDefinition.FormulaFields[I].Text;
                J++;
            }

            if (rpt.Subreports.Count > 0)
            {
                for (int I = 0; I < rpt.Subreports.Count; I++)
                {
                    FormulaFieldDefinitions subFlds = rpt.Subreports[I].DataDefinition.FormulaFields;
                    for (int K = 0; K < subFlds.Count; K++)
                    {
                        FrmlaFlds[J] = subFlds[K].FormulaName;
                        J++;
                        FrmlaFlds[J] = subFlds[K].Text;
                        J++;
                    }
                    subFlds = null;
                }
            }

            return FrmlaFlds;
        }


        public string GetUsedFields(ReportDocument rpt)
        {
            Tables[] rptTables = new Tables[20];
            rptTables[0] = rpt.Database.Tables;
            if (rpt.Subreports.Count > 0)
            {
                for (int I = 0; I < rpt.Subreports.Count; I++)
                {
                    rptTables[I + 1] = rpt.Subreports[I].Database.Tables;
                }
            }

            Dictionary<string, string> fields = new Dictionary<string, string>();

            string fieldStr = "";
            //loop through reports/subreports
            for (int I = 0; I < rpt.Subreports.Count + 1; I++)
            {
                //loop through tables in report/subreport
                for (int J = 0; J < rptTables[I].Count; J++)
                {
                    if (fields.ContainsKey(rptTables[I][J].Name) == false) { fields.Add(rptTables[I][J].Name, rptTables[I][J].Name + ":"); }
                    //loop through fields in a table in a report/subreport
                    for (int K = 0; K < rptTables[I][J].Fields.Count; K++)
                    {
                        if (rptTables[I][J].Fields[K].UseCount > 0)
                        {
                            if (fields[rptTables[I][J].Name].Equals(rptTables[I][J].Name + ":")) 
                            { 
                                fields[rptTables[I][J].Name] = fields[rptTables[I][J].Name] + rptTables[I][J].Fields[K].Name; 
                            }
                            else
                            {
                                if (fields[rptTables[I][J].Name].Contains(rptTables[I][J].Fields[K].Name) == false)
                                {
                                    fields[rptTables[I][J].Name] = fields[rptTables[I][J].Name] + "," + rptTables[I][J].Fields[K].Name;
                                }
                            }

                            //fields += rptTables[I][J].Name + ":" + rptTables[I][J].Fields[K].Name;
                            //if ( !(I==rpt.Subreports.Count & J==rptTables[rpt.Subreports.Count-1].Count-1 & K==rptTables[rpt.Subreports.Count][rptTables[rpt.Subreports.Count-1].Count].Fields.Count-1))
                            //{
                            //    fields += ",";
                            //}
                        }
                    }
                }
            }

            foreach( KeyValuePair<string,string> kvp in fields)
            {
                fieldStr = fieldStr + kvp.Value + ";";
            }

            return fieldStr;
        }


        /// <summary>
        /// Handles connecting to the same data source on all of the tables in a report. Will not work properly if done before LoadRpt. Must be done first for all other methods to work.
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="serverName"></param> For ODBC this equates to the data source
        /// <param name="databaseName"></param> If using an ODBC connection, will have no effect. For SQL it is required.
        /// <param name="userName"></param>
        /// <param name="password"></param>
        public void SetDBLogon(ReportDocument rpt, String serverName, String databaseName, String userName, String password)
        {
            ConnectionInfo connectionInfo = new ConnectionInfo();
            connectionInfo.ServerName = serverName;
            connectionInfo.DatabaseName = databaseName;
            connectionInfo.UserID = userName;
            connectionInfo.Password = password;

            Tables tables = rpt.Database.Tables;
            foreach (Table table in tables)
            {
                TableLogOnInfo newLogonInfo = table.LogOnInfo;
                newLogonInfo.ConnectionInfo = connectionInfo;
                table.ApplyLogOnInfo(newLogonInfo);
            }
        }

        /// <summary>
        /// Overloaded version for connecting to a report that has a Subreport that needs a SQL connection. Will check if any of the SR's need SQL and then connect to all their tables.
        /// </summary>
        /// See above method for first 4 parameters. Last 4 are the same but for the Subreport(s).
        public void SetDBLogon(ReportDocument rpt, String serverName, String databaseName, String userName, String password,
                                                   String serverNameSQL, String databaseNameSQL, String userNameSQL, String passwordSQL)
        {
            int rptTyp = 1;

            ConnectionInfo connectionInfo = new ConnectionInfo();
            Tables tables = rpt.Database.Tables;

            if (rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.ServerName.Contains("CUSTOM") == true)
            {
                //MessageBox.Show("The main report requires a SQL connection.", "Status", MessageBoxButtons.OK);
                rptTyp = 2;
                connectionInfo.ServerName = serverNameSQL;
                connectionInfo.DatabaseName = databaseNameSQL;
                connectionInfo.UserID = userNameSQL;
                connectionInfo.Password = passwordSQL;
            }
            else
            {
                connectionInfo.ServerName = serverName;
                connectionInfo.DatabaseName = databaseName;
                connectionInfo.UserID = userName;
                connectionInfo.Password = password;
            }

            foreach (Table table in tables)
            {
                TableLogOnInfo newLogonInfo = table.LogOnInfo;

                newLogonInfo.ConnectionInfo = connectionInfo;
                table.ApplyLogOnInfo(newLogonInfo);
            }

            Subreports srs = rpt.Subreports;
            foreach (ReportDocument sr in srs)
            {
                bool needLogon = false;
                ConnectionInfo SRConnectionInfo = new ConnectionInfo();
                if ((sr.Database.Tables[0].LogOnInfo.ConnectionInfo.ServerName.Contains("CUSTOM") == true) & (rptTyp == 1))
                {
                    //MessageBox.Show("A subreport requires a SQL connection.", "Status", MessageBoxButtons.OK);
                    if (serverNameSQL != "") SRConnectionInfo.ServerName = serverNameSQL;
                    if (databaseNameSQL != "") SRConnectionInfo.DatabaseName = databaseNameSQL;
                    SRConnectionInfo.UserID = userNameSQL;
                    SRConnectionInfo.Password = passwordSQL;
                    needLogon = true;
                }
                else if (rptTyp == 2)
                {
                    if (serverName != "") SRConnectionInfo.ServerName = serverName;
                    if (databaseName != "") SRConnectionInfo.DatabaseName = databaseName;
                    SRConnectionInfo.UserID = userName;
                    SRConnectionInfo.Password = password;
                    needLogon = true;
                }

                if (needLogon)
                {
                    Tables SRTables = sr.Database.Tables;
                    foreach (Table SRTable in SRTables)
                    {
                        TableLogOnInfo SRLogonInfo = SRTable.LogOnInfo;

                        SRLogonInfo.ConnectionInfo = SRConnectionInfo;
                        SRTable.ApplyLogOnInfo(SRLogonInfo);
                    }
                }
             }
        }

        #region Export

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="dest"></param> Path to destination folder
        /// <param name="name"></param> Desired name for file including file extension
        /// <param name="exType"></param> Export method. Either PDF or EXCEL
        /// <param name="XLSX"></param> Whether to make an EXCEL export an XLSX file. False will generate a XLS file.
        /// <returns></returns> Value NumberOfRecordRead from the report
        public int ExpRpt(ReportDocument rpt, String dest, String name, String exType, String delim, String sep, Boolean XLSX)
        {
            if (!System.IO.Directory.Exists(dest))
            {
                System.IO.Directory.CreateDirectory(dest);
            }
            

            if (exType.ToUpper() == "PDF")
            {
                ExpRptPdf(rpt, dest, name);
            }
            else if (exType.ToUpper() == "EXCEL")
            {
                ExpRptExcel(rpt, dest, name, XLSX);
            }
            else if (exType.ToUpper() == "HTML32")
            {
                ExpRptHTML(rpt, dest, name, "32");   
            }
            else if (exType.ToUpper() == "HTML40")
            {
                ExpRptHTML(rpt, dest, name, "40");
            }
            else if (exType.ToUpper() == "CSV")
            {
                ExpRptCSV(rpt, dest, name, delim, sep);
            }
            return rpt.ReportRequestStatus.NumberOfRecordRead;
        }

        /// <summary>
        /// Exports the current report to a PDF. If the specified destination folder does not exist it will be created
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="dest"></param> Path to destination folder
        /// <param name="name"></param> Desired name for PDF including file extension
        public void ExpRptPdf(ReportDocument rpt, String dest, String name)
        {
            if (!System.IO.Directory.Exists(dest))
            {
                System.IO.Directory.CreateDirectory(dest);
            }
        
            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, dest + name);
        }

        /// <summary>
        /// Exports the current report to a Microsoft Excel document. If the specified destination folder does not exist it will be created
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="dest"></param> Path to destination folder
        /// <param name="name"></param> Desired name for Excel document including file extension
        public void ExpRptExcel(ReportDocument rpt, String dest, String name, Boolean xlsx)
        {
            if (!System.IO.Directory.Exists(dest))
            {
                System.IO.Directory.CreateDirectory(dest);
            }

            ExportOptions exOpt = new ExportOptions();
            DiskFileDestinationOptions dfDestOpt = new DiskFileDestinationOptions();

            if (xlsx)
            {
                exOpt.ExportFormatType = ExportFormatType.ExcelWorkbook;
                if (name.Contains(".xlsx") == false) { name.Replace(".xls", ".xlsx"); }
            }
            else
            {
                exOpt.ExportFormatType = ExportFormatType.Excel;
                if (name.Contains(".xlsx") == true) { name.Replace(".xlsx", ".xls"); }
            }
            dfDestOpt.DiskFileName = dest + name;

            exOpt.ExportDestinationType = ExportDestinationType.DiskFile;
            exOpt.ExportDestinationOptions = dfDestOpt;

            exOpt.ExportFormatOptions = excelOpt;

            rpt.Export(exOpt);
        }

        /// <summary>
        /// Exports the current report to an HTML document
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="dest"></param> Path to destination folder
        /// <param name="name"></param> Desired name for HTML document including file extension
        /// <param name="htmlVer"></param> specify whether you want to use HTML 3.2 or HTML 4
        public void ExpRptHTML(ReportDocument rpt, String dest, String name, String htmlVer)
        {
            ExportFormatType expForm = new ExportFormatType();

            if (htmlVer == "32")
            {
                expForm = ExportFormatType.HTML32; 
            }
            else
            {
                expForm = ExportFormatType.HTML40;    
            }
            
            rpt.ExportToDisk(expForm, dest + name);
        }

        public void ExpRptCSV(ReportDocument rpt, String dest, String name, String delim, String sep)
        {
            DiskFileDestinationOptions diskOpt = new DiskFileDestinationOptions
            {
                DiskFileName = dest + name
            };

            CharacterSeparatedValuesFormatOptions csvOpt = new CharacterSeparatedValuesFormatOptions
            {
                Delimiter = delim,
                SeparatorText = sep,
                ExportMode = CsvExportMode.Standard,
                GroupSectionsOption = CsvExportSectionsOption.DoNotExport,
                ReportSectionsOption = CsvExportSectionsOption.ExportIsolated
            };

            ExportOptions expOpt = new ExportOptions
            {
                ExportDestinationType = ExportDestinationType.DiskFile,
                ExportDestinationOptions = diskOpt,
                ExportFormatType = ExportFormatType.CharacterSeparatedValues,
                ExportFormatOptions = csvOpt
            };

            rpt.Export(expOpt);
        }

        /// <summary>
        /// Sets the ExcelAreaGroupNumber property in the ExcelFormatOptions object. This should be set before ExpRptExcel is called to take effect.
        /// </summary>
        /// <param name="arGrpNum"></param> Desired value
        public void SetArGrpNum(short arGrpNum)
        {
            excelOpt.ExcelAreaGroupNumber = arGrpNum;
        }

        /// <summary>
        /// Sets the ExcelAreaType property in the ExcelFormatOptions object. This should be set before ExpRptExcel is called to take effect.
        /// </summary>
        /// <param name="areaTyp"></param> Desired value as integer. The field looks for an enumerated type so the translation is done here.
        public void SetAreaTyp(int areaTyp)
        {
            switch (areaTyp)
            {
                case 0:
                    excelOpt.ExcelAreaType = AreaSectionKind.Invalid;
                    break;
                case 1:
                    excelOpt.ExcelAreaType = AreaSectionKind.ReportHeader;
                    break;
                case 2:
                    excelOpt.ExcelAreaType = AreaSectionKind.PageHeader;
                    break;
                case 3:
                    excelOpt.ExcelAreaType = AreaSectionKind.GroupHeader;
                    break;
                case 4:
                    excelOpt.ExcelAreaType = AreaSectionKind.Detail;
                    break;
                case 5:
                    excelOpt.ExcelAreaType = AreaSectionKind.GroupFooter;
                    break;
                case 7:
                    excelOpt.ExcelAreaType = AreaSectionKind.PageFooter;
                    break;
                case 8:
                    excelOpt.ExcelAreaType = AreaSectionKind.ReportFooter;
                    break;
                case 255:
                    excelOpt.ExcelAreaType = AreaSectionKind.WholeReport;
                    break;
            }
        }

        /// <summary>
        /// Sets the ExcelTabHasColumnHeadings property in the ExcelFormatOptions object. This should be set before ExpRptExcel is called to take effect.
        /// </summary>
        /// <param name="tabHsColHead"></param> Desired value
        public void SetTabHsColHead(Boolean tabHsColHead)
        {
            excelOpt.ExcelTabHasColumnHeadings = tabHsColHead;
        }

        /// <summary>
        /// Sets the ExcelUseConstantColumnWidth property in the ExcelFormatOptions object. This should be set before ExpRptExcel is called to take effect.
        /// </summary>
        /// <param name="useConstColW"></param> Desired value
        public void SetUseConstColW(Boolean useConstColW)
        {
            excelOpt.ExcelUseConstantColumnWidth = useConstColW;
        }

        #endregion

        #region Print

        /// <summary>
        /// Prints the current report to the default system printer, if a different printer has not been set already. Only prints one copy, prints whole document, and is not collated
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        public void PrintOut(ReportDocument rpt)
        {
            if (rpt.PrintOptions.PrinterName == "")
            {
                PrintDocument printDocument = new PrintDocument();
                rpt.PrintOptions.PrinterName = printDocument.PrinterSettings.PrinterName;
            }
            rpt.PrintToPrinter(1, false, 0, 0);
        }

        /// <summary>
        /// Sets the printer to be used to print the report. Must be called before PrintOut is called to have an effect.
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="printerName"></param> Name of desired printer. Must be precise. If incorrect nothing will happen.
        public void SelPrinter(ReportDocument rpt, String printerName)
        {
            PrinterSettings.StringCollection printers = PrinterSettings.InstalledPrinters;
            foreach (String s in printers)
                if (printerName == s) rpt.PrintOptions.PrinterName = printerName;
        }

        public void SetPrtSize(ReportDocument rpt, String sz)
        {
            string size = sz.ToLower();
            if (size == "legal") rpt.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperLegal;
            else if (size == "letter") rpt.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperLetter;
            else if (size == "default") rpt.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize;
        }

        public void SetPrtOrient(ReportDocument rpt, String orient)
        {
            string orientation = orient.ToLower();
            if (orientation == "portrait") rpt.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait;
            else if (orientation == "landscape") rpt.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape;
            else if (orientation == "default") rpt.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation;
        }

        #endregion

        #region Parameters

        /// <summary>
        /// Either adds to or replaces the Record Selection Formula for the report.
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="selFormula"></param> String containing selection formula. If invalid the Crystal Reports engine will throw an error.
        /// <param name="ow"></param> True or false that determines whether to overwrite current value
        public void SetRecSelFormula(ReportDocument rpt, String selFormula, Boolean ow)
        {
            if (ow)
            {
                rpt.DataDefinition.RecordSelectionFormula = selFormula;
            }
            else
            {
                rpt.DataDefinition.RecordSelectionFormula += " " + selFormula;
            }            
        }

        /// <summary>
        /// Clear the report parameter corresponding to the specified name paramName
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="paramName"></param>
        public void ClearParam(ReportDocument rpt, String paramName)
        {
            if (rpt.ParameterFields[paramName] != null) rpt.ParameterFields[paramName].CurrentValues.Clear();
            else throw new System.Exception("No parameter with that name exists");
        }

        /// <summary>
        /// Clear all report parameters
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        public void ClearAllParams(ReportDocument rpt)
        {
            /*
            if (rpt.ParameterFields.Count > 0)
            {
                ParameterFields paramFlds = rpt.ParameterFields;
                foreach (ParameterField field in paramFlds)
                {

                    field.CurrentValues.Clear();
                }
            }
            */
            
            if (rpt.DataDefinition.ParameterFields.Count > 0)
            {
                ParameterFieldDefinitions paramFlds = rpt.DataDefinition.ParameterFields;
                foreach (ParameterFieldDefinition field in paramFlds)
                {
                    if (field.IsLinked() == false)
                    {
                        field.CurrentValues.Clear();
                    }
                }
            }

            
        }

        /// <summary>
        /// Set the value for the specified Discrete type parameter. Always overwrites previous value
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="paramName"></param> Name of parameter
        /// <param name="val"></param> Desired value passed as a string
        public void SetDiscVal(ReportDocument rpt, String paramName, String val)
        {
            if (rpt.ParameterFields[paramName] != null) rpt.SetParameterValue(paramName, val);
            else throw new System.Exception("No parameter with that name exists");
        }

        /// <summary>
        /// Set multiple values for the specified Discrete type parameter. Always overwrites previous values
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="paramName"></param> Name of parameter. Throws exception if cant be found
        /// <param name="valList"></param> A string of comma ',' delimited parameter values
        public void SetDiscVals(ReportDocument rpt, String paramName, String valList)
        {

            if (rpt.ParameterFields[paramName] != null)
            {
                char[] sep = new char[] { ',' };
                string[] arr = valList.Split(sep);
                rpt.SetParameterValue(paramName, arr);
            }
            else throw new System.Exception("No parameter with that name exists");
        }

        /// <summary>
        /// Set a specified Range type parameter.
        /// </summary>
        /// <param name="rpt"></param> ReportDocument object from the plugin
        /// <param name="paramName"></param> Name of parameter. Throws exception if cant be found
        /// <param name="low"></param> Lower bound of range
        /// <param name="hi"></param> Upper bound of range
        /// <param name="bnds"></param> integer value determining bound inclusion
        public void SetRangeVal(ReportDocument rpt, String paramName, String low, String hi, int bnds)
        {
            if (rpt.ParameterFields[paramName] != null)
            {
                ParameterRangeValue rangeValue = new ParameterRangeValue();
                rangeValue.StartValue = low;
                rangeValue.EndValue = hi;

                if (bnds == 0)
                {
                    rangeValue.LowerBoundType = RangeBoundType.BoundExclusive;
                    rangeValue.UpperBoundType = RangeBoundType.BoundExclusive;
                }
                else if (bnds == 1)
                {
                    rangeValue.LowerBoundType = RangeBoundType.BoundExclusive;
                    rangeValue.UpperBoundType = RangeBoundType.BoundInclusive;
                }
                else if (bnds == 2)
                {
                    rangeValue.LowerBoundType = RangeBoundType.BoundInclusive;
                    rangeValue.UpperBoundType = RangeBoundType.BoundExclusive;
                }
                else if (bnds == 12 | bnds == 21)
                {
                    rangeValue.LowerBoundType = RangeBoundType.BoundInclusive;
                    rangeValue.UpperBoundType = RangeBoundType.BoundInclusive;
                }
                else if (bnds == 4)
                {
                    rangeValue.LowerBoundType = RangeBoundType.BoundExclusive;
                    rangeValue.UpperBoundType = RangeBoundType.NoBound;
                }
                else if (bnds == 8)
                {
                    rangeValue.LowerBoundType = RangeBoundType.NoBound;
                    rangeValue.UpperBoundType = RangeBoundType.BoundExclusive;
                }
                else if (bnds == 48 | bnds == 84)
                {
                    rangeValue.LowerBoundType = RangeBoundType.NoBound;
                    rangeValue.UpperBoundType = RangeBoundType.NoBound;
                }
                rpt.SetParameterValue(paramName, rangeValue);
            }
            else throw new Exception("No parameter with that name exists");
        }


        #endregion
    }
}
