using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;
using System.Drawing.Printing;

namespace OrderBooking_LaxmiSteel
{
    public class ClsCRDisplay
    {

        private static string _Server;
        private static string _Database;
        private static string _UserID;
        private static string _Password;
        private static SqlConnection _SQLCon;
        private static string _SQLQuery;

        #region PROPERTY DECLARATIONS

        public static SqlConnection SQLConnection
        {
            set { _SQLCon = value; }
        }

        public static string SQLServerName
        {
            set { _Server = value; }
        }

        public static string Database
        {
            set { _Database = value; }
        }

        public static string UserID
        {
            set { _UserID = value; }
        }

        public static string Password
        {
            set { _Password = value; }
        }

        #endregion

        #region SIMPLE REPORT HANDLING

        private static void ReportConnectionInfo(CrystalDecisions.CrystalReports.Engine.ReportDocument oReport)
        {
            ConnectionInfo connection = new ConnectionInfo();
            connection.DatabaseName = _Database;
            connection.ServerName = _Server;
            connection.UserID = _UserID;
            connection.Password = _Password;
            foreach (CrystalDecisions.CrystalReports.Engine.Table table in oReport.Database.Tables)
            {
                TableLogOnInfo logOnInfo = table.LogOnInfo;
                logOnInfo.ConnectionInfo = connection;
                table.ApplyLogOnInfo(logOnInfo);
            }
        }

        public static ReportDocument GetReport(DataSet objDataSet, string ReportPath, bool ReportHeaders, string PaperSize, CrystalDecisions.Shared.PaperOrientation Orientation)
        {
            try
            {
                ReportDocument oReportDoc = new ReportDocument();
                oReportDoc.Load(ReportPath);
                ReportConnectionInfo(oReportDoc);
                //clsCRDisplay.ReportConnectionInfo(oReportDoc);
                if (objDataSet.Tables[0].Rows.Count == 0)
                {
                    throw new Exception("No, data available to view.");
                    return null;
                }
                else
                {
                    #region REPORT HEADERS
                    if (ReportHeaders)
                    {
                        _SQLQuery = "Select * from dbo.GetCompany('L')";
                        string ConString = "Server=" + _Server + ";Database=" + _Database + ";UID=" + _UserID + ";PWD=" + _Password;
                        DataTable dt = SqlHelper.ExecuteDataset(ConString, CommandType.Text, _SQLQuery).Tables[0];
                        string CompanyName = dt.Rows[0]["CompName"].ToString();
                        string CompanyAddr = dt.Rows[0]["Address"].ToString();

                        oReportDoc.DataDefinition.FormulaFields["CompanyName"].Text = "'" + CompanyName + "'";
                        oReportDoc.DataDefinition.FormulaFields["CompanyAddrs"].Text = "'" + CompanyAddr + "'";
                    }
                    #endregion

                    #region PRINTING SETTINGS

                    PrintDocument PrintDoc = new PrintDocument();
                    System.Drawing.Printing.PaperSize PkSize = new System.Drawing.Printing.PaperSize(PaperSize, 12, 12);
                    for (int i = 0; i <= PrintDoc.PrinterSettings.PaperSizes.Count - 1; i++)
                    {
                        if (PrintDoc.PrinterSettings.PaperSizes[i].PaperName.ToString().ToUpper() == PaperSize.ToUpper())
                        {
                            PkSize = PrintDoc.PrinterSettings.PaperSizes[i];
                            break;
                        }
                    }
                    int SizeID = GetRawKind(PkSize);
                    oReportDoc.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default;
                    oReportDoc.PrintOptions.PaperSize = (CrystalDecisions.Shared.PaperSize)SizeID;
                    oReportDoc.PrintOptions.PaperOrientation = Orientation;

                    #endregion

                    oReportDoc.SetDataSource(objDataSet.Tables[0]);
                    oReportDoc.Refresh();
                    return oReportDoc;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static ReportDocument GetSubReport(DataSet objDataSet, string SubReportName, CrystalDecisions.CrystalReports.Engine.ReportDocument oMainReport)
        {
            try
            {
                Sections crSections;
                ReportObjects crReportObjects;
                SubreportObject crSubreportObject;
                ReportDocument crSubReportDoc = null;

                crSections = oMainReport.ReportDefinition.Sections;

                foreach (Section crSection in crSections)
                {
                    crReportObjects = crSection.ReportObjects;
                    foreach (ReportObject crReportObject in crReportObjects)
                    {
                        if (crReportObject.Kind.ToString() == "SubreportObject")
                        {
                            crSubreportObject = (SubreportObject)crReportObject;
                            crSubReportDoc = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);
                            if (crSubReportDoc.Name.ToUpper() == SubReportName.ToUpper())
                            {
                                ClsCRDisplay.ReportConnectionInfo(crSubReportDoc);
                                crSubReportDoc.SetDataSource(objDataSet.Tables[0]);
                            }
                        }
                    }
                }
                return crSubReportDoc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region REPORT HANDLING THROUGH XML

        public static ReportDocument GetReport(string ReportPath, string SourceXMLPath, string PaperSize, CrystalDecisions.Shared.PaperOrientation Orientation, bool DirectPrint)
        {
            CrystalDecisions.CrystalReports.Engine.SubreportObject oSubReport;
            CrystalDecisions.CrystalReports.Engine.ReportDocument rptSubReportDoc;
            CrystalDecisions.CrystalReports.Engine.ReportDocument rptDoc = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            rptDoc.Load(ReportPath);
            foreach (CrystalDecisions.CrystalReports.Engine.Table oMainReportTable in rptDoc.Database.Tables)
                oMainReportTable.Location = SourceXMLPath;

            foreach (CrystalDecisions.CrystalReports.Engine.Section rptSection in rptDoc.ReportDefinition.Sections)
            {
                foreach (CrystalDecisions.CrystalReports.Engine.ReportObject rptObject in rptSection.ReportObjects)
                {
                    if (rptObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                    {
                        oSubReport = (CrystalDecisions.CrystalReports.Engine.SubreportObject)rptObject;
                        rptSubReportDoc = oSubReport.OpenSubreport(oSubReport.SubreportName);

                        foreach (CrystalDecisions.CrystalReports.Engine.Table oSubTable in rptSubReportDoc.Database.Tables)
                            oSubTable.Location = SourceXMLPath;
                    }
                }
            }

            #region PRINTING SETTINGS

            PrintDocument PrintDoc = new PrintDocument();
            System.Drawing.Printing.PaperSize PkSize = new System.Drawing.Printing.PaperSize(PaperSize, 12, 12);
            for (int i = 0; i <= PrintDoc.PrinterSettings.PaperSizes.Count - 1; i++)
            {
                if (PrintDoc.PrinterSettings.PaperSizes[i].PaperName.ToString().ToUpper() == PaperSize.ToUpper())
                {
                    PkSize = PrintDoc.PrinterSettings.PaperSizes[i];
                    break;
                }
            }
            int SizeID = GetRawKind(PkSize);
            rptDoc.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default;
            rptDoc.PrintOptions.PaperSize = (CrystalDecisions.Shared.PaperSize)SizeID;
            rptDoc.PrintOptions.PaperOrientation = Orientation;
            //Add the defalut printer Setting
            //if (DirectPrint)
            //    rptDoc.PrintOptions.PrinterName = "";
            //rptDoc.PrintToPrinter(1, false, 0, 0);
            #endregion

            rptDoc.Refresh();
            return rptDoc;
        }

        //public static ReportDocument GetReport(string ReportPath, int month,int year, string PaperSize, CrystalDecisions.Shared.PaperOrientation Orientation)
        //{
        //    string SourceXMLPath = string.Empty ;
        //    ConnectionInfo objConnectionInfo = new ConnectionInfo();
        //    ParameterFields myParams = new ParameterFields(); 


        //    CrystalDecisions.CrystalReports.Engine.SubreportObject oSubReport;
        //    CrystalDecisions.CrystalReports.Engine.ReportDocument rptSubReportDoc;
        //    CrystalDecisions.CrystalReports.Engine.ReportDocument rptDoc = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        //    objConnectionInfo.DatabaseName = Utilities.Application.Company.CompanyDB;
        //    objConnectionInfo.UserID =  Constants.gUSER_ID  ;
        //    objConnectionInfo.Password = Constants.gUSER_PASSWORD;
        //    objConnectionInfo.ServerName = Constants.gSERVER;
        //    rptDoc.Load(ReportPath);
        //    rptDoc.SetDatabaseLogon(Constants.gUSER_ID  , Constants.gUSER_PASSWORD, Constants.gSERVER, Utilities.Application.Company.CompanyDB);

        //    Tables myTables = rptDoc.Database.Tables;
        //    foreach (Table myTable in myTables)
        //    {
        //        TableLogOnInfo myTablesLogonInfo = myTable.LogOnInfo;
        //        myTablesLogonInfo.ConnectionInfo = objConnectionInfo;
        //        myTable.ApplyLogOnInfo(myTablesLogonInfo);
        //    } 
        //     ParameterDiscreteValue     myDiscreteValue =  new ParameterDiscreteValue ();
        //        myDiscreteValue.Value=month ; 


        //    rptDoc.ParameterFields[0].CurrentValues.Clear();// =  myDiscreteValue
        //    rptDoc.ParameterFields[0].CurrentValues.Add(myDiscreteValue);
        //    ParameterDiscreteValue myDiscreteValue1 = new ParameterDiscreteValue();

        //    rptDoc.ParameterFields[1].CurrentValues.Clear();// =  myDiscreteValue
        //    myDiscreteValue1.Value = year;
        //    rptDoc.ParameterFields[1].CurrentValues.Add(myDiscreteValue1);

        //    //foreach (CrystalDecisions.CrystalReports.Engine.Table oMainReportTable in rptDoc.Database.Tables)
        //    //    oMainReportTable.Location = SourceXMLPath;

        //    //foreach (CrystalDecisions.CrystalReports.Engine.Section rptSection in rptDoc.ReportDefinition.Sections)
        //    //{
        //    //    foreach (CrystalDecisions.CrystalReports.Engine.ReportObject rptObject in rptSection.ReportObjects)
        //    //    {
        //    //        if (rptObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
        //    //        {
        //    //            oSubReport = (CrystalDecisions.CrystalReports.Engine.SubreportObject)rptObject;
        //    //            rptSubReportDoc = oSubReport.OpenSubreport(oSubReport.SubreportName);

        //    //            foreach (CrystalDecisions.CrystalReports.Engine.Table oSubTable in rptSubReportDoc.Database.Tables)
        //    //                oSubTable.Location = SourceXMLPath;
        //    //        }
        //    //    }
        //    //}

        //    #region PRINTING SETTINGS

        //    PrintDocument PrintDoc = new PrintDocument();
        //    System.Drawing.Printing.PaperSize PkSize = new System.Drawing.Printing.PaperSize(PaperSize, 12, 12);
        //    for (int i = 0; i <= PrintDoc.PrinterSettings.PaperSizes.Count - 1; i++)
        //    {
        //        if (PrintDoc.PrinterSettings.PaperSizes[i].PaperName.ToString().ToUpper() == PaperSize.ToUpper())
        //        {
        //            PkSize = PrintDoc.PrinterSettings.PaperSizes[i];
        //            break;
        //        }
        //    }
        //    int SizeID = GetRawKind(PkSize);
        //    rptDoc.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default;
        //    rptDoc.PrintOptions.PaperSize = (CrystalDecisions.Shared.PaperSize)SizeID;
        //    rptDoc.PrintOptions.PaperOrientation = Orientation;
        //    //Add the defalut printer Setting
        //    //if (DirectPrint)
        //    //    rptDoc.PrintOptions.PrinterName = "";
        //    //rptDoc.PrintToPrinter(1, false, 0, 0);
        //    #endregion
        //    rptDoc.VerifyDatabase();
        //   // rptDoc.Refresh();
        //    return rptDoc;
        //}


        #endregion

        #region GET COMPANY INFORMATION

        public static void GetReportHeader(SqlConnection objCon, out string CompanyName, out string CompanyAddr)
        {
            try
            {
                string Sql = "Select CompnyName, CompnyAddr From OADM,ADM1";

                SqlCommand cmd = new SqlCommand(Sql, objCon);
                SqlDataAdapter ad = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                ad.Fill(dt);
                CompanyName = dt.Rows[0][0].ToString();
                CompanyAddr = dt.Rows[0][1].ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        public static int GetRawKind(System.Drawing.Printing.PaperSize psize)
        {
            Type t = psize.GetType();
            System.Reflection.FieldInfo fi = t.GetField("kind", System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Instance);
            return (int)fi.GetValue(psize);
        }

        public void SetCrystalLogin(string sUser, string sPassword, string sServer, string sCompanyDB, CrystalDecisions.CrystalReports.Engine.ReportDocument oRpt)
        {
            CrystalDecisions.Shared.ConnectionInfo oConnectInfo = new CrystalDecisions.Shared.ConnectionInfo();

            oConnectInfo.DatabaseName = sCompanyDB;
            oConnectInfo.ServerName = sServer;
            oConnectInfo.UserID = sUser;
            oConnectInfo.Password = sPassword;

            // Set the logon credentials for all tables
            SetCrystalTablesLogin(oConnectInfo, oRpt.Database.Tables);

            // Check for subreports
            foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in oRpt.ReportDefinition.Sections)
            {
                foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                {
                    if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                    {
                        // This is a subreport so set the logon credentials for this report's tables
                        CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;

                        // Open the subreport
                        CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                        SetCrystalTablesLogin(oConnectInfo, oSubRpt.Database.Tables);
                    }
                }
            }

            oRpt.Refresh();

            oRpt.SetDatabaseLogon(sUser, sPassword, sServer, sCompanyDB, false);
            oRpt.VerifyDatabase();

            oRpt.Refresh();

        }

        public static void SetCrystalTablesLogin(CrystalDecisions.Shared.ConnectionInfo oConnectInfo, Tables oTables)
        {
            foreach (CrystalDecisions.CrystalReports.Engine.Table oTable in oTables)
            {
                CrystalDecisions.Shared.TableLogOnInfo oLogonInfo = oTable.LogOnInfo;
                oLogonInfo.ConnectionInfo = oConnectInfo;

                oTable.ApplyLogOnInfo(oLogonInfo);
            }
        }
    }
}
