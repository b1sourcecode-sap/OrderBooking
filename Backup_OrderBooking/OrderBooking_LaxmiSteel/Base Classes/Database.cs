using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;

namespace OrderBooking_LaxmiSteel
{
    /// <summary>
    /// Summary description for Database.
    /// </summary>
    public sealed class Database
    {
        private Database()
        {
        }

        #region INITIALISE DATABASE
        public static void InitializeDatabase()
        {
            SAPbouiCOM.ProgressBar oProgressBar = null;
            string[] oChildTables = new string[3];
            string[] oFindColumns = new string[3];
            string[] oColumnName = new string[2];
            try
            {
                oProgressBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Initializing database...Please wait.", 50, false);
                oProgressBar.Value = 1;

                if (!Utilities.Application.Company.InTransaction)
                    Utilities.Application.Company.StartTransaction();


                #region Order Booking Rate Header

                CreateTable("OBRTH", "Order Booking Rate Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("OBRTH", "PartyType", "Party Type", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "BPGrpCode", "BP Group Code", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "BPGrpName", "BP Group Name", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "UsedCRLmt", "Used Credit Limit", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "AvlCRLmt", "Available Credit Limit", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "Remarks", "Remarks", 254, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "FrtSbSdy", "Fereight Subsidy", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "SbSdyAmt", "Fereight Subsidy Amount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "DocNum", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "ValidUpt", "Valid UpTo", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "Qty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "Rate", "Rate", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "RemQty", "Remaining Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "Status", "Doc Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTH", "BookType", "Booking Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
               
                #endregion

                #region Order Booking Rate Line

                CreateTable("OBRTL", "Order Booking Rate Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("OBRTL", "BPCode", "BP Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "BPName", "BP Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "ItemCode", "Item Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "ItmsGrp", "Item Group", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "DocNum", "Sales Order No", 11, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "DocEntry", "Sales Order Entry", 11, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "DocDate", "SO Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "Status", "SO Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("OBRTL", "Quantity", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
               
                #endregion

                AddColumn("RDR1", "RtBookNo", "Rate Booking No", 11, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OADM", "ValidDays", "Rate Booking Valid Days", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);

                CreateTable("BOOKTYPE", "Booking Type", SAPbobsCOM.BoUTBTableType.bott_NoObject);

                #region UDF - OUSR
                AddColumn("OUSR", "OrdBooking", "Order Booking Approver", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region CREATE UDO

                oChildTables[0] = "OBRTL";
                oFindColumns[0] = "U_DocNum";
                CreateUDO("OBRTH", "Order Booking Rate Header", "OBRTH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);

                #endregion


                //Utilities.CreateFunction_RemoveSpecialChar();

                if (Utilities.Application.Company.InTransaction)
                    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                // Initialize Connection String
                //Constants.gobjSQLCon = new SqlConnection("user id=" + Constants.gUSER_ID + ";data source=" + Constants.gSERVER + ";pwd=" + Constants.gUSER_PASSWORD + ";initial catalog=" + Utilities.Application.Company.CompanyDB);
                //Constants.gConnStr = "user id=" + Constants.gUSER_ID + ";data source=" + Constants.gSERVER + ";pwd=" + Constants.gUSER_PASSWORD + ";initial catalog=" + Utilities.Application.Company.CompanyDB;
                oProgressBar.Stop();
            }
            catch (Exception ex)
            {
                oProgressBar.Stop();

                if (Utilities.Application.Company.InTransaction)
                    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                oProgressBar = null;
                GC.Collect();
            }
        }


        #endregion

        #region DATABASE CREATION FUNCTIONS

        #region CREATE USER DEFINED TABLE

        private static void CreateTable(string oTable, string oDescription, SAPbobsCOM.BoUTBTableType oType)
        {
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTable = null;

            try
            {
                oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!oUserTable.GetByKey(oTable))
                {
                    oUserTable.TableName = oTable;
                    oUserTable.TableDescription = oDescription;
                    oUserTable.TableType = oType;

                    if (oUserTable.Add() != 0)

                        throw new Exception(Utilities.Application.Company.GetLastErrorDescription());

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
            }
        }

        private static void UpdateTable(string oTable, SAPbobsCOM.BoUTBTableType oType)
        {
            SAPbobsCOM.UserTablesMD oUserTable = null;

            try
            {
                oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUserTable.GetByKey(oTable) != null)
                {
                    if (oUserTable.TableType != SAPbobsCOM.BoUTBTableType.bott_MasterData)
                    {
                        //oUserTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData;
                        if (oUserTable.Remove() != 0)

                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
            }
        }
        #endregion

        #region ADD COLUMN
        internal static void AddColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.UserFieldsMD oUserField = null;
            int Fieldid;
            try
            {

                if (ColumnExists(oTable, oName, oIsSystemTable, out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tNO)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                    oUserField.TableName = oTable;
                    oUserField.Name = oName;
                    oUserField.Description = oDescription;
                    oUserField.Type = oType;
                    oUserField.SubType = oSubType;

                    if (oType != SAPbobsCOM.BoFieldTypes.db_Float)
                        oUserField.EditSize = oSize;
                    else
                        oUserField.EditSize = 0;


                    oUserField.Mandatory = oMandatory;

                    if ((oTable == "OITM" && oName == "B1FIPACK") || (oTable == "B1F_FSAM" && oName == "B1FPACK") || (oTable == "B1F_FSAC" && oName == "B1FPFS"))
                    {
                        oUserField.ValidValues.Value = "N";
                        oUserField.ValidValues.Description = "N";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Y";
                        oUserField.ValidValues.Description = "Y";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "N";

                    }
                    else if (oTable == "OINV" && oName == "SaleType")
                    {
                        oUserField.ValidValues.Value = "Domestic";
                        oUserField.ValidValues.Description = "Domestic";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Export";
                        oUserField.ValidValues.Description = "Export";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "Domestic";
                    }
                    else if (oTable == "OUSR" && oName == "OrdBooking")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }

                    


                    if (oUserField.Add() != 0)
                        throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oUserField != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();
            }
        }

        internal static void UpdateColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.UserFieldsMD oUserField = null;
            string oTable1 = oTable;
            int Fieldid = 0;
            try
            {

                if (ColumnExists(oTable, oName, oIsSystemTable, out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    if (oTable != "OCRD")
                        oTable1 = @"@" + oTable;

                    oUserField.GetByKey(oTable1, Fieldid);
                    if (oTable == "B1F_GNDR" && (oName == "B1FINACT"))
                    {
                        if (oUserField.DefaultValue != "N")
                        {
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }
                    }
                    else if ((oTable == "OCRD" && oName == "Season") || (oTable == "B1F_ITEMSEAS" && oName == "B1FSEASD") || (oTable == "B1F_ITEMCAPS" && (oName == "B1FCAPSD" || oName == "B1FDTST" || oName == "B1FDTCP")))
                    {
                        if (oUserField.Remove() != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());

                    }
                    else if (oTable == "OCRD" && (oName == "B1FBUYINP" || oName == "B1FBUYINPN"))
                    {
                        if (oUserField.ValidValues.Value == "")
                        {
                            oUserField.ValidValues.Value = "N";
                            oUserField.ValidValues.Description = "No";
                            oUserField.ValidValues.Add();
                            oUserField.ValidValues.Value = "Y";
                            oUserField.ValidValues.Description = "Yes";
                            oUserField.ValidValues.Add();
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                    else
                    {
                        if ((oTable == "B1F_FSPM" && oName == "B1FSKUNUM") || (oTable == "B1F_FSPM" && oName == "B1FATLEN") || (oTable == "B1F_FSAC" && oName == "B1FSEQ") || (oTable == "B1F_FSIM" && oName == "B1FSKULN") || (oTable == "B1F_VAT1" && oName == "B1FGRPSQ") || (oTable == "B1F_VAT2" && oName == "B1FGRPSQ"))
                        {
                            if (oUserField.EditSize < oSize)
                            {
                                oUserField.EditSize = oSize;
                                if (oUserField.Update() != 0)
                                    throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                            }
                        }
                        if (oUserField.Description != oDescription)
                        {
                            oUserField.Description = oDescription;
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                }


            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oUserField != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();
            }
        }

        #endregion

        //#region IS COLUMN EXISTS
        //private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable,out int FieldId)
        //{
        //    SAPbobsCOM.Recordset oRSColumn = null;
        //    string               oSQL      = string.Empty;
        //    FieldId = 0;
        //    try
        //    {
        //        if( oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO )
        //            oTable = @"@" + oTable;

        //        oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
        //        Utilities.ExecuteSQL(ref oRSColumn, oSQL);

        //        if ((int)oRSColumn.Fields.Item(0).Value == 0)
        //            return SAPbobsCOM.BoYesNoEnum.tNO;

        //        else
        //        {
        //            FieldId = (int)oRSColumn.Fields.Item(1).Value;
        //            return SAPbobsCOM.BoYesNoEnum.tYES;
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        throw ex;
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
        //        oRSColumn = null;
        //        GC.Collect();
        //    }
        //}
        //#endregion

        #region IS COLUMN EXISTS
        private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable, out int FieldId)
        {
            SAPbobsCOM.Recordset oRSColumn = null;
            string oSQL = string.Empty;
            FieldId = 0;
            try
            {
                if (oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO)
                    oTable = @"@" + oTable;

                //oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
                oSQL = @"Select " +
                        @"Count(*) "
                        + " , "
                        + @"""FieldID"""
                        + "From CUFD Where "
                        + @"""TableID"""
                         + "="
                        + "'" + oTable.Trim() + "'"
                        + " And "
                        + @"""AliasID"""
                        + "="
                        + "'" + oColumn.Trim() + "'"
                        + " group by "
                        + @"""FieldID""".Trim();


                Utilities.ExecuteSQL(ref oRSColumn, oSQL);

                if ((int)oRSColumn.Fields.Item(0).Value == 0)
                    return SAPbobsCOM.BoYesNoEnum.tNO;

                else
                {
                    FieldId = (int)oRSColumn.Fields.Item(1).Value;
                    return SAPbobsCOM.BoYesNoEnum.tYES;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
                oRSColumn = null;
                GC.Collect();
            }
        }
        #endregion


        #region CREATE UDO
        private static void CreateUDO(string oUniqueID, string oDescription, string oTable, SAPbobsCOM.BoUDOObjType oType, string[] oFindColumns, string[] oChildTables)
        {
            SAPbobsCOM.UserObjectsMD oUserObject = null;
            int oCount;

            try
            {
                oUserObject = (SAPbobsCOM.UserObjectsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                if (!oUserObject.GetByKey(oUniqueID))
                {
                    oUserObject.Code = oUniqueID;
                    oUserObject.Name = oDescription;
                    oUserObject.ObjectType = oType;
                    oUserObject.TableName = oTable;

                    oUserObject.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObject.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObject.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObject.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObject.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (oFindColumns != null)
                    {
                        if (oFindColumns.Length > 0)
                        {
                            oUserObject.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

                            for (oCount = 0; oCount <= oFindColumns.Length - 1; oCount++)
                            {
                                if (oFindColumns[oCount] != null)
                                {
                                    oUserObject.FindColumns.ColumnAlias = oFindColumns[oCount];
                                    oUserObject.FindColumns.Add();
                                }

                                oFindColumns[oCount] = null;
                            }
                        }
                    }

                    if (oChildTables != null)
                    {
                        if (oChildTables.Length > 0)
                        {
                            for (oCount = 0; oCount <= oChildTables.Length - 1; oCount++)
                            {
                                if (oChildTables[oCount] != null)
                                {
                                    oUserObject.ChildTables.TableName = oChildTables[oCount];

                                    if (oCount != oChildTables.Length - 1)
                                        oUserObject.ChildTables.Add();
                                }

                                oChildTables[oCount] = null;
                            }
                        }
                    }

                    if (oUserObject.Add() != 0)
                        throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObject);
                oUserObject = null;
                GC.Collect();
            }
        }
        #endregion

        public static void CreateUserObject(string CodeID, string Name, string TableName, string Child) //used for registration of user defined table
        {

            int lRetCode = 0;
            string sErrMsg = null;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            try
            {
                if (oUserObjectMD == null)
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

                if (oUserObjectMD.GetByKey(CodeID) == true)
                    return; //vijay 170708

                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;


                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.TableName = TableName;

                if (Child != "")
                {
                    oUserObjectMD.ChildTables.TableName = Child;
                    //oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                lRetCode = oUserObjectMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { }
                    else
                    { Utilities.Application.Company.GetLastError(out lRetCode, out sErrMsg); }
                else
                { }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #region IS KEY EXISTS
        private static SAPbobsCOM.BoYesNoEnum KeyExists(string oTable, string oKeyName, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.Recordset oRSColumn = null;
            string oSQL = string.Empty;

            try
            {
                if (oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO)
                    oTable = @"@" + oTable;
                oSQL = "Select Count(*) From OUKD Where TableName = '" + oTable + "' AND KeyName = '" + oKeyName + "' ";
                Utilities.ExecuteSQL(ref oRSColumn, oSQL);

                if ((int)oRSColumn.Fields.Item(0).Value == 0)
                    return SAPbobsCOM.BoYesNoEnum.tNO;

                else
                    return SAPbobsCOM.BoYesNoEnum.tYES;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
                oRSColumn = null;
                GC.Collect();
            }
        }
        #endregion

        #region ADD KEYS
        private static void AddKey(string oTable, string[] oColumnName, string oKeyName, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.UserKeysMD oUserKey = null;
            int oCount;
            try
            {
                oUserKey = (SAPbobsCOM.UserKeysMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
                if (KeyExists(oTable, oKeyName, oIsSystemTable) == SAPbobsCOM.BoYesNoEnum.tNO)
                {
                    oUserKey.TableName = oTable;
                    oUserKey.KeyName = oKeyName;
                    if (oColumnName != null)
                    {
                        if (oColumnName.Length > 0)
                        {
                            for (oCount = 0; oCount <= oColumnName.Length - 1; oCount++)
                            {
                                if (oColumnName[oCount] != null)
                                {
                                    oUserKey.Elements.ColumnAlias = oColumnName[oCount];
                                    oUserKey.Elements.Add();
                                }

                                oColumnName[oCount] = null;
                            }
                        }
                    }
                    oUserKey.Unique = SAPbobsCOM.BoYesNoEnum.tYES;
                    if (oUserKey.Add() != 0)
                        throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKey);
                oUserKey = null;
                GC.Collect();
            }
        }
        #endregion

        #endregion

    }
}
