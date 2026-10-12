using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;

namespace OrderBooking_LaxmiSteel
{
    class ClsOrderBooking : Base
    {

        #region Variable declaration

        SAPbobsCOM.Recordset _RS = null, _RSDelData = null, _RSDaysCount = null, _RS3 = null, _RSCheck = null;
        private SAPbouiCOM.Matrix _Matrix, _Matrix1, _Matrix2, _Matrix3;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbouiCOM.ComboBox _ComboBox;
        public string _CardCode = string.Empty;
        public string _DocEntry = string.Empty;
        public static string _SDocNum = string.Empty;
        public static string _SDocEntry = string.Empty;
        private int RowToDelete;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        public bool _Status = true;
        public static bool _CmbStatus = false;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsOrderBooking()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsOrderBooking()
        {
        }

        #endregion

        #region Item Events

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {
                    #region et_ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            string QueryDeleteData = "delete from \"@OBRTL\" where \"DocEntry\" = (select \"DocEntry\" from \"@OBRTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("9").Specific).Value.ToString().Trim() + "') ";
                            Utilities.ExecuteSQL(ref _RSDelData, QueryDeleteData);
                        }
                        break;
                    #endregion

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "4")
                        {
                            this.FilterBPtYPE("");
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "1000003" && _CmbStatus == true)
                        {
                            string strSQL = "select \"U_OrdBooking\" from OUSR where \"USER_CODE\" = '" + Utilities.Application.Company.UserName + "'";
                            Utilities.ExecuteSQL(ref _RS, strSQL);
                            string test = _RS.Fields.Item(0).Value.ToString().Trim();
                            if (_RS.Fields.Item(0).Value.ToString().Trim() != "Yes")
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("You are not Authorized user for approval.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        break;

                }
            }
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            _Form.Freeze(true);
                            _Form.Items.Item("9").Enabled = false;
                            LoadDataInMatrix();
                            _Form.Freeze(false);
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.ActionSuccess == true)
                        {
                            AddMoveValues();
                        }
                        break;
                    #endregion

                    #region et_COMBO_SELECT

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "32" && pVal.Action_Success == true)
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                            {
                                _Form.Items.Item("8").Enabled = true;
                            }
                            else
                            {
                                _Form.Items.Item("8").Enabled = false;
                                ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value = "0";
                            }
                        }

                        break;

                    #endregion

                    #region Set CFL Values
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "4" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_BPGrpName", 0, _SysCFLEvent.SelectedObjects.GetValue("GroupName", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_BPGrpCode", 0, _SysCFLEvent.SelectedObjects.GetValue("GroupCode", 0).ToString().Trim());

                        }

                        break;
                    #endregion

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "13" && pVal.CharPressed == 9 && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            string BookedQty = ((SAPbouiCOM.EditText)_Form.Items.Item("13").Specific).Value.ToString();
                            ((SAPbouiCOM.EditText)_Form.Items.Item("15").Specific).Value = BookedQty.ToString();
                        }
                        break;
                }
            }
            #endregion
        }

        #endregion

        #region Menu Event

        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            _Form.Freeze(true);
            try
            {
                #region Before Action

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        //case Constants.System_Menus.mnu_Cancelled:
                        //    {
                        //        string StrSQL = "UPDATE \"@DELCH\" set \"U_Status\" = 'Cancelled' where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "'";
                        //        Utilities.ExecuteSQL(ref _RS, StrSQL);
                        //    }
                        //    break;
                    }
                }

                #endregion

                #region After Action

                else
                {
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIND:
                            {
                                _Form.Items.Item("9").Enabled = true;

                                // LoadDataInMatrix();
                            }
                            break;
                        case Constants.System_Menus.mnu_ADD:
                            {
                                //if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                                //    _Form.Items.Item("8").Enabled = true;
                                //else
                                //    _Form.Items.Item("8").Enabled = false;

                                AddMoveValues();
                            }
                            break;
                        case Constants.System_Menus.mnu_FIRST:
                            {
                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                                    _Form.Items.Item("8").Enabled = true;
                                else
                                    _Form.Items.Item("8").Enabled = false;

                                LoadDataInMatrix();
                            }
                            break;
                        case Constants.System_Menus.mnu_LAST:
                            {
                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                                    _Form.Items.Item("8").Enabled = true;
                                else
                                    _Form.Items.Item("8").Enabled = false;

                                LoadDataInMatrix();
                            }
                            break;
                        case Constants.System_Menus.mnu_NEXT:
                            {
                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                                    _Form.Items.Item("8").Enabled = true;
                                else
                                    _Form.Items.Item("8").Enabled = false;

                                LoadDataInMatrix();
                            }
                            break;
                        case Constants.System_Menus.mnu_PREVIOUS:
                            {
                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                                    _Form.Items.Item("8").Enabled = true;
                                else
                                    _Form.Items.Item("8").Enabled = false;

                                LoadDataInMatrix();
                            }
                            break;

                        case Constants.System_Menus.mnu_Refresh:
                            {
                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("32").Specific).Selected.Value == "Yes")
                                    _Form.Items.Item("8").Enabled = true;
                                else
                                    _Form.Items.Item("8").Enabled = false;

                                LoadDataInMatrix();
                            }
                            break;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion

        #region FormDefault

        public void FormDefault()
        {
            _Form.Freeze(true);

            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("1000001").Specific;
            _Matrix1 = (SAPbouiCOM.Matrix)Form.Items.Item("Item_9").Specific;
            _Form.DataBrowser.BrowseBy = "9";
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@OBRTH");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@OBRTL");
            _ChildDBDataSource1 = _Form.DataSources.DBDataSources.Item("@OBRTL1");
            _CmbStatus = false;
            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("3").Specific;
            _ComboBox.ValidValues.Add("Individual", "Individual");
            _ComboBox.ValidValues.Add("Agency", "Agency");
            _ComboBox.Select("Individual", BoSearchKey.psk_ByDescription);

            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("1000003").Specific;
            _ComboBox.ValidValues.Add("Draft", "Draft");
            _ComboBox.ValidValues.Add("Open", "Open");
            _ComboBox.ValidValues.Add("Approved", "Approved");
            _ComboBox.ValidValues.Add("Closed", "Closed");
            _ComboBox.ValidValues.Add("Expired", "Expired");
            _ComboBox.ValidValues.Add("Terminated", "Terminated");
            _ComboBox.Select("Draft", BoSearchKey.psk_ByDescription);

            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("34").Specific;
            Utilities.FillCombo(ref _ComboBox, "select \"Code\",\"Name\" from \"@BOOKTYPE\" ");

            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("32").Specific;
            _ComboBox.ValidValues.Add("Yes", "Yes");
            _ComboBox.ValidValues.Add("No", "No");
            _ComboBox.Select("No", BoSearchKey.psk_ByDescription);

            ((SAPbouiCOM.EditText)_Form.Items.Item("9").Specific).Value = Utilities.getMaxColumnValueNum("@OBRTH", "U_DocNum");

            int _Days = 0;
            Utilities.ExecuteSQL(ref _RSDaysCount, "select \"U_ValidDays\" from OADM");
            _Days = Convert.ToInt16(_RSDaysCount.Fields.Item("U_ValidDays").Value);

            ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value = DateTime.Today.ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = (DateTime.Today.AddDays(_Days)).ToString("yyyyMMdd");

            _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Items.Item("9").Enabled = false;
            _Form.Items.Item("8").Enabled = false;
            _Form.Items.Item("5").Enabled = false;
            _Form.Items.Item("15").Enabled = false;
            _Form.Items.Item("11").Enabled = false;
            _Form.Items.Item("12").Enabled = false;
            _Status = true;
            Utilities.SerializedMartix(ref _Form, ref _Matrix);
            _CmbStatus = true;
            _Form.Freeze(false);
        }

        #endregion

        #region FormDefault

        public void AddMoveValues()
        {
            _Form.Freeze(true);

            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("1000001").Specific;
            _Matrix1 = (SAPbouiCOM.Matrix)Form.Items.Item("Item_9").Specific;
            _Form.DataBrowser.BrowseBy = "9";
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@OBRTH");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@OBRTL");
            _ChildDBDataSource1 = _Form.DataSources.DBDataSources.Item("@OBRTL1");
            _CmbStatus = false;
            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("3").Specific;
            _ComboBox.Select("Individual", BoSearchKey.psk_ByDescription);

            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("1000003").Specific;
            _ComboBox.Select("Draft", BoSearchKey.psk_ByDescription);

            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("32").Specific;
            _ComboBox.Select("No", BoSearchKey.psk_ByDescription);

            ((SAPbouiCOM.EditText)_Form.Items.Item("9").Specific).Value = Utilities.getMaxColumnValueNum("@OBRTH", "U_DocNum");
            int _Days = 0;
            Utilities.ExecuteSQL(ref _RSDaysCount, "select \"U_ValidDays\" from OADM");
            _Days = Convert.ToInt16(_RSDaysCount.Fields.Item("U_ValidDays").Value);

            ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value = DateTime.Today.ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = (DateTime.Today.AddDays(_Days)).ToString("yyyyMMdd");

            _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Items.Item("9").Enabled = false;
            _Form.Items.Item("8").Enabled = false;
            _Form.Items.Item("5").Enabled = false;
            _Form.Items.Item("15").Enabled = false;
            _Form.Items.Item("11").Enabled = false;
            _Form.Items.Item("12").Enabled = false;
            _Status = true;
            Utilities.SerializedMartix(ref _Form, ref _Matrix);
            _CmbStatus = true;
            _Form.Freeze(false);

        }

        #endregion

        #region FILTER BP Type
        internal void FilterBPtYPE(string CFLId)
        {
            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "GroupType";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "C";
                _Con.BracketCloseNum = 1;
                _Form.ChooseFromLists.Item("CFL_1").SetConditions(_Cons);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region LoadDataInMatrix

        public void LoadDataInMatrix1()
        {
            _Matrix.Clear();
            //Utilities.Application.SBO_Application.StatusBar.SetText("Please wait", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            string QueryFillData = "select T0.\"CardCode\",T0.\"CardName\",T1.\"ItemCode\",T1.\"Dscription\", T3.\"ItmsGrpNam\", " +
                                   " T0.\"DocEntry\",T0.\"DocNum\",T0.\"DocDate\",case when T0.\"DocStatus\" = 'O' then 'Open' " +
                                   "when T0.\"DocStatus\" = 'C' AND T0.\"CANCELED\" = 'N' then 'Closed' when T0.\"DocStatus\" = 'C' AND T0.\"CANCELED\" = 'Y' then 'Canceled' END as \"DocStatus\",T1.\"Quantity\" " +
                                   " FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                   " INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" " +
                                   " INNER JOIN OITB T3 ON T2.\"ItmsGrpCod\" = T3.\"ItmsGrpCod\" where  T1.\"U_RtBookNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("999").Specific).Value.ToString().Trim() + "' ";

            Utilities.ExecuteSQL(ref _RS, QueryFillData);

            //Utilities.Application.SBO_Application.StatusBar.SetText("Please wait " + _RS.RecordCount + " lines data loading", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            if (_RS.RecordCount > 0)
            {
                for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix.AddRow(1, -1);
                    _Matrix.FlushToDataSource();
                    _ChildDBDataSource.SetValue("U_BPCode", i, _RS.Fields.Item("CardCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_BPName", i, _RS.Fields.Item("CardName").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemCode", i, _RS.Fields.Item("ItemCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemName", i, _RS.Fields.Item("Dscription").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItmsGrp", i, _RS.Fields.Item("ItmsGrpNam").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocEntry", i, _RS.Fields.Item("DocEntry").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocNum", i, _RS.Fields.Item("DocNum").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocDate", i, DateTime.Parse(_RS.Fields.Item("DocDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_Status", i, _RS.Fields.Item("DocStatus").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Quantity", i, _RS.Fields.Item("Quantity").Value.ToString());
                    _Matrix.LoadFromDataSource();
                    _RS.MoveNext();
                }
            }
            // _Matrix.Columns.Item(0).Editable = false;
            //_Matrix.Columns.Item(1).Editable = false;
            //_Matrix.Columns.Item(2).Editable = false;
            //_Matrix.Columns.Item(3).Editable = false;
            //_Matrix.Columns.Item(4).Editable = false;
            //_Matrix.Columns.Item(5).Editable = false;
            //_Matrix.Columns.Item(6).Editable = false;
            //_Matrix.Columns.Item(7).Editable = false;
            //_Matrix.Columns.Item(8).Editable = false;
            //_Matrix.Columns.Item(9).Editable = false;
            //_Matrix.Columns.Item(10).Editable = false;
            for (int i = 0; i < _Form.Items.Count; i++)
            {
                try
                {
                    _Form.Items.Item(i).Enabled = false;
                }
                catch { }
            }

            // Keep OK and Cancel enabled
            _Form.Items.Item("1").Enabled = true;
            _Form.Items.Item("2").Enabled = true;

            // Keep tabs enabled
            _Form.Items.Item("Item_7").Enabled = true;
            _Form.Items.Item("Item_8").Enabled = true;
            Utilities.SerializedMartix(ref _Form, ref _Matrix);
            _Form.PaneLevel = 0;
            SAPbouiCOM.ComboBox partyType = (SAPbouiCOM.ComboBox)_Form.Items.Item("3").Specific;
            string partyTypes = partyType.Selected != null
                ? partyType.Selected.Value.Trim()
                : "";
            if (partyTypes == "Agency")
            {
                _Matrix1.Clear();
                string QueryFillData1 = "select T0.\"CardCode\",T0.\"CardName\" from \"OCRD\" T0 where  T0.\"GroupCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value.ToString().Trim() + "' ";

                Utilities.ExecuteSQL(ref _RS3, QueryFillData1);
                if (_RS3.RecordCount > 0)
                {
                    for (int i = 0; i < _RS3.RecordCount; i++)
                    {
                        _Matrix1.AddRow(1, -1);
                        _Matrix1.FlushToDataSource();
                        _ChildDBDataSource1.SetValue("U_BPCode", i, _RS3.Fields.Item("CardCode").Value.ToString());
                        _ChildDBDataSource1.SetValue("U_BPName", i, _RS3.Fields.Item("CardName").Value.ToString());
                        _Matrix1.LoadFromDataSource();
                        _RS3.MoveNext();
                    }
                    _Matrix1.Columns.Item(0).Editable = false;
                    _Matrix1.Columns.Item(1).Editable = false;
                    _Matrix1.Columns.Item(2).Editable = false;
                    Utilities.SerializedMartix(ref _Form, ref _Matrix1);
                    _Form.PaneLevel = 1;
                }
                
            }
        }

        public void LoadDataInMatrix()
        {
            _Matrix.Clear();

            string QueryFillData =
                "SELECT T0.\"CardCode\", T0.\"CardName\", T1.\"ItemCode\", T1.\"Dscription\", " +
                "T3.\"ItmsGrpNam\", T0.\"DocEntry\", T0.\"DocNum\", T0.\"DocDate\", " +
                "CASE " +
                "WHEN T0.\"DocStatus\" = 'O' THEN 'Open' " +
                "WHEN T0.\"DocStatus\" = 'C' AND T0.\"CANCELED\" = 'N' THEN 'Closed' " +
                "WHEN T0.\"DocStatus\" = 'C' AND T0.\"CANCELED\" = 'Y' THEN 'Canceled' " +
                "END AS \"DocStatus\", T1.\"Quantity\" " +
                "FROM ORDR T0 " +
                "INNER JOIN RDR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                "INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" " +
                "INNER JOIN OITB T3 ON T2.\"ItmsGrpCod\" = T3.\"ItmsGrpCod\" " +
                "WHERE T1.\"U_RtBookNo\" = '" +
                ((SAPbouiCOM.EditText)_Form.Items.Item("9").Specific).Value.Trim() + "'";

            Utilities.ExecuteSQL(ref _RS, QueryFillData);

            if (_RS.RecordCount > 0)
            {
                for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix.AddRow(1, -1);
                    _Matrix.FlushToDataSource();

                    _ChildDBDataSource.SetValue("U_BPCode", i, _RS.Fields.Item("CardCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_BPName", i, _RS.Fields.Item("CardName").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemCode", i, _RS.Fields.Item("ItemCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemName", i, _RS.Fields.Item("Dscription").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItmsGrp", i, _RS.Fields.Item("ItmsGrpNam").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocEntry", i, _RS.Fields.Item("DocEntry").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocNum", i, _RS.Fields.Item("DocNum").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocDate", i,
                        DateTime.Parse(_RS.Fields.Item("DocDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_Status", i, _RS.Fields.Item("DocStatus").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Quantity", i, _RS.Fields.Item("Quantity").Value.ToString());

                    _Matrix.LoadFromDataSource();
                    _RS.MoveNext();
                }
            }

            Utilities.SerializedMartix(ref _Form, ref _Matrix);

            // Read Party Type BEFORE disabling fields
            SAPbouiCOM.ComboBox partyType =
                (SAPbouiCOM.ComboBox)_Form.Items.Item("3").Specific;

            string partyTypes = partyType.Selected != null
                ? partyType.Selected.Value.Trim()
                : "";

            // ---------------------------------------------------------
            // Make normal form fields read-only
            // ---------------------------------------------------------
            for (int i = 0; i < _Form.Items.Count; i++)
            {
                try
                {
                    _Form.Items.Item(i).Enabled = false;
                }
                catch
                {
                }
            }

            // Keep buttons enabled
            _Form.Items.Item("1").Enabled = true;
            _Form.Items.Item("2").Enabled = true;

            // Keep tabs enabled
            _Form.Items.Item("Item_7").Enabled = true;
            _Form.Items.Item("Item_8").Enabled = true;

            // IMPORTANT:
            // Keep Matrix enabled so LinkedObject/BP arrow works
            _Form.Items.Item("1000001").Enabled = true;

            // Make first matrix read-only
            for (int i = 0; i < _Matrix.Columns.Count; i++)
            {
                _Matrix.Columns.Item(i).Editable = false;
            }

            // ---------------------------------------------------------
            // Agency
            // ---------------------------------------------------------
            if (partyTypes == "Agency")
            {
                _Matrix1.Clear();

                string QueryFillData1 =
                    "SELECT T0.\"CardCode\", T0.\"CardName\" " +
                    "FROM \"OCRD\" T0 " +
                    "WHERE T0.\"GroupCode\" = '" +
                    ((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value.Trim() + "'";

                Utilities.ExecuteSQL(ref _RS3, QueryFillData1);

                if (_RS3.RecordCount > 0)
                {
                    for (int i = 0; i < _RS3.RecordCount; i++)
                    {
                        _Matrix1.AddRow(1, -1);
                        _Matrix1.FlushToDataSource();

                        _ChildDBDataSource1.SetValue(
                            "U_BPCode",
                            i,
                            _RS3.Fields.Item("CardCode").Value.ToString());

                        _ChildDBDataSource1.SetValue(
                            "U_BPName",
                            i,
                            _RS3.Fields.Item("CardName").Value.ToString());

                        _Matrix1.LoadFromDataSource();
                        _RS3.MoveNext();
                    }
                }

                Utilities.SerializedMartix(ref _Form, ref _Matrix1);

                // IMPORTANT: Matrix must remain enabled for LinkedObject
                _Form.Items.Item("Item_9").Enabled = true;

                // Make columns read-only
                for (int i = 0; i < _Matrix1.Columns.Count; i++)
                {
                    _Matrix1.Columns.Item(i).Editable = false;
                }

                // Keep Customer Details tab enabled
                _Form.Items.Item("Item_8").Enabled = true;

                // Show Customer Details pane
                _Form.PaneLevel = 1;
            }
            else
            {
                // Show Details pane
                _Form.PaneLevel = 0;
            }
        }
        #endregion
    }

}

