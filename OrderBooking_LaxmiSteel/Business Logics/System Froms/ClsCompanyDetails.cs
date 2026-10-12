using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SAPbobsCOM;

namespace OrderBooking_LaxmiSteel
{
    class ClsCompanyDetails : Base
    {
        #region Variable declaration

        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null, _RS3 = null, _RSCheck = null;
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
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.Button _Button;
        SAPbouiCOM.StaticText _StaticText;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsCompanyDetails()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsCompanyDetails()
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

                    //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                    //    break;

                    #endregion

                }
            }
            #endregion

            #region After Action
            else

                switch (pVal.EventType)
                {
                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "PrcStatus")
                        {
                            string _StrQry = "update \"@OBRTH\" set \"U_Status\" = 'Expired' where \"U_Status\" not in ('Closed','Terminated') AND \"U_ValidUpt\" < (SELECT CURRENT_DATE FROM DUMMY)";
                            Utilities.ExecuteSQL(ref _RS, _StrQry);

                            string _StrQry1 = "update \"@OBRTH\" set \"U_Status\" = 'Closed' where IFNULL(\"U_RemQty\",0) <= 0 and \"U_Status\" = 'Open'";
                            Utilities.ExecuteSQL(ref _RS1, _StrQry1);
                        }
                        break;

                    #endregion

                }
            #endregion
        }

        #endregion

        #region Form Default
        internal void FormDefault()
        {
            try
            {
                //string strSQL = "select \"U_LCostReq\" from OUSR where \"USER_CODE\" = '" + Utilities.Application.Company.UserName + "'";
                //Utilities.ExecuteSQL(ref _RS, strSQL);
                //string test = _RS.Fields.Item(0).Value.ToString().Trim();
                //if (_RS.Fields.Item(0).Value.ToString().Trim() == "YES")
                //{
                    _ExistItem = _Form.Items.Item("2");
                    _NewItem = _Form.Items.Add("PrcStatus", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                    _NewItem.Width = _ExistItem.Width + 10;
                    _NewItem.Top = _ExistItem.Top;
                    _NewItem.Height = _ExistItem.Height;
                    _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                    _Button.Caption = "Process Status";
                //}

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion
    }
}
