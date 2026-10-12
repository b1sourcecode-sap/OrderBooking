using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SAPbobsCOM;

namespace OrderBooking_LaxmiSteel
{
    class ClsSalesOrder : Base
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

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsSalesOrder()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsSalesOrder()
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

                        break;
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
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                        }
                        break;
                    #endregion

                    #region et_KEY_DOWN

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ColUID == "1" && pVal.Action_Success == true)
                        {
                          //  _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("38").Specific;
                          //string ItemCode =  ((SAPbouiCOM.EditText)_Matrix.Columns.Item("1").Cells.Item(pVal.Row).Specific).Value.ToString();
                          //  string _StrQry = "Select * From \"@BOOKINGCRM\" Where \"U_SODocNum\" = '" + DocNum + "' AND \"U_DocEntry\" = '" + DocEntry + "'";
                          //  Utilities.ExecuteSQL(ref _RS, _StrQry);
                        }

                        break;

                    #endregion
                }
            #endregion
        }

        #endregion

    }
}
