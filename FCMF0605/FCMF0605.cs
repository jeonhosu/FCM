using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0605
{
    public partial class FCMF0605 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0605()
        {
            InitializeComponent();
        }

        public FCMF0605(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void Set_Default_Value()
        {
            // Budget Select Type.
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "BUDGET_SELECT_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();

            SELECT_TYPE_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
            SELECT_TYPE_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
        }

        private void SearchDB()
        {
            if (iString.ISNull(PERIOD_NAME_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PERIOD_NAME_0.Focus();
                return;
            }

            idaBUDGET_MOVE.Fill();
            Set_Total_Amount();
            igrBUDGET_MOVE.Focus();
            icbALL_RECORD_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
        }

        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter_W(object pGroupCode, object pWhere, object pEnabled_YN)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Total_Amount()
        {
            decimal vTotal_Amount = 0;
            object vAmount;
            int vIDXCol = igrBUDGET_MOVE.GetColumnToIndex("AMOUNT");
            for (int r = 0; r < idaBUDGET_MOVE.SelectRows.Count; r++)
            {
                vAmount = 0;
                vAmount = igrBUDGET_MOVE.GetCellValue(r, vIDXCol);
                vTotal_Amount = vTotal_Amount + iString.ISDecimaltoZero(vAmount);
            }
            TOTAL_AMOUNT.EditValue = vTotal_Amount;
        }

        private void Set_CheckBox()
        {
            int mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("CHECK_YN");
            object mCheck_YN = icbCHECK_YN.CheckBoxValue;
            for (int r = 0; r < igrBUDGET_MOVE.RowCount; r++)
            {
                igrBUDGET_MOVE.SetCellValue(r, mIDX_Col, mCheck_YN);
            }
        }

        private void Set_Grid_Item_Status(DataRow pDataRow)
        {
            bool mEnabled_YN = true;
            int mIDX_Col;
            
            // 신청
            icbCHECK_YN.Enabled = false;
            mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("CHECK_YN");
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 신청금액.
            mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("AMOUNT");
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 신청사유.
            mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("CAUSE_NAME");
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 비고.
            mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("DESCRIPTION");
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            if (pDataRow != null)
            {
                if (iString.ISNull(icbALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString() ||
                    (iString.ISNull(pDataRow["APPROVE_STATUS"]) != "A".ToString() &&
                    iString.ISNull(pDataRow["APPROVE_STATUS"]) != "N".ToString()))
                {
                    mEnabled_YN = false;
                }
                if (iString.ISNull(icbALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString())
                {
                    // 신청
                    icbCHECK_YN.Enabled = true;
                    mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("CHECK_YN");
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
                if (mEnabled_YN == true)
                {                    
                    // 신청금액.
                    mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("AMOUNT");
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 신청사유.
                    mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("CAUSE_NAME");
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 비고.
                    mIDX_Col = igrBUDGET_MOVE.GetColumnToIndex("DESCRIPTION");
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_MOVE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
            }
            igrBUDGET_MOVE.ResetDraw = true;
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaBUDGET_MOVE.IsFocused)
                    {
                        idaBUDGET_MOVE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBUDGET_MOVE.IsFocused)
                    {
                        idaBUDGET_MOVE.Cancel();
                        icbALL_RECORD_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBUDGET_MOVE.IsFocused)
                    {
                        idaBUDGET_MOVE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0605_Load(object sender, EventArgs e)
        {
            idaBUDGET_MOVE.FillSchema();
        }

        private void FCMF0605_Shown(object sender, EventArgs e)
        {
            Set_Default_Value();
            PERIOD_NAME_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            EMAIL_STATUS.EditValue = "N";
            icbALL_RECORD_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
        }

        private void irbAPPROVE_STATUS_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            APPROVE_STATUS_0.EditValue = iStatus.RadioCheckedString;
            SearchDB();
        }

        private void icbCHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Set_CheckBox();
        }

        private void ibtOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }

            idaBUDGET_MOVE.SetUpdateParamValue("P_APPROVE_FLAG", "OK");
            idaBUDGET_MOVE.Update();

            SearchDB();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "C".ToString())
            {
                EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }

            idaBUDGET_MOVE.SetUpdateParamValue("P_APPROVE_FLAG", "CANCEL");
            idaBUDGET_MOVE.Update();

            SearchDB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "N");
            ildDEPT.SetLookupParamValue("W_CHECK_CAPACITY", "Y");
        }

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaSELECT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUDGET_SELECT_TYPE", DBNull.Value, "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBUDGET_MOVE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaBUDGET_MOVE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Item_Status(pBindingManager.DataRow);
        }

        #endregion

    }
}