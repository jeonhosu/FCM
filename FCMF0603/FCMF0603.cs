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

namespace FCMF0603
{
    public partial class FCMF0603 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0603()
        {
            InitializeComponent();
        }

        public FCMF0603(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(PERIOD_NAME_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PERIOD_NAME_0.Focus();
                return;
            }
            idaBUDGET_ADD.Fill();
            Set_Total_Amount();
            igrBUDGET_ADD.Focus();
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
            int vIDXCol = igrBUDGET_ADD.GetColumnToIndex("AMOUNT");
            for (int r = 0; r < idaBUDGET_ADD.SelectRows.Count; r++)
            {
                vAmount = 0;
                vAmount = igrBUDGET_ADD.GetCellValue(r, vIDXCol);
                vTotal_Amount = vTotal_Amount + iString.ISDecimaltoZero(vAmount);
            }
            TOTAL_AMOUNT.EditValue = vTotal_Amount;
        }

        private void Set_CheckBox()
        {
            int mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("CHECK_YN");
            object mCheck_YN = icbCHECK_YN.CheckBoxValue;
            for (int r = 0; r < igrBUDGET_ADD.RowCount; r++)
            {
                igrBUDGET_ADD.SetCellValue(r, mIDX_Col, mCheck_YN);
            }
        }

        private void Set_Grid_Item_Status(DataRow pDataRow)
        {
            bool mEnabled_YN = true;
            int mIDX_Col;

            // 선택.
            icbCHECK_YN.Enabled = false;
            mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("CHECK_YN");
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 신청금액.
            mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("AMOUNT");
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 신청사유.
            mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("CAUSE_NAME");
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 비고.
            mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("DESCRIPTION");
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            if (pDataRow != null)
            {
                if (iString.ISNull(icbALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString() ||
                    (iString.ISNull(pDataRow["APPROVE_STATUS"]) != "A".ToString() &&
                    iString.ISNull(pDataRow["APPROVE_STATUS"]) != "N".ToString()))
                {
                    mEnabled_YN = false;
                }
                if (iString.ISNull(icbALL_RECORD_FLAG.CheckBoxValue) != "Y".ToString())
                {
                    // 선택.
                    icbCHECK_YN.Enabled = true;
                    mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("CHECK_YN");
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
                if (mEnabled_YN == true)
                {                    
                    // 신청금액.
                    mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("AMOUNT");
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 신청사유.
                    mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("CAUSE_NAME");
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 비고.
                    mIDX_Col = igrBUDGET_ADD.GetColumnToIndex("DESCRIPTION");
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrBUDGET_ADD.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
            }
            igrBUDGET_ADD.ResetDraw = true;
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
                    if (idaBUDGET_ADD.IsFocused)
                    {
                        idaBUDGET_ADD.SetUpdateParamValue("P_APPROVE_FLAG", DBNull.Value);
                        idaBUDGET_ADD.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBUDGET_ADD.IsFocused)
                    {
                        idaBUDGET_ADD.Cancel();
                        icbALL_RECORD_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBUDGET_ADD.IsFocused)
                    {
                        idaBUDGET_ADD.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0603_Load(object sender, EventArgs e)
        {
            idaBUDGET_ADD.FillSchema();
        }

        private void FCMF0603_Shown(object sender, EventArgs e)
        {
            PERIOD_NAME_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            EMAIL_STATUS.EditValue = "N";
            icbALL_RECORD_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
        }

        private void irbAPPROVE_STATUS_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            APPROVE_STATUS_0.EditValue = iStatus.RadioButtonString;
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

            idaBUDGET_ADD.SetUpdateParamValue("P_APPROVE_FLAG", "OK");
            idaBUDGET_ADD.Update();

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

            idaBUDGET_ADD.SetUpdateParamValue("P_APPROVE_FLAG", "CANCEL");
            idaBUDGET_ADD.Update();

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

        private void ilaBUDGET_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'ADD'", "N");
        }

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBUDGET_ADD_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaBUDGET_ADD_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Item_Status(pBindingManager.DataRow);
        }

        #endregion

    }
}