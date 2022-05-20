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

namespace FCMF0912
{
    public partial class FCMF0912 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0912()
        {
            InitializeComponent();
        }

        public FCMF0912(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void SearchDB()
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {// 예산부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            idaENDING_AMOUNT.Fill();
            idaCLOSING_AMOUNT.Fill();
            idaCLOSING_SLIP.Fill(); 

            SLIP_SUM();

            if (itbCLOSING.SelectedIndex == 0)
            {
                igrENDING_AMOUNT.Focus();
            }
            else if (itbCLOSING.SelectedIndex == 1)
            {
                igrCLOSING_AMOUNT.Focus();
            }
            else if (itbCLOSING.SelectedIndex == 2)
            {
                igrCLOSING_SLIP.Focus();
            }
            else if(itbCLOSING.SelectedTab.TabIndex == TP_ETC.TabIndex)
            {
                IDA_ETC_TRX_AMOUNT.Fill();
            }
            else if (itbCLOSING.SelectedTab.TabIndex == TP_ETC_SLIP.TabIndex)
            {
                IDA_ETC_TRX_SLIP.Fill();
                ETC_TRX_SLIP_SUM();
            }
        }

        private void SLIP_SUM()
        {
            // 분개 합계 표시
            idcCLOSING_SLIP_SUM.ExecuteNonQuery();
            DR_AMOUNT.EditValue = idcCLOSING_SLIP_SUM.GetCommandParamValue("O_DR_AMOUNT");
            CR_AMOUNT.EditValue = idcCLOSING_SLIP_SUM.GetCommandParamValue("O_CR_AMOUNT");
            GAP_AMOUNT.EditValue = idcCLOSING_SLIP_SUM.GetCommandParamValue("O_GAP_AMOUNT");
        }

        private void ETC_TRX_SLIP_SUM()
        {
            // 기타 분개 합계 표시
            IDC_ETC_TRX_SLIP_SUM_P.ExecuteNonQuery();
            V_ETC_DR_AMOUNT.EditValue = IDC_ETC_TRX_SLIP_SUM_P.GetCommandParamValue("O_DR_AMOUNT");
            V_ETC_CR_AMOUNT.EditValue = IDC_ETC_TRX_SLIP_SUM_P.GetCommandParamValue("O_CR_AMOUNT");
            V_ETC_GAP_AMOUNT.EditValue = IDC_ETC_TRX_SLIP_SUM_P.GetCommandParamValue("O_GAP_AMOUNT");
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetInsert_EndingAmount()
        {
            igrENDING_AMOUNT.SetCellValue("PERIOD_NAME", W_PERIOD_NAME.EditValue);

            igrENDING_AMOUNT.CurrentCellMoveTo(igrENDING_AMOUNT.GetColumnToIndex("ACCOUNT_CODE"));
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
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
                    //if (idaENDING_AMOUNT.IsFocused)
                    //{
                    //    idaENDING_AMOUNT.AddOver();
                    //    SetInsert_EndingAmount();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if (idaENDING_AMOUNT.IsFocused)
                    //{
                    //    idaENDING_AMOUNT.AddUnder();
                    //    SetInsert_EndingAmount();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaENDING_AMOUNT.IsFocused)
                    {
                        idaENDING_AMOUNT.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaENDING_AMOUNT.IsFocused)
                    {
                        idaENDING_AMOUNT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaENDING_AMOUNT.IsFocused)
                    {
                        idaENDING_AMOUNT.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0912_Load(object sender, EventArgs e)
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);

            object vOPERATION_DIVISION_FLAG = 0;
            IDC_GET_OPERATION_DIV_FLAG_P.ExecuteNonQuery();
            string vOPERATION_DIV_FLAG = iString.ISNull(IDC_GET_OPERATION_DIV_FLAG_P.GetCommandParamValue("O_OPERATION_DIV_FLAG"));
            if (vOPERATION_DIV_FLAG == "Y")
            {
                W_OPERATION_DIV_NAME.Visible = true;
                vOPERATION_DIVISION_FLAG = 1;
            }
            else
            {
                W_OPERATION_DIV_NAME.Visible = false;
                vOPERATION_DIVISION_FLAG = 0;
            } 
            int vIDX_OPERATION_DIVISION_NAME = igrENDING_AMOUNT.GetColumnToIndex("OPERATION_DIVISION_NAME");
            igrENDING_AMOUNT.GridAdvExColElement[vIDX_OPERATION_DIVISION_NAME].Visible = vOPERATION_DIVISION_FLAG;
            igrENDING_AMOUNT.ResetDraw = true;

            vIDX_OPERATION_DIVISION_NAME = igrCLOSING_AMOUNT.GetColumnToIndex("OPERATION_DIVISION_NAME");
            igrCLOSING_AMOUNT.GridAdvExColElement[vIDX_OPERATION_DIVISION_NAME].Visible = vOPERATION_DIVISION_FLAG;
            igrCLOSING_AMOUNT.ResetDraw = true;

            vIDX_OPERATION_DIVISION_NAME = igrCLOSING_SLIP.GetColumnToIndex("OPERATION_DIVISION_NAME");
            igrCLOSING_SLIP.GridAdvExColElement[vIDX_OPERATION_DIVISION_NAME].Visible = vOPERATION_DIVISION_FLAG;
            igrCLOSING_SLIP.ResetDraw = true;
        }

        private void FCMF0912_Shown(object sender, EventArgs e)
        {
            idaCLOSING_AMOUNT.FillSchema();
            idaCLOSING_SLIP.FillSchema();
            idaENDING_AMOUNT.FillSchema();
        }

        private void BTN_CST_ENDING_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string vStatus = "F";
            string vMessage = string.Empty;
            IDC_INIT_CLSOING_ENDING_AMOUNT.ExecuteNonQuery();
            vStatus = iString.ISNull(IDC_INIT_CLSOING_ENDING_AMOUNT.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(IDC_INIT_CLSOING_ENDING_AMOUNT.GetCommandParamValue("O_MESSAGE"));

            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            if (IDC_INIT_CLSOING_ENDING_AMOUNT.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_INIT_CLSOING_ENDING_AMOUNT.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (iString.ISNull(vMessage) != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            idaENDING_AMOUNT.Fill();
        }

        private void BTN_GET_ETC_TRX_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string vStatus = "F";
            string vMessage = string.Empty;
            IDC_MISC_ACCT_CREATE.ExecuteNonQuery();
            vStatus = iString.ISNull(IDC_MISC_ACCT_CREATE.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(IDC_MISC_ACCT_CREATE.GetCommandParamValue("O_MESSAGE"));

            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            if (IDC_MISC_ACCT_CREATE.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_INIT_CLSOING_ENDING_AMOUNT.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (iString.ISNull(vMessage) != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            IDA_ETC_TRX_AMOUNT.Fill();
        }

        private void ibtnSET_CLOSING_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string vStatus = "F";
            string vMessage = string.Empty;
            idcCLOSING_SET.ExecuteNonQuery();
            vStatus = iString.ISNull(idcCLOSING_SET.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(idcCLOSING_SET.GetCommandParamValue("O_MESSAGE"));
            
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            if (idcCLOSING_SET.ExcuteError)
            {
                MessageBoxAdv.Show(idcCLOSING_SET.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (iString.ISNull(vMessage) != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            SearchDB();
        }

        private void btnCREATE_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string vStatus = "F";
            string vMessage = string.Empty;
            idcINSERT_CLOSING_SLIP.ExecuteNonQuery();
            vStatus = iString.ISNull(idcINSERT_CLOSING_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(idcINSERT_CLOSING_SLIP.GetCommandParamValue("O_MESSAGE"));

            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            if (idcINSERT_CLOSING_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(idcINSERT_CLOSING_SLIP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if (iString.ISNull(vMessage) != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            idaCLOSING_AMOUNT.Fill();
        }

        private void btnDELETE_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string vStatus = "F";
            string vMessage = string.Empty;

            idcDELETE_CLOSING_SLIP.ExecuteNonQuery();
            vStatus = iString.ISNull(idcDELETE_CLOSING_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(idcDELETE_CLOSING_SLIP.GetCommandParamValue("O_MESSAGE"));

            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            if (idcDELETE_CLOSING_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(idcDELETE_CLOSING_SLIP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if (iString.ISNull(vMessage) != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            idaCLOSING_SLIP.Fill();
        }

        private void BTN_CANCEL_ETC_TRX_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_CANCEL_ETC_TRX_SLIP.ExecuteNonQuery();
            vStatus = iString.ISNull(IDC_CANCEL_ETC_TRX_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(IDC_CANCEL_ETC_TRX_SLIP.GetCommandParamValue("O_MESSAGE"));

            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            if (IDC_CANCEL_ETC_TRX_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(idcDELETE_CLOSING_SLIP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if (iString.ISNull(vMessage) != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            IDA_ETC_TRX_SLIP.Fill();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaCLOSING_ENDING_ACCOUNT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCLOSING_ENDING_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCLOSING_ACCOUNT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CLOSING_ACCOUNT_TYPE", "Y");
        }

        private void ilaPERIOD_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildPERIOD.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 6)));
        }

        private void ilaPERIOD_SelectedRowData(object pSender)
        {
            SearchDB();
        }

        private void ILA_OPERATION_DIVISION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OPERATION_DIVISION.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_TRANSACTION_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_TRANSACTION_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CLOSING_GROUP_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CLOSING_GROUP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaCLOSING_ACCOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERIOD_NAME"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrENDING_AMOUNT.CurrentCellMoveTo(igrENDING_AMOUNT.GetColumnToIndex("ACCOUNT_CODE"));
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrENDING_AMOUNT.CurrentCellMoveTo(igrENDING_AMOUNT.GetColumnToIndex("ACCOUNT_CODE"));
                return;
            }
            if (iString.ISNull(e.Row["ENDING_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10074", "&&VALUE:=Ending Amount(기말금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrENDING_AMOUNT.CurrentCellMoveTo(igrENDING_AMOUNT.GetColumnToIndex("ENDING_AMOUNT"));
                return;
            }
        }

        private void idaCLOSING_ACCOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["PERIOD_NAME"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10226"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrENDING_AMOUNT.CurrentCellMoveTo(igrENDING_AMOUNT.GetColumnToIndex("ACCOUNT_CODE"));
                return;
            }
        }

        #endregion

    }
}