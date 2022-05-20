using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0284
{
    public partial class FCMF0284 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        bool mSearch_Flag = false;

        #endregion;

        #region ----- Constructor -----

        public FCMF0284()
        {
            InitializeComponent();
        }

        public FCMF0284(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                if (iConv.ISNull(W_REG_DATE_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_REG_DATE_FR.Focus();
                    return;
                }

                if (iConv.ISNull(W_REG_DATE_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_REG_DATE_TO.Focus();
                    return;
                }

                if (mSearch_Flag == false)
                    return;

                IDA_CMS_TRAN_LIST.Fill();
                IGR_CMS_TRAN_LIST.Focus();
            }
            else if(TB_MAIN.SelectedTab.TabIndex == TP_DTL.TabIndex)
            {
                Search_DB_DTL(CMS_TRAN_HEADER_ID.EditValue);
            }
        }

        private void Search_DB_DTL(object pCMS_TRAN_HEADER_ID)
        {
            if (iConv.ISNull(pCMS_TRAN_HEADER_ID) == string.Empty)
            {
                return;
            }

            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            IDA_CMS_TRAN_HEADER.SetSelectParamValue("W_CMS_TRAN_HEADER_ID", pCMS_TRAN_HEADER_ID);
            IDA_CMS_TRAN_HEADER.Fill();

            REG_DATE.Focus();
        }

        private void Insert_Header()
        {
            REG_DATE.EditValue = iDate.ISGetDate();

            IDC_TRAN_TYPE.SetCommandParamValue("W_GROUP_CODE", "CMS_TRAN_TYPE");
            IDC_TRAN_TYPE.ExecuteNonQuery();
            TRAN_TYPE.EditValue = IDC_TRAN_TYPE.GetCommandParamValue("O_VALUE1");
            CMS_TRAN_TYPE_NAME.EditValue = IDC_TRAN_TYPE.GetCommandParamValue("O_CODE_NAME");

            IDC_CMS_REMARK_TYPE.SetCommandParamValue("W_GROUP_CODE", "CMS_REMARK_TYPE");
            IDC_CMS_REMARK_TYPE.ExecuteNonQuery();
            WDRW_REMARK_TYPE.EditValue = IDC_CMS_REMARK_TYPE.GetCommandParamValue("O_CODE");
            CMS_REMARK_TYPE_NAME.EditValue = IDC_CMS_REMARK_TYPE.GetCommandParamValue("O_CODE_NAME");

            IDC_CMS_TRAN_SET_TYPE.SetCommandParamValue("W_GROUP_CODE", "CMS_TRAN_SET_TYPE");
            IDC_CMS_TRAN_SET_TYPE.ExecuteNonQuery();
            TRAN_SET_TYPE.EditValue = IDC_CMS_TRAN_SET_TYPE.GetCommandParamValue("O_CODE");
            CMS_TRAN_SET_TYPE_NAME.EditValue = IDC_CMS_TRAN_SET_TYPE.GetCommandParamValue("O_CODE_NAME");

            IDC_CMS_KOR_BANK_USE.SetCommandParamValue("W_GROUP_CODE", "CMS_KOR_BANK_USE");
            IDC_CMS_KOR_BANK_USE.ExecuteNonQuery();
            KOR_BANK_USE_YN.EditValue = IDC_CMS_KOR_BANK_USE.GetCommandParamValue("O_CODE");
            CMS_KOR_BANK_USE_NAME.EditValue = IDC_CMS_KOR_BANK_USE.GetCommandParamValue("O_CODE_NAME");

            IDC_CMS_SAME_WDRW_ACCT_TYPE.SetCommandParamValue("W_GROUP_CODE", "CMS_SAME_WDRW_ACCT_TYPE");
            IDC_CMS_SAME_WDRW_ACCT_TYPE.ExecuteNonQuery();
            SAME_WDRW_ACCT_TYPE.EditValue = IDC_CMS_SAME_WDRW_ACCT_TYPE.GetCommandParamValue("O_CODE");
            CMS_SAME_WDRW_ACCT_TYPE_NAME.EditValue = IDC_CMS_SAME_WDRW_ACCT_TYPE.GetCommandParamValue("O_CODE_NAME");

            REG_DATE.Focus();
        }

        private void Insert_Line()
        {
            IGR_CMS_TRAN_LINE.SetCellValue("WDRW_ACCT_NO", WDRW_ACCT_NO.EditValue);
            IGR_CMS_TRAN_LINE.SetCellValue("WDRW_REMARK", WDRW_REMARK.EditValue);
            
            IDC_CMS_FEE_TYPE.SetCommandParamValue("W_GROUP_CODE", "CMS_FEE_TYPE");
            IDC_CMS_FEE_TYPE.ExecuteNonQuery();
            IGR_CMS_TRAN_LINE.SetCellValue("FEE_TYPE", IDC_CMS_FEE_TYPE.GetCommandParamValue("O_CODE"));
            IGR_CMS_TRAN_LINE.SetCellValue("CMS_FEE_TYPE_NAME", IDC_CMS_FEE_TYPE.GetCommandParamValue("O_CODE_NAME")); 

            IGR_CMS_TRAN_LINE.Focus();
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
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
                    mSearch_Flag = true;
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_CMS_TRAN_HEADER.IsFocused)
                    {
                        IDA_CMS_TRAN_HEADER.AddUnder();
                        Insert_Header();
                    }
                    else if (IDA_CMS_TRAN_LINE.IsFocused)
                    {
                        IDA_CMS_TRAN_LINE.AddUnder();
                        Insert_Line();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_CMS_TRAN_HEADER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CMS_TRAN_HEADER.IsFocused)
                    {
                        IDA_CMS_TRAN_LINE.Cancel();
                        IDA_CMS_TRAN_HEADER.Cancel();
                    }
                    else if (IDA_CMS_TRAN_LINE.IsFocused)
                    {
                        IDA_CMS_TRAN_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CMS_TRAN_LIST.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return;

                        IDC_DELETE_CMS_TRAN_LIST.SetCommandParamValue("W_CMS_TRAN_HEADER_ID", IGR_CMS_TRAN_LIST.GetCellValue("CMS_TRAN_HEADER_ID"));
                        IDC_DELETE_CMS_TRAN_LIST.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_CMS_TRAN_LIST.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_CMS_TRAN_LIST.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_DELETE_CMS_TRAN_LIST.ExcuteError)
                        {
                            MessageBoxAdv.Show(IDC_DELETE_CMS_TRAN_LIST.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if (vSTATUS.Equals("F"))
                        {
                            if (vMESSAGE != string.Empty)
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        Search_DB();
                    }
                    else if (IDA_CMS_TRAN_HEADER.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return;

                        IDC_DELETE_CMS_TRAN_LIST.SetCommandParamValue("W_CMS_TRAN_HEADER_ID", CMS_TRAN_HEADER_ID.EditValue);
                        IDC_DELETE_CMS_TRAN_LIST.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_CMS_TRAN_LIST.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_CMS_TRAN_LIST.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_DELETE_CMS_TRAN_LIST.ExcuteError)
                        {
                            MessageBoxAdv.Show(IDC_DELETE_CMS_TRAN_LIST.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if (vSTATUS.Equals("F"))
                        {
                            if (vMESSAGE != string.Empty)
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        Search_DB_DTL(CMS_TRAN_HEADER_ID.EditValue);
                    }
                    else if (IDA_CMS_TRAN_LINE.IsFocused)
                    {
                        IDA_CMS_TRAN_LINE.Delete();
                    }
                } 
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0284_Load(object sender, EventArgs e)
        {
            GB_STATUS.BringToFront();

            W_REG_DATE_FR.EditValue = iDate.ISMonth_1st(iDate.ISGetDate());
            W_REG_DATE_TO.EditValue = iDate.ISGetDate();

            mSearch_Flag = false;
            W_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_TRANS_YN.EditValue = W_RB_ALL.RadioCheckedString;            
        }

        private void FCMF0284_Shown(object sender, EventArgs e)
        {
            IDA_CMS_TRAN_LIST.FillSchema();
            IDA_CMS_TRAN_HEADER.FillSchema();
            IDA_CMS_TRAN_LINE.FillSchema();
        }

        private void W_RB_ALL_Click(object sender, EventArgs e)
        {
            if (W_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
                W_TRANS_YN.EditValue = W_RB_ALL.RadioCheckedString;
        }

        private void W_RB_OK_Click(object sender, EventArgs e)
        {
            if (W_RB_OK.CheckedState == ISUtil.Enum.CheckedState.Checked)
                W_TRANS_YN.EditValue = W_RB_OK.RadioCheckedString;
        }

        private void W_RB_NO_Click(object sender, EventArgs e)
        {
            if (W_RB_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
                W_TRANS_YN.EditValue = W_RB_NO.RadioCheckedString;
        }

        private void IGR_CMS_TRAN_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_CMS_TRAN_LIST.RowIndex < 0)
                return;

            Search_DB_DTL(IGR_CMS_TRAN_LIST.GetCellValue("CMS_TRAN_HEADER_ID"));
        }

        private void WDRW_ACCT_NO_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            object vWDRW_ACCT_NO = WDRW_ACCT_NO.EditValue;
            if (iConv.ISNull(vWDRW_ACCT_NO) == string.Empty)
                return;

            int vIDX_WDRW_ACCT_NO = IGR_CMS_TRAN_LINE.GetColumnToIndex("WDRW_ACCT_NO");
            for (int r = 0; r < IGR_CMS_TRAN_LINE.RowCount;r++)
            {
                IGR_CMS_TRAN_LINE.SetCellValue(r, vIDX_WDRW_ACCT_NO, vWDRW_ACCT_NO);    
            }
        }

        private void BTN_EXEC_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iConv.ISNull(CMS_TRAN_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(CMS_TRAN_HEADER_ID))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            IDC_SET_CMS_TRAN_CLOSED.SetCommandParamValue("P_CMS_TRAN_HEADER_ID", CMS_TRAN_HEADER_ID.EditValue);
            IDC_SET_CMS_TRAN_CLOSED.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_CMS_TRAN_CLOSED.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_CMS_TRAN_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if(IDC_SET_CMS_TRAN_CLOSED.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_CMS_TRAN_CLOSED.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS.Equals("F"))
            {
                if(vMESSAGE != string.Empty)
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; 
            }

            Search_DB_DTL(CMS_TRAN_HEADER_ID.EditValue);
        }

        private void BTN_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(CMS_TRAN_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(CMS_TRAN_HEADER_ID))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            IDC_CANCEL_CMS_TRAN_CLOSED.SetCommandParamValue("P_CMS_TRAN_HEADER_ID", CMS_TRAN_HEADER_ID.EditValue);
            IDC_CANCEL_CMS_TRAN_CLOSED.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CANCEL_CMS_TRAN_CLOSED.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CANCEL_CMS_TRAN_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CANCEL_CMS_TRAN_CLOSED.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CANCEL_CMS_TRAN_CLOSED.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS.Equals("F"))
            {
                if (vMESSAGE != string.Empty)
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Search_DB_DTL(CMS_TRAN_HEADER_ID.EditValue);
        }

        private void BTN_EXEC_CB_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(CMS_TRAN_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(CMS_TRAN_HEADER_ID))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            IDC_SET_CMS_TRAN_CYBER_BRANCH.SetCommandParamValue("P_CMS_TRAN_HEADER_ID", CMS_TRAN_HEADER_ID.EditValue);
            IDC_SET_CMS_TRAN_CYBER_BRANCH.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_CMS_TRAN_CYBER_BRANCH.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_CMS_TRAN_CYBER_BRANCH.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_CMS_TRAN_CYBER_BRANCH.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_CMS_TRAN_CYBER_BRANCH.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS.Equals("F"))
            {
                if (vMESSAGE != string.Empty)
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Search_DB_DTL(CMS_TRAN_HEADER_ID.EditValue);
        }

        private void BTN_CANCEL_CB_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(CMS_TRAN_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(CMS_TRAN_HEADER_ID))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            IDC_CANCEL_CMS_TRAN_CYBER_BRANCH.SetCommandParamValue("P_CMS_TRAN_HEADER_ID", CMS_TRAN_HEADER_ID.EditValue);
            IDC_CANCEL_CMS_TRAN_CYBER_BRANCH.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CANCEL_CMS_TRAN_CYBER_BRANCH.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CANCEL_CMS_TRAN_CYBER_BRANCH.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CANCEL_CMS_TRAN_CYBER_BRANCH.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CANCEL_CMS_TRAN_CYBER_BRANCH.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS.Equals("F"))
            {
                if (vMESSAGE != string.Empty)
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Search_DB_DTL(CMS_TRAN_HEADER_ID.EditValue);
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_CMS_TRAN_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CMS_TRAN_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CMS_PROC_STATUS_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CMS_PROC_STATUS.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CMS_REMARK_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CMS_REMARK_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_TRAN_SET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
           ILD_TRAN_SET_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VENDOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR.SetLookupParamValue("W_REG_DATE", REG_DATE.EditValue);
            ILD_VENDOR.SetLookupParamValue("W_CURRENCY_CODE", DBNull.Value);
        }

        private void ILA_REV_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CMS_FEE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CMS_FEE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_KOR_BANK_USE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_KOR_BANK_USE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_WAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_WAGE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_SAME_WDRW_ACCT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SAME_WDRW_ACCT_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_CMS_TRAN_HEADER_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
                return;
             
        }

        private void IDA_CMS_TRAN_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["REG_DATE"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REG_DATE))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["TRAN_TYPE"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CMS_TRAN_TYPE_NAME))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["TRAN_SET_TYPE"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CMS_TRAN_SET_TYPE_NAME))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["TRAN_SET_TYPE"]).Equals("R"))
            {
                if (iConv.ISNull(e.Row["TRAN_SET_DATE"]).Equals(""))
                {
                    e.Cancel = true;
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(TRAN_SET_DATE))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (iConv.ISNull(e.Row["TRAN_SET_TIME_H"]).Equals(""))
                {
                    e.Cancel = true;
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(TRAN_SET_TIME_H))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (iConv.ISNull(e.Row["TRAN_SET_TIME_M"]).Equals(""))
                {
                    e.Cancel = true;
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(TRAN_SET_TIME_M))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (iConv.ISNull(e.Row["KOR_BANK_USE_YN"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CMS_KOR_BANK_USE_NAME))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void IDA_CMS_TRAN_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["WDRW_ACCT_NO"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_CMS_TRAN_LINE, IGR_CMS_TRAN_LINE.GetColumnToIndex("WDRW_ACCT_NO")))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["RCV_BANK_CODE"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_CMS_TRAN_LINE, IGR_CMS_TRAN_LINE.GetColumnToIndex("RCV_BANK_NAME")))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["RCV_ACCT_NO"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_CMS_TRAN_LINE, IGR_CMS_TRAN_LINE.GetColumnToIndex("RCV_ACCT_NO")))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["TRAN_AMOUNT"]).Equals(""))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_CMS_TRAN_LINE, IGR_CMS_TRAN_LINE.GetColumnToIndex("TRAN_AMOUNT")))), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion

    }
}