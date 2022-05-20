using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0961
{
    public partial class FCMF0961 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0961()
        {
            InitializeComponent();
        }

        public FCMF0961(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (TB_DIST.SelectedTab.TabIndex == TP_DIST_OP.TabIndex)
            {
                IDA_OPERATION_DIST_OP.Fill();
                ISG_OPERATION_DIST_OP.Focus();
            }
            else if(TB_DIST.SelectedTab.TabIndex == TP_DIST_CC.TabIndex)
            {
                IDA_OPERATION_DIST_CC.Fill();
                ISG_OPERATION_DIST_CC.Focus();
            }
        }

        private void SearchDB_OP( object pPERIOD_NAME
                                , object pACCOUNT_CONTROL_ID, object pOPERATION_ADJ_RULE_ID
                                , object pCOST_CENTER_ID, object pOPERATION_DIVISION)
        {
            IDA_OPERATION_DIST_OP_SLIP.SetSelectParamValue("W_PERIOD_NAME", pPERIOD_NAME);
            IDA_OPERATION_DIST_OP_SLIP.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_OPERATION_DIST_OP_SLIP.SetSelectParamValue("W_OPERATION_ADJ_RULE_ID", pOPERATION_ADJ_RULE_ID);
            IDA_OPERATION_DIST_OP_SLIP.SetSelectParamValue("W_COST_CENTER_ID", pCOST_CENTER_ID);
            IDA_OPERATION_DIST_OP_SLIP.SetSelectParamValue("W_OPERATION_DIVISION", pOPERATION_DIVISION); 
            IDA_OPERATION_DIST_OP_SLIP.Fill();
        }

        private void SearchDB_CC(object pPERIOD_NAME
                                , object pACCOUNT_CONTROL_ID, object pOPERATION_ADJ_RULE_ID
                                , object pCOST_CENTER_ID, object pOPERATION_DIVISION)
        {
            IDA_OPERATION_DIST_CC_SLIP.SetSelectParamValue("W_PERIOD_NAME", pPERIOD_NAME);
            IDA_OPERATION_DIST_CC_SLIP.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_OPERATION_DIST_CC_SLIP.SetSelectParamValue("W_OPERATION_ADJ_RULE_ID", pOPERATION_ADJ_RULE_ID);
            IDA_OPERATION_DIST_CC_SLIP.SetSelectParamValue("W_COST_CENTER_ID", pCOST_CENTER_ID);
            IDA_OPERATION_DIST_CC_SLIP.SetSelectParamValue("W_OPERATION_DIVISION", pOPERATION_DIVISION);
            IDA_OPERATION_DIST_CC_SLIP.Fill();
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
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_OPERATION_DIST_OP.IsFocused)
                    {
                        IDA_OPERATION_DIST_OP.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
            }
        }

        #endregion;

        #region ---- From Event -----


        private void FCMF0961_Load(object sender, EventArgs e)
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(iDate.ISGetDate());
        }

        private void FCMF0961_Shown(object sender, EventArgs e)
        {

        }

        private void BTN_EXEC_DIST_RATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;
            
            IDC_SET_OPERATION_DIST_RATE.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_OPERATION_DIST_RATE.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_OPERATION_DIST_RATE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_OPERATION_DIST_RATE.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            SearchDB();
        }

        private void BTN_EXEC_DIST_AMOUNT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_SET_OPERATION_DIST_AMOUNT.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_OPERATION_DIST_AMOUNT.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_OPERATION_DIST_AMOUNT.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_OPERATION_DIST_AMOUNT.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            SearchDB();
        }

        private void BTN_EXEC_DIST_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            
            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_SET_OPERATION_DIST_CLOSED.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_OPERATION_DIST_CLOSED.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_OPERATION_DIST_CLOSED.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_OPERATION_DIST_CLOSED.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            SearchDB();
        }

        private void BTN_CANCEL_DIST_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_CANCEL_OPERATION_DIST_CLOSED.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_CANCEL_OPERATION_DIST_CLOSED.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_CANCEL_OPERATION_DIST_CLOSED.GetCommandParamValue("O_MESSAGE"));
            
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_OPERATION_DIST_CLOSED.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            SearchDB();
        }

        private void BTN_EXEC_DIST_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            
            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_SET_OPERATION_DIST_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_OPERATION_DIST_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_OPERATION_DIST_SLIP.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_OPERATION_DIST_SLIP.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            SearchDB();
        }

        private void BTN_CANCEL_DIST_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_CANCEL_OPERATION_DIST_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_CANCEL_OPERATION_DIST_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_CANCEL_OPERATION_DIST_SLIP.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_OPERATION_DIST_SLIP.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            SearchDB();
        }

        #endregion


        #region ----- Lookup Event -----

        private void ILA_PERIOD_NAME_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD_NAME_W.SetLookupParamValue("W_START_YYYYMM", "2010-01");
        }

        private void ILA_ACCOUNT_CONTROL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
             ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_OPERATION_ADJ_RULE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OPERATION_ADJ_RULE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_OPERATION_ADJ_RULE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OPERATION_ADJ_RULE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_OPERATION_DIST_OP_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                SearchDB_OP(W_PERIOD_NAME.EditValue, -1, -1, -1, "-1");
            }
            else
            {
                SearchDB_OP(pBindingManager.DataRow["PERIOD_NAME"]         //PERIOD_NAME
                            , pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]  //ACCOUNT_CONTROL_ID
                            , pBindingManager.DataRow["OPERATION_ADJ_RULE_ID"]  //OPERATION_ADJ_RULE_ID
                            , DBNull.Value  // COST_CENTER_ID
                            , DBNull.Value  //OPERATION_DIVISION 
                            );
            }
        }

        private void IDA_OPERATION_DIST_CC_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                SearchDB_CC(W_PERIOD_NAME.EditValue, -1, -1, -1, "-1");
            }
            else
            {
                SearchDB_CC(pBindingManager.DataRow["PERIOD_NAME"]         //PERIOD_NAME
                            , pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]  //ACCOUNT_CONTROL_ID
                            , pBindingManager.DataRow["OPERATION_ADJ_RULE_ID"]  //OPERATION_ADJ_RULE_ID
                            , pBindingManager.DataRow["COST_CENTER_ID"]  // COST_CENTER_ID
                            , DBNull.Value  //OPERATION_DIVISION 
                            );
            }
        }
         
        #endregion


    }
}