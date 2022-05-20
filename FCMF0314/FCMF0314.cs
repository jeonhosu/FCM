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

namespace FCMF0314
{
    public partial class FCMF0314 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0314()
        {
            InitializeComponent();
        }

        public FCMF0314(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            V_SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            IGR_ASSET_MASTER_CIP.LastConfirmChanges();
            IDA_ASSET_MASTER_CIP.OraSelectData.AcceptChanges();
            IDA_ASSET_MASTER_CIP.Refillable = true;

            IDA_ASSET_MASTER_CIP.Fill();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_ASSET_MASTER_CIP.IsFocused)
                    {
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ASSET_MASTER_CIP.IsFocused)
                    {
                         
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ASSET_HISTORY_CIP.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_MASTER_CIP.IsFocused)
                    {
                        IDA_ASSET_MASTER_CIP.Cancel();
                        IDA_ASSET_HISTORY_CIP.Cancel();
                    }
                    else if (IDA_ASSET_HISTORY_CIP.IsFocused)
                    {
                        IDA_ASSET_HISTORY_CIP.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
            }
        }

        #endregion;

        #region ----- Form Event ----

        private void FCMF0314_Load(object sender, EventArgs e)
        {
            W_REGISTER_PERIOD.EditValue = iDate.ISYearMonth(iDate.ISGetDate());
            W_CIP_YES.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_CIP_FLAG.EditValue = W_CIP_YES.RadioCheckedString;

            BTN_CANCEL_ASSET.Enabled = false;

            IDA_ASSET_MASTER_CIP.FillSchema();
            IDA_ASSET_HISTORY_CIP.FillSchema();
        }

        private void W_CIP_YES_Click(object sender, EventArgs e)
        {
            if (W_CIP_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_YES.RadioCheckedString;

                BTN_SET_ASSET.Enabled = true;
                BTN_CANCEL_ASSET.Enabled = false;
            }
        }

        private void W_CIP_NO_Click(object sender, EventArgs e)
        {
            if (W_CIP_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_NO.RadioCheckedString;

                BTN_SET_ASSET.Enabled = false;
                BTN_CANCEL_ASSET.Enabled = true;
            }
        }

        private void V_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IGR_ASSET_MASTER_CIP.RowCount < 1)
            {
                return;
            }

            int vIDX_SELECT_YN = IGR_ASSET_MASTER_CIP.GetColumnToIndex("SELECT_YN"); 
            for (int vRow = 0; vRow < IGR_ASSET_MASTER_CIP.RowCount; vRow++)
            {
                if (iConv.ISNull(IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_SELECT_YN)) == iConv.ISNull(V_SELECT_YN.CheckBoxValue))
                {
                   
                }
                else
                {
                    IGR_ASSET_MASTER_CIP.SetCellValue(vRow, vIDX_SELECT_YN, V_SELECT_YN.CheckBoxValue);
                }
            }
            IGR_ASSET_MASTER_CIP.LastConfirmChanges();
            IDA_ASSET_MASTER_CIP.OraSelectData.AcceptChanges(); 
            IDA_ASSET_MASTER_CIP.Refillable = true;
        }

        private void BTN_SET_ASSET_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_ASSET_MASTER_CIP.RowCount < 1)
            {
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_YN = IGR_ASSET_MASTER_CIP.GetColumnToIndex("SELECT_YN");
            int vIDX_ASSET_ID = IGR_ASSET_MASTER_CIP.GetColumnToIndex("ASSET_ID");
            int vIDX_REPLACE_DATE = IGR_ASSET_MASTER_CIP.GetColumnToIndex("REPLACE_DATE");
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IGR_ASSET_MASTER_CIP.LastConfirmChanges();
            IDA_ASSET_MASTER_CIP.OraSelectData.AcceptChanges();
            IDA_ASSET_MASTER_CIP.Refillable = true;

            //TRANSACTION BEGIN
            isDataTransaction1.BeginTran();

            //대상 선택//
            for (int vRow = 0; vRow < IGR_ASSET_MASTER_CIP.RowCount; vRow++)
            {
                if (iConv.ISNull(IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_SELECT_YN)) == "Y")
                {
                    IDC_SET_ASSET_MASTER_CIP.SetCommandParamValue("W_SELECT_YN", IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_SELECT_YN));
                    IDC_SET_ASSET_MASTER_CIP.SetCommandParamValue("W_ASSET_ID", IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_ASSET_ID));
                    IDC_SET_ASSET_MASTER_CIP.SetCommandParamValue("W_REPLACE_DATE", IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_REPLACE_DATE));
                    IDC_SET_ASSET_MASTER_CIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_ASSET_MASTER_CIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_ASSET_MASTER_CIP.GetCommandParamValue("O_MESSAGE"));
                    if(IDC_SET_ASSET_MASTER_CIP.ExcuteError)
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_SET_ASSET_MASTER_CIP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if(vSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            } 

            //commit
            isDataTransaction1.Commit();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void BTN_CANCEL_ASSET_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_ASSET_MASTER_CIP.RowCount < 1)
            {
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_YN = IGR_ASSET_MASTER_CIP.GetColumnToIndex("SELECT_YN");
            int vIDX_ASSET_ID = IGR_ASSET_MASTER_CIP.GetColumnToIndex("ASSET_ID");
            int vIDX_REPLACE_DATE = IGR_ASSET_MASTER_CIP.GetColumnToIndex("REPLACE_DATE");
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IGR_ASSET_MASTER_CIP.LastConfirmChanges();
            IDA_ASSET_MASTER_CIP.OraSelectData.AcceptChanges();
            IDA_ASSET_MASTER_CIP.Refillable = true;

            //TRANSACTION BEGIN
            isDataTransaction1.BeginTran();

            //대상 선택//
            for (int vRow = 0; vRow < IGR_ASSET_MASTER_CIP.RowCount; vRow++)
            {
                if (iConv.ISNull(IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_SELECT_YN)) == "Y")
                {
                    IDC_SET_ASSET_MASTER_CIP.SetCommandParamValue("W_SELECT_YN", IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_SELECT_YN));
                    IDC_SET_ASSET_MASTER_CIP.SetCommandParamValue("W_ASSET_ID", IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_ASSET_ID));
                    IDC_SET_ASSET_MASTER_CIP.SetCommandParamValue("W_REPLACE_DATE", IGR_ASSET_MASTER_CIP.GetCellValue(vRow, vIDX_REPLACE_DATE)); 
                    IDC_SET_ASSET_MASTER_CIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_ASSET_MASTER_CIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_ASSET_MASTER_CIP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_ASSET_MASTER_CIP.ExcuteError)
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_SET_ASSET_MASTER_CIP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            //commit
            isDataTransaction1.Commit();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_PERIOD_NAME_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            
        }

        private void ilaASSET_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_ASSET_CHARGE_CIP_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_CHARGE_CIP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion


        #region ----- Adapter Event -----

        private void IDA_ASSET_HISTORY_CIP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["CHARGE_DATE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10223"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["CHARGE_ID"]) == string.Empty)
            {
                e.Cancel=true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10224"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion
    }
}