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

namespace FCMF0343
{
    public partial class FCMF0343 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0343()
        {
            InitializeComponent();
        }

        public FCMF0343(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Set_DefaultValues()
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_PERIOD_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_PERIOD_DATE_TO.EditValue = iDate.ISMonth_Last(DateTime.Today);

            // INTERFACE STATUS //
            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "INTERFACE_STATUS");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            if (IDC_DEFAULT_VALUE.ExcuteError)
            {
                return;
            }
            W_INTERFACE_FLAG.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
            W_INTERFACE_FLAG_DESC.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
        }

        private void Set_BTN_Visible(int pTab_Index)
        {
            if (pTab_Index == 0)
            {
                BTN_ASSET_INTERFACE.Visible = false;
                BTN_GET_ASSET_LIST.Visible = true;                
            }
            else if (pTab_Index == 1)
            {
                BTN_GET_ASSET_LIST.Visible = false;                
                BTN_ASSET_INTERFACE.Visible = true;                
            }
        }

        private void Search_DB()
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_DATE_FR.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_DATE_TO.Focus();
                return;
            }

            if (TB_MAIN.SelectedIndex == 0)
            {
                string vASSET_DESC = iConv.ISNull(IGR_ASSET_IF_LIST.GetCellValue("ASSET_DESC"));
                int vIDX_ASSET_DESC = IGR_ASSET_IF_LIST.GetColumnToIndex("ASSET_DESC");

                IDA_ASSET_IF_LIST.Fill();

                if (iConv.ISNull(vASSET_DESC) != string.Empty)
                {
                    for (int i = 0; i < IGR_ASSET_IF_LIST.RowCount; i++)
                    {
                        if (vASSET_DESC == iConv.ISNull(IGR_ASSET_IF_LIST.GetCellValue(i, vIDX_ASSET_DESC)))
                        {
                            IGR_ASSET_IF_LIST.CurrentCellMoveTo(i, vIDX_ASSET_DESC);
                            return;
                        }
                    }
                }
            }
            else if (TB_MAIN.SelectedIndex == 1)
            {
                string vASSET_DESC = iConv.ISNull(IGR_ASSET_INTERFACE.GetCellValue("ASSET_DESC"));
                int vIDX_ASSET_DESC = IGR_ASSET_INTERFACE.GetColumnToIndex("ASSET_DESC");

                IDA_ASSET_INTERFACE.Fill();
                
                if (iConv.ISNull(vASSET_DESC) != string.Empty)
                {
                    for (int i = 0; i < IGR_ASSET_INTERFACE.RowCount; i++)
                    {
                        if (vASSET_DESC == iConv.ISNull(IGR_ASSET_INTERFACE.GetCellValue(i, vIDX_ASSET_DESC)))
                        {
                            IGR_ASSET_INTERFACE.CurrentCellMoveTo(i, vIDX_ASSET_DESC);
                            return;
                        }
                    }
                }
            }            
        }

        private decimal Get_Last_Book_Amount(object pAsset_Amount, object pLast_Book_Rate)
        {
            IDC_GET_LAST_BOOK_AMOUNT.SetCommandParamValue("W_ASSET_AMOUNT", pAsset_Amount);
            IDC_GET_LAST_BOOK_AMOUNT.SetCommandParamValue("W_LAST_BOOK_RATE", pLast_Book_Rate);
            IDC_GET_LAST_BOOK_AMOUNT.ExecuteNonQuery();
            decimal vLast_Book_Amount = iConv.ISDecimaltoZero(IDC_GET_LAST_BOOK_AMOUNT.GetCommandParamValue("O_LAST_BOOK_AMOUNT")); 
            return vLast_Book_Amount;
        }

        private void Get_DPR_Rate(object pDPR_TYPE, object pUSEFUL_TYPE, object pUSEFUL_LIFE)
        {
            IDC_DPR_RATE.SetCommandParamValue("W_DPR_TYPE", pDPR_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_TYPE", pUSEFUL_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_LIFE", pUSEFUL_LIFE);
            IDC_DPR_RATE.ExecuteNonQuery();
            IGR_ASSET_IF_LIST.SetCellValue("DPR_RATE", IDC_DPR_RATE.GetCommandParamValue("O_DPR_RATE"));
        }

        private void Get_IFRS_DPR_Rate(object pIFRS_DPR_TYPE, object pIFRS_USEFUL_TYPE, object pIFRS_USEFUL_LIFE)
        {
            IDC_DPR_RATE.SetCommandParamValue("W_DPR_TYPE", pIFRS_DPR_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_TYPE", pIFRS_USEFUL_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_LIFE", pIFRS_USEFUL_LIFE);
            IDC_DPR_RATE.ExecuteNonQuery();
            IGR_ASSET_IF_LIST.SetCellValue("IFRS_DPR_RATE", IDC_DPR_RATE.GetCommandParamValue("O_DPR_RATE"));
        }

        #endregion;

        #region ----- Set Parameter -----

        private void SetCommonPara(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        #endregion

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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_ASSET_IF_LIST.IsFocused)
                    {
                        IDA_ASSET_IF_LIST.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_IF_LIST.IsFocused)
                    {
                        IDA_ASSET_IF_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ASSET_IF_LIST.IsFocused)
                    {
                        IDA_ASSET_IF_LIST.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        
        private void FCMF0343_Load(object sender, EventArgs e)
        {
            Set_DefaultValues();
            IDA_ASSET_IF_LIST.FillSchema();
        }

        private void FCMF0343_Shown(object sender, EventArgs e)
        {
            Set_BTN_Visible(TB_MAIN.SelectedIndex);
        }

        private void TB_MAIN_Click(object sender, EventArgs e)
        {
            Set_BTN_Visible(TB_MAIN.SelectedIndex);
        }

        private void IGR_ASSET_IF_LIST_CellKeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                if (IGR_ASSET_IF_LIST.ColIndex == IGR_ASSET_IF_LIST.GetColumnToIndex("ASSET_DESC"))
                {
                    IGR_ASSET_IF_LIST.CurrentCellMoveTo(IGR_ASSET_IF_LIST.GetColumnToIndex("AST_CATEGORY_NAME"));
                    IGR_ASSET_IF_LIST.CurrentCellMoveTo(IGR_ASSET_IF_LIST.GetColumnToIndex("ASSET_DESC"));
                }
            }
        }

        private void IGR_ASSET_IF_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_LAST_BOOK_RATE = IGR_ASSET_IF_LIST.GetColumnToIndex("DPR_LAST_BOOK_RATE");
            decimal vASSET_AMOUNT = 0;
            if (vIDX_LAST_BOOK_RATE == e.ColIndex)
            {
                vASSET_AMOUNT = iConv.ISDecimaltoZero(IGR_ASSET_IF_LIST.GetCellValue("AMOUNT"), 0);
                decimal vLAST_BOOK_AMOUNT = Get_Last_Book_Amount(vASSET_AMOUNT, e.NewValue);

                IGR_ASSET_IF_LIST.SetCellValue("DPR_LAST_BOOK_AMOUNT", vLAST_BOOK_AMOUNT);
            }

            int vIDX_IFRS_LAST_BOOK_RATE = IGR_ASSET_IF_LIST.GetColumnToIndex("IFRS_DPR_LAST_BOOK_RATE");
            if (vIDX_IFRS_LAST_BOOK_RATE == e.ColIndex)
            {
                vASSET_AMOUNT = iConv.ISDecimaltoZero(IGR_ASSET_IF_LIST.GetCellValue("AMOUNT"), 0); 
                decimal vLAST_BOOK_AMOUNT = Get_Last_Book_Amount(vASSET_AMOUNT, e.NewValue);

                IGR_ASSET_IF_LIST.SetCellValue("IFRS_DPR_LAST_BOOK_AMOUNT", vLAST_BOOK_AMOUNT);
            }
        }

        private void IGR_ASSET_IF_LIST_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            int vIDX_USEFUL_TYPE = IGR_ASSET_IF_LIST.GetColumnToIndex("USEFUL_TYPE");
            int vIDX_USEFUL_LIFE = IGR_ASSET_IF_LIST.GetColumnToIndex("USEFUL_LIFE");
            int vIDX_IFRS_USEFUL_TYPE = IGR_ASSET_IF_LIST.GetColumnToIndex("IFRS_USEFUL_TYPE");
            int vIDX_IFRS_USEFUL_LIFE = IGR_ASSET_IF_LIST.GetColumnToIndex("IFRS_USEFUL_LIFE");

            if (e.ColIndex == vIDX_USEFUL_TYPE)
            {
                Get_DPR_Rate(IGR_ASSET_IF_LIST.GetCellValue("DPR_TYPE"), e.CellValue, IGR_ASSET_IF_LIST.GetCellValue("USEFUL_LIFE"));
            }
            else if (e.ColIndex == vIDX_USEFUL_LIFE)
            {
                Get_DPR_Rate(IGR_ASSET_IF_LIST.GetCellValue("DPR_TYPE"), IGR_ASSET_IF_LIST.GetCellValue("USEFUL_TYPE"), e.CellValue);
            }
            else if (e.ColIndex == vIDX_IFRS_USEFUL_TYPE)
            {
                Get_IFRS_DPR_Rate(IGR_ASSET_IF_LIST.GetCellValue("IFRS_DPR_TYPE"), e.CellValue, IGR_ASSET_IF_LIST.GetCellValue("IFRS_USEFUL_LIFE"));
            }
            else if (e.ColIndex == vIDX_IFRS_USEFUL_LIFE)
            {
                Get_IFRS_DPR_Rate(IGR_ASSET_IF_LIST.GetCellValue("IFRS_DPR_TYPE"), IGR_ASSET_IF_LIST.GetCellValue("IFRS_USEFUL_TYPE"), e.CellValue);
            }
        }

        private void IGR_ASSET_INTERFACE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_ASSET_INTERFACE.GetColumnToIndex("SELECT_YN"))
            {
                IGR_ASSET_INTERFACE.LastConfirmChanges();
                IDA_ASSET_INTERFACE.OraSelectData.AcceptChanges();
                IDA_ASSET_INTERFACE.Refillable = true;
            }
        }

        private void CB_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IGR_ASSET_INTERFACE.RowCount < 1)
            {
                return;
            }

            int vIDX_SELECT_YN = IGR_ASSET_INTERFACE.GetColumnToIndex("SELECT_YN");
            int vIDX_INTERFACE_FLAG = IGR_ASSET_INTERFACE.GetColumnToIndex("INTERFACE_FLAG");
            for(int vRow =0;vRow < IGR_ASSET_INTERFACE.RowCount;vRow++)
            {
                if (iConv.ISNull(IGR_ASSET_INTERFACE.GetCellValue(vRow, vIDX_INTERFACE_FLAG)) == "Y")
                {

                }
                else
                {
                    IGR_ASSET_INTERFACE.SetCellValue(vRow, vIDX_SELECT_YN, CB_SELECT_YN.CheckBoxValue);
                }
            }
            IGR_ASSET_INTERFACE.LastConfirmChanges();
            IDA_ASSET_INTERFACE.OraSelectData.AcceptChanges();
            IDA_ASSET_INTERFACE.Refillable = true;
        }

        private void BTN_GET_ASSET_LIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_DATE_FR.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_DATE_TO.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IDC_GET_SLIP_ASSET_LIST.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_GET_SLIP_ASSET_LIST.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_GET_SLIP_ASSET_LIST.GetCommandParamValue("O_MESSAGE"));
            if (IDC_GET_SLIP_ASSET_LIST.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            // 다시 조회 //
            Search_DB();
        }

        private void BTN_ASSET_INTERFACE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            //1. 전송대상 FLAG UPDATE //
            int vIDX_SELECT_YN = IGR_ASSET_INTERFACE.GetColumnToIndex("SELECT_YN");
            int vIDX_SOURCE_TABLE= IGR_ASSET_INTERFACE.GetColumnToIndex("SOURCE_TABLE");
            int vIDX_SOURCE_LINE_ID = IGR_ASSET_INTERFACE.GetColumnToIndex("SOURCE_LINE_ID");
            int vIDX_ASSET_DESC = IGR_ASSET_INTERFACE.GetColumnToIndex("ASSET_DESC");

            isDataTransaction1.BeginTran();
            for (int vRow = 0; vRow < IGR_ASSET_INTERFACE.RowCount; vRow++)
            {
                if (iConv.ISNull(IGR_ASSET_INTERFACE.GetCellValue(vRow, vIDX_SELECT_YN)) == "Y")
                {
                    IDC_UPDATE_FLAG_ASSET_IF.SetCommandParamValue("P_SOURCE_TABLE", IGR_ASSET_INTERFACE.GetCellValue(vRow, vIDX_SOURCE_TABLE));
                    IDC_UPDATE_FLAG_ASSET_IF.SetCommandParamValue("P_SOURCE_LINE_ID", IGR_ASSET_INTERFACE.GetCellValue(vRow, vIDX_SOURCE_LINE_ID));
                    IDC_UPDATE_FLAG_ASSET_IF.SetCommandParamValue("P_ASSET_DESC", IGR_ASSET_INTERFACE.GetCellValue(vRow, vIDX_ASSET_DESC));
                    IDC_UPDATE_FLAG_ASSET_IF.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_UPDATE_FLAG_ASSET_IF.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_UPDATE_FLAG_ASSET_IF.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_UPDATE_FLAG_ASSET_IF.ExcuteError || vSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }
            
            //2. 선택 자산대장 전송 //
            IDC_SET_IF_ASSET_MASTER.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_IF_ASSET_MASTER.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_IF_ASSET_MASTER.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_IF_ASSET_MASTER.ExcuteError || vSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            // 다시 조회 //
            Search_DB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_W_PERIOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", "2010-01");
        }

        private void ILA_W_ASSET_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", null);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_W_INTERFACE_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonPara("INTERFACE_STATUS", "N");
        }

        private void ILA_ASSET_CATEGORY_1ST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY_1ST.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ASSET_CATE_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATE_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ASSET_CODE_CIP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
             
        }

        private void ILA_ASSET_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonPara("ASSET_STATUS", "Y");
        }

        private void ILA_EXPENSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonPara("EXPENSE_TYPE", "Y");
        }

        private void ILA_SUPP_CUST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SUPP_CUST.SetLookupParamValue("W_SUPP_CUST_TYPE", "S");
            ILD_SUPP_CUST.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_FLOOR_MANAGE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FLOOR_CC.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_FLOOR_USE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FLOOR_CC.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_COST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_OPERATION_DIVISION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonPara("OPERATION_DIVISION", "Y");
        }

        private void ILA_ASSET_CHARGE_CIP_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_CHARGE_CIP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_USEFUL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonPara("USEFUL_TYPE", "Y");
        }

        private void ILA_IFRS_USEFUL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonPara("USEFUL_TYPE", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_ASSET_IF_LIST_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["AST_CATE_1ST_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10101"), "Error");
                e.Cancel = true;
                return;
            } 
            if (iConv.ISNull(e.Row["ASSET_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10201"), "Error");
                e.Cancel = true;
                return;
            }
            //if (iConv.ISNull(e.Row["AST_CATEGORY_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10095"), "Error");
            //    e.Cancel = true;
                //return;
            //}
            if (iConv.ISNull(e.Row["EXPENSE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10216"), "Error");
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ACQUIRE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10203"), "Error");
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10208"), "Error");
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ASSET_STATUS_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10410"), "Error");
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DPR_YN"]) == "Y" && iConv.ISDecimaltoZero(e.Row["DPR_USEFUL_LIFE"]) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Error");
                e.Cancel = true;
                return;
            } 
            if (iConv.ISNull(e.Row["IFRS_DPR_YN"]) == "Y" && iConv.ISDecimaltoZero(e.Row["IFRS_DPR_USEFUL_LIFE"]) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Error");
                e.Cancel = true;
                return;
            }
        }

        private void IDA_ASSET_INTERFACE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            int vIDX_SELECT_YN = IGR_ASSET_INTERFACE.GetColumnToIndex("SELECT_YN");
            if (iConv.ISNull(pBindingManager.DataRow["INTERFACE_FLAG"]) == "Y")
            {
                IGR_ASSET_INTERFACE.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 0;
                IGR_ASSET_INTERFACE.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 0;
            }
            else
            {
                IGR_ASSET_INTERFACE.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 1;
                IGR_ASSET_INTERFACE.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 1;
            }
        }


        #endregion

    }
}