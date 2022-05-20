using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0322
{
    public partial class FCMF0322 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();
         
        #endregion;

        #region ----- Constructor -----

        public FCMF0322()
        {
            InitializeComponent();
        }

        public FCMF0322(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
         
        private void Sync_BTN_Status(string pSLIP_FLAG)
        {
            if(pSLIP_FLAG == "N")
            { 
                BTN_SET_ASSET.Enabled = true;
                BTN_SET_DIST.Enabled = true;
            }
            else if(pSLIP_FLAG == "Y")
            {
                BTN_SET_DIST.Enabled = false;
                BTN_SET_ASSET.Enabled = false;
            }
            else
            {
                BTN_SET_DIST.Enabled = false;
                BTN_SET_ASSET.Enabled = false;
            }
        }

        private void Search_DB()
        {
            string vASSET_CIP_CODE = iConv.ISNull(IGR_ASSET_CIP_LIST.GetCellValue("ASSET_CIP_CODE"));
            int vCOL_IDX = IGR_ASSET_CIP_LIST.GetColumnToIndex("ASSET_CIP_CODE");

            IGR_ASSET_CIP_DIST.LastConfirmChanges();
            IDA_ASSET_CIP_DIST.OraSelectData.AcceptChanges();
            IDA_ASSET_CIP_DIST.Refillable = true;

            IGR_ASSET_CIP_LIST.LastConfirmChanges();
            IDA_ASSET_CIP_LIST.OraSelectData.AcceptChanges();
            IDA_ASSET_CIP_LIST.Refillable = true;

            IDA_ASSET_CIP_LIST.Fill();
            if (iConv.ISNull(vASSET_CIP_CODE) != string.Empty)
            {
                for (int i = 0; i < IGR_ASSET_CIP_LIST.RowCount; i++)
                {
                    if (vASSET_CIP_CODE == iConv.ISNull(IGR_ASSET_CIP_LIST.GetCellValue(i, vCOL_IDX)))
                    {
                        IGR_ASSET_CIP_LIST.CurrentCellMoveTo(i, vCOL_IDX);
                        IGR_ASSET_CIP_LIST.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
            IGR_ASSET_CIP_LIST.Focus();
        }

        private decimal Get_Last_Book_Amount(object pAsset_Amount, object pLast_Book_Rate)
        {
            IDC_GET_LAST_BOOK_AMOUNT.SetCommandParamValue("W_ASSET_AMOUNT", pAsset_Amount);
            IDC_GET_LAST_BOOK_AMOUNT.SetCommandParamValue("W_LAST_BOOK_RATE", pLast_Book_Rate);
            IDC_GET_LAST_BOOK_AMOUNT.ExecuteNonQuery();
            decimal vLast_Book_Amount = iConv.ISDecimaltoZero(IDC_GET_LAST_BOOK_AMOUNT.GetCommandParamValue("O_LAST_BOOK_AMOUNT"));
            return vLast_Book_Amount;
        } 

        private void Init_Total_Amount()
        {            
            decimal vAsset_Amount = Convert.ToDecimal(0); 
            foreach (DataRow vRow in IDA_ASSET_CIP_DIST.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Deleted)
                {
                    vAsset_Amount = vAsset_Amount + iConv.ISDecimaltoZero(vRow["ASSET_AMOUNT"]);  
                }
            }
            V_DIST_SUM_AMOUNT.EditValue = iConv.ISDecimaltoZero(vAsset_Amount);

            V_GAP_AMOUNT.EditValue = Math.Abs(iConv.ISDecimaltoZero(ASSET_AMOUNT.EditValue,0) - 
                                                iConv.ISDecimaltoZero(V_DIST_SUM_AMOUNT.EditValue,0)) * -1;
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
                    IGR_ASSET_CIP_LIST.Focus(); 
                    IDA_ASSET_CIP_LIST.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_CIP_LIST.IsFocused)
                    {
                        IDA_ASSET_CIP_DIST.Cancel();
                        IDA_ASSET_CIP_LIST.Cancel();
                    }
                    else
                    {
                        IDA_ASSET_CIP_DIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                { 

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0322_Load(object sender, EventArgs e)
        {
            W_ASSET_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_ASSET_IF_FLAG.EditValue = W_ASSET_NO.RadioButtonString;
            GB_IF_FLAG.BringToFront();

            BTN_RE_CAL_DIST.BringToFront();
            V_DIST_SUM_AMOUNT.BringToFront();
            ASSET_AMOUNT.BringToFront();
            V_GAP_AMOUNT.BringToFront();

            IDA_ASSET_CIP_LIST.FillSchema();
            IDA_ASSET_CIP_DIST.FillSchema();
        }

        private void FCMF0322_Shown(object sender, EventArgs e)
        {
             
        }

        private void IGR_ASSET_CIP_DIST_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vIDX_ASSET_AMOUNT = IGR_ASSET_CIP_DIST.GetColumnToIndex("ASSET_AMOUNT");
            int vIDX_LAST_BOOK_RATE = IGR_ASSET_CIP_DIST.GetColumnToIndex("LAST_BOOK_RATE");
            if (vIDX_ASSET_AMOUNT == e.ColIndex)
            {
                IGR_ASSET_CIP_DIST.LastConfirmChanges();
                Init_Total_Amount();
                decimal vLAST_BOOK_RATE = iConv.ISDecimaltoZero(IGR_ASSET_CIP_DIST.GetCellValue("LAST_BOOK_RATE"), 0);
                if(vLAST_BOOK_RATE == 0)
                {
                    return;
                }
                decimal vLAST_BOOK_AMOUNT = Get_Last_Book_Amount(e.NewValue, vLAST_BOOK_RATE);
                IGR_ASSET_CIP_DIST.SetCellValue("LAST_BOOK_AMOUNT", vLAST_BOOK_AMOUNT);
            }
            else if (vIDX_LAST_BOOK_RATE == e.ColIndex)
            {
                if (iConv.ISDecimaltoZero(e.NewValue) == 0)
                {
                    return;
                }
                decimal vASSET_AMOUNT = iConv.ISDecimaltoZero(IGR_ASSET_CIP_DIST.GetCellValue("ASSET_AMOUNT"), 0);
                decimal vLAST_BOOK_AMOUNT = Get_Last_Book_Amount(vASSET_AMOUNT, e.NewValue);

                IGR_ASSET_CIP_DIST.SetCellValue("LAST_BOOK_AMOUNT", vLAST_BOOK_AMOUNT);
            } 
        }

        private void W_CIP_ALL_Click(object sender, EventArgs e)
        {
            if (W_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_ASSET_IF_FLAG.EditValue = W_ALL.RadioButtonString;
                Sync_BTN_Status("ALL");
            }
        }

        private void W_CIP_NO_Click(object sender, EventArgs e)
        {
            if (W_ASSET_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_ASSET_IF_FLAG.EditValue = W_ASSET_NO.RadioButtonString;
                Sync_BTN_Status("N");
            }
        }

        private void W_CIP_YES_Click(object sender, EventArgs e)
        {
            if (W_ASSET_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_ASSET_IF_FLAG.EditValue = W_ASSET_YES.RadioButtonString;
                Sync_BTN_Status("Y");
            }
        }
         
        private void BTN_SET_DIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {                          
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            int vREC_CNT = 0;

            IDA_ASSET_CIP_LIST.Update();

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vSYS_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE"));

            int vIDX_SELECT_FLAG = IGR_ASSET_CIP_LIST.GetColumnToIndex("SELECT_FLAG");
            int vIDX_ASSET_CIP_ID = IGR_ASSET_CIP_LIST.GetColumnToIndex("ASSET_CIP_ID"); 
             
            for (int r = 0; r < IGR_ASSET_CIP_LIST.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_ASSET_CIP_LIST.GetCellValue(r, vIDX_SELECT_FLAG)))
                {
                    vREC_CNT++;
                    IGR_ASSET_CIP_LIST.CurrentCellMoveTo(r, vIDX_SELECT_FLAG);
                    IGR_ASSET_CIP_LIST.CurrentCellActivate(r, vIDX_SELECT_FLAG);

                    //전표 일괄생성 패키지에서 처리하므로 폼에서는 제어할수 없음
                    //해당 체크박스는 숨김 처리
                    IDC_SET_ASSET_CIP_DIST.SetCommandParamValue("W_ASSET_CIP_ID", IGR_ASSET_CIP_LIST.GetCellValue(r, vIDX_ASSET_CIP_ID)); 
                    IDC_SET_ASSET_CIP_DIST.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_ASSET_CIP_DIST.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_ASSET_CIP_DIST.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_ASSET_CIP_DIST.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_SET_ASSET_CIP_DIST.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    IGR_ASSET_CIP_LIST.SetCellValue(r, vIDX_SELECT_FLAG, "N");

                    IGR_ASSET_CIP_LIST.LastConfirmChanges();
                    IDA_ASSET_CIP_LIST.OraSelectData.AcceptChanges();
                    IDA_ASSET_CIP_LIST.Refillable = true;
                }
            } 
            if(vREC_CNT == 0)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                return;
            }
             

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();   
        }

        private void BTN_SET_ASSET_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IDA_ASSET_CIP_LIST.Update();

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vSYS_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE"));

            int vIDX_SELECT_FLAG = IGR_ASSET_CIP_LIST.GetColumnToIndex("SELECT_FLAG");
            int vIDX_ASSET_CIP_ID = IGR_ASSET_CIP_LIST.GetColumnToIndex("ASSET_CIP_ID");

            for (int r = 0; r < IGR_ASSET_CIP_LIST.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_ASSET_CIP_LIST.GetCellValue(r, vIDX_SELECT_FLAG)))
                {
                    IGR_ASSET_CIP_LIST.CurrentCellMoveTo(r, vIDX_SELECT_FLAG);
                    IGR_ASSET_CIP_LIST.CurrentCellActivate(r, vIDX_SELECT_FLAG);

                    if (iConv.ISNull(IGR_ASSET_CIP_LIST.GetCellValue("ASSET_CIP_ID")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10037"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    //전표 일괄생성 패키지에서 처리하므로 폼에서는 제어할수 없음
                    //해당 체크박스는 숨김 처리
                    IDC_TRANS_CIP_ASSET_MASTER.SetCommandParamValue("W_ASSET_CIP_ID", IGR_ASSET_CIP_LIST.GetCellValue(r, vIDX_ASSET_CIP_ID));
                    IDC_TRANS_CIP_ASSET_MASTER.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_TRANS_CIP_ASSET_MASTER.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_TRANS_CIP_ASSET_MASTER.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_TRANS_CIP_ASSET_MASTER.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_TRANS_CIP_ASSET_MASTER.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    IGR_ASSET_CIP_LIST.SetCellValue(r, vIDX_SELECT_FLAG, "N");

                    IGR_ASSET_CIP_LIST.LastConfirmChanges();
                    IDA_ASSET_CIP_LIST.OraSelectData.AcceptChanges();
                    IDA_ASSET_CIP_LIST.Refillable = true;
                }
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents(); 

            Search_DB(); 
        }

        private void BTN_RE_CAL_DIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_ASSET_CIP_DIST.RowCount < 0)
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

            IDA_ASSET_CIP_LIST.Update();

            object vASSET_CIP_ID = IGR_ASSET_CIP_LIST.GetCellValue("ASSET_CIP_ID");
            IDC_CAL_ASSET_CIP_DIST_AMT.SetCommandParamValue("W_ASSET_CIP_ID", vASSET_CIP_ID);
            IDC_CAL_ASSET_CIP_DIST_AMT.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CAL_ASSET_CIP_DIST_AMT.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CAL_ASSET_CIP_DIST_AMT.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CAL_ASSET_CIP_DIST_AMT.ExcuteError)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(IDC_CAL_ASSET_CIP_DIST_AMT.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            IDA_ASSET_CIP_DIST.Fill();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- Lookup Event -----

        private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void ILA_ASSET_CATEGORY_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY_W.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ILD_ASSET_CATEGORY_W.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY_W.SetLookupParamValue("W_ENABLED_YN", "N");
        }
         
        private void ILA_COSTCENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_EXPENSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("EXPENSE_TYPE", "Y");
        }
         
        private void ILA_USEFUL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("USEFUL_TYPE", "Y");
        }

        private void ILA_VENDOR_SUPPLIER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR_SUPPLIER.SetLookupParamValue("W_SUPP_CUST_TYPE", "S");
            ILD_VENDOR_SUPPLIER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VENDOR_MAKER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR_MAKER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_OPERATION_DIVISION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("OPERATION_DIVISION", "Y");
        }

        private void ILA_DPR_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("DPR_TYPE", "Y");
        }
         
        private void ILA_COSTCENTER_D_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_EXPENSE_TYPE_D_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("EXPENSE_TYPE", "Y");
        }

        private void ILA_USEFUL_TYPE_D_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("USEFUL_TYPE", "Y");
        }

        private void ILA_OPERATION_DIVISION_D_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("OPERATION_DIVISION", "Y");
        }

        private void ILA_DPR_TYPE_D_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("DPR_TYPE", "Y");
        }
         
        #endregion

        #region ----- Adapter Event -----

        private void IDA_ASSET_CIP_DIST_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Total_Amount();
        }

        private void IDA_ASSET_CIP_LIST_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["DPR_YN"]) == "Y" && iConv.ISNull(e.Row["DPR_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10221"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DPR_YN"]) == "Y" && iConv.ISDecimaltoZero(e.Row["USEFUL_LIFE"]) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DPR_YN"]) == "N" && iConv.ISNull(e.Row["DPR_TYPE"]) != String.Empty)
            {
                int vIDX_DPR_YN = IGR_ASSET_CIP_LIST.GetColumnToIndex("DPR_YN");
                int vIDX_DPR_TYPE = IGR_ASSET_CIP_LIST.GetColumnToIndex("DPR_TYPE_DESC");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10217", String.Format("&&VALUE1:={0} &&VALUE2:={1}", Get_Grid_Prompt(IGR_ASSET_CIP_LIST, vIDX_DPR_YN), Get_Grid_Prompt(IGR_ASSET_CIP_LIST, vIDX_DPR_TYPE))), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DPR_YN"]) == "N" && iConv.ISDecimaltoZero(e.Row["USEFUL_LIFE"]) != 0)
            {
                int vIDX_DPR_YN = IGR_ASSET_CIP_LIST.GetColumnToIndex("DPR_YN");
                int vIDX_DPR_USEFUL_LIFE = IGR_ASSET_CIP_LIST.GetColumnToIndex("USEFUL_LIFE");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10217", String.Format("&&VALUE1:={0} &&VALUE2:={1}", Get_Grid_Prompt(IGR_ASSET_CIP_LIST, vIDX_DPR_YN), Get_Grid_Prompt(IGR_ASSET_CIP_LIST, vIDX_DPR_USEFUL_LIFE))), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_CMS_SLIP_DETAIL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {// 계정과목.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["MANAGEMENT1"]) == string.Empty && iConv.ISNull(e.Row["MANAGEMENT1_YN"], "N") == "Y".ToString())
            {// 관리항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["MANAGEMENT1_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["MANAGEMENT2"]) == string.Empty && iConv.ISNull(e.Row["MANAGEMENT2_YN"], "N") == "Y".ToString())
            {// 관리항목2 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["MANAGEMENT2_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER1"]) == string.Empty && iConv.ISNull(e.Row["REFER1_YN"], "N") == "Y".ToString())
            {// 참고항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER1_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER2"]) == string.Empty && iConv.ISNull(e.Row["REFER2_YN"], "N") == "Y".ToString())
            {// 참고항목2 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER2_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER3"]) == string.Empty && iConv.ISNull(e.Row["REFER3_YN"], "N") == "Y".ToString())
            {// 참고항목3 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER3_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER4"]) == string.Empty && iConv.ISNull(e.Row["REFER4_YN"], "N") == "Y".ToString())
            {// 참고항목4 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER4_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER5"]) == string.Empty && iConv.ISNull(e.Row["REFER5_YN"], "N") == "Y".ToString())
            {// 참고항목5 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER5_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER6"]) == string.Empty && iConv.ISNull(e.Row["REFER6_YN"], "N") == "Y".ToString())
            {// 참고항목6 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER6_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER7"]) == string.Empty && iConv.ISNull(e.Row["REFER7_YN"], "N") == "Y".ToString())
            {// 참고항목7 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER7_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER8"]) == string.Empty && iConv.ISNull(e.Row["REFER8_YN"], "N") == "Y".ToString())
            {// 참고항목8 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER8_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}