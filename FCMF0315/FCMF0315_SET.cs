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

namespace FCMF0315
{
    public partial class FCMF0315_SET : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
         
        #endregion;

        #region ----- Constructor -----

        public FCMF0315_SET()
        {
            InitializeComponent();
        }

        public FCMF0315_SET(Form pMainForm, ISAppInterface pAppInterface 
                            , object pSALE_HEADER_ID, object pSALE_NUM, object pSALE_DATE, object pSALE_AMOUNT)
        {
            InitializeComponent(); 
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_SALE_HEADER_ID.EditValue = pSALE_HEADER_ID;            
            V_SALE_NUM.EditValue = pSALE_NUM;
            V_SALE_DATE.EditValue = pSALE_DATE;
            V_SALE_AMOUNT.EditValue = pSALE_AMOUNT; 
        }

        #endregion;

        #region ----- Private Methods ----
         
        private void Search_DB()
        {
            if (iConv.ISNull(V_SALE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_SALE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(V_SALE_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
            V_SELECTED.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            IDA_ASSET_SALE_LIST.Fill();
            IGR_ASSET_SALE_LIST.Focus();
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_GRID_STATUS(object pSELECTED_FLAG, object pMODIFY_YN)
        {
            int vSTATUS = 0;                // INSERTABLE, UPDATABLE; 
            int vIDX_SALE_AMOUNT = IGR_ASSET_SALE_LIST.GetColumnToIndex("SALE_AMOUNT");
            int vIDX_DESCRIPTION = IGR_ASSET_SALE_LIST.GetColumnToIndex("DESCRIPTION");

            if (iConv.ISNull(pSELECTED_FLAG) == "N")
            {
                vSTATUS = 0;
            }
            else
            {
                if (iConv.ISNull(pMODIFY_YN) == "Y")
                {
                    vSTATUS = 1;
                }
                else
                {
                    vSTATUS = 0;
                }
            }

            IGR_ASSET_SALE_LIST.GridAdvExColElement[vIDX_SALE_AMOUNT].Insertable = vSTATUS;
            IGR_ASSET_SALE_LIST.GridAdvExColElement[vIDX_SALE_AMOUNT].Updatable = vSTATUS;

            IGR_ASSET_SALE_LIST.GridAdvExColElement[vIDX_DESCRIPTION].Insertable = vSTATUS;
            IGR_ASSET_SALE_LIST.GridAdvExColElement[vIDX_DESCRIPTION].Updatable = vSTATUS;
            
            // 범위를 지정해서 LOOP 이용//
            //int mGRID_START_COL = 17;   // 그리드 시작 COLUMN INDEX.
            //int mMax_Column = 24;       // 종료 COLUMN INDEX.

            //if (iConvert.ISNull(pSELECT_FLAG) == "Y")
            //{
            //    vSTATUS = 1;
            //}
            //else
            //{
            //    vSTATUS = 0;
            //}

            //for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            //{
            //    IGR_SCM_WYFC.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Insertable = vSTATUS;
            //    IGR_SCM_WYFC.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Updatable = vSTATUS;
            //}
        }

        private void Init_Sale_Amount()
        {
            decimal vSale_Amount = 0;
            foreach (System.Data.DataRow vRow in IDA_ASSET_SALE_LIST.CurrentRows)
            {
                if (iConv.ISNull(vRow["SELECTED_YN"]) == "Y")
                {
                    vSale_Amount = vSale_Amount +
                                    iConv.ISDecimaltoZero(vRow["SALE_AMOUNT"]);
                } 
            }
            V_SELECTED_SALE_AMOUNT.EditValue = vSale_Amount;
        }

        private bool Save_Asset_Sale_List()
        {             
            IGR_ASSET_SALE_LIST.LastConfirmChanges();
            IDA_ASSET_SALE_LIST.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_LIST.Refillable = true;

            string vStatus = "F";
            string vMessage = string.Empty;
            foreach (System.Data.DataRow vRow in IDA_ASSET_SALE_LIST.CurrentRows)
            {
                if (iConv.ISNull(vRow["SELECTED_YN"]) == "Y")
                {
                    if (iConv.ISDecimaltoZero(vRow["SALE_AMOUNT"], 0) == 0)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10592"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }

                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_SELECTED_YN", vRow["SELECTED_YN"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_DPR_TYPE", vRow["DPR_TYPE"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_ASSET_ID", vRow["ASSET_ID"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_ASSET_AMOUNT", vRow["ASSET_AMOUNT"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_DPR_PERIOD_NAME", vRow["DPR_PERIOD_NAME"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_DPR_SUM_AMOUNT", vRow["DPR_SUM_AMOUNT"]); 
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_BOOK_AMOUNT", vRow["BOOK_AMOUNT"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_SALE_AMOUNT", vRow["SALE_AMOUNT"]);
                    IDC_SAVE_SALE_ASSET_LIST.SetCommandParamValue("P_DESCRIPTION", vRow["DESCRIPTION"]);
                    IDC_SAVE_SALE_ASSET_LIST.ExecuteNonQuery();
                    vStatus = iConv.ISNull(IDC_SAVE_SALE_ASSET_LIST.GetCommandParamValue("O_STATUS"));
                    vMessage = iConv.ISNull(IDC_SAVE_SALE_ASSET_LIST.GetCommandParamValue("O_MESSAGE"));

                    if (IDC_SAVE_SALE_ASSET_LIST.ExcuteError)
                    {
                        if (IDC_SAVE_SALE_ASSET_LIST.ExcuteErrorMsg != string.Empty)
                        {
                            MessageBoxAdv.Show(IDC_SAVE_SALE_ASSET_LIST.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return false;
                    }
                    else if (vStatus == "F")
                    {
                        if (vMessage != string.Empty)
                        {
                            MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return false;
                    } 
                }
            } 
            return true;
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event ----

        private void FCMF0315_SET_Load(object sender, EventArgs e)
        {
            V_SALE_DATE.BringToFront();
            V_SALE_NUM.BringToFront();

            IDA_ASSET_SALE_LIST.FillSchema(); 
        }

        private void BTN_INQUIRY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Search_DB();
        }

        private void BTN_SELECTED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Save_Asset_Sale_List() == true)
            {
                IDC_CAL_PROFIT_AMOUNT.ExecuteNonQuery();
                string vSTATUS = iConv.ISNull(IDC_CAL_PROFIT_AMOUNT.GetCommandParamValue("O_STATUS"));
                string vMESSAGE = iConv.ISNull(IDC_CAL_PROFIT_AMOUNT.GetCommandParamValue("O_MESSAGE"));
                if (IDC_CAL_PROFIT_AMOUNT.ExcuteError)
                {
                    MessageBoxAdv.Show(IDC_CAL_PROFIT_AMOUNT.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else if(vSTATUS == "F")
                {
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    return;
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void IGR_ASSET_SALE_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (IGR_ASSET_SALE_LIST.RowIndex < 0)
            {
                return;
            }
             
            IGR_ASSET_SALE_LIST.LastConfirmChanges();
            IDA_ASSET_SALE_LIST.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_LIST.Refillable = true;

            Set_GRID_STATUS(IGR_ASSET_SALE_LIST.GetCellValue("SELECTED_YN"), IGR_ASSET_SALE_LIST.GetCellValue("MODIFY_YN"));

            Init_Sale_Amount();
        }


        private void V_SELECTED_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IGR_ASSET_SALE_LIST.RowCount < 1)
            {
                return;
            }

            int vIDX_SELECTED_YN = IGR_ASSET_SALE_LIST.GetColumnToIndex("SELECTED_YN");
            for (int vRow = 0; vRow < IGR_ASSET_SALE_LIST.RowCount; vRow++)
            {
                if (iConv.ISNull(IGR_ASSET_SALE_LIST.GetCellValue(vRow, vIDX_SELECTED_YN)) == iConv.ISNull(V_SELECTED.CheckBoxValue))
                {

                }
                else
                {
                    IGR_ASSET_SALE_LIST.SetCellValue(vRow, vIDX_SELECTED_YN, V_SELECTED.CheckBoxValue);
                }
            }  
            IGR_ASSET_SALE_LIST.LastConfirmChanges();
            IDA_ASSET_SALE_LIST.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_LIST.Refillable = true;

            Set_GRID_STATUS(IGR_ASSET_SALE_LIST.GetCellValue("SELECTED_YN"), IGR_ASSET_SALE_LIST.GetCellValue("MODIFY_YN")); 
        }
         
        private void FCMF0315_SET_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        #endregion

        #region ----- Lookup Event -----
         
        private void ILA_VENDOR_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }	 
        
        private void ILA_VENDOR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_ASSET_CATEGORY_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_MANAGE_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_DEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_COSTCENTER_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_DPR_TYPE_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("DPR_TYPE", "Y");
        }

        private void ILA_ASSET_TYPE_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("ASSET_TYPE", "Y");
        }

        private void ILA_EXPENSE_TYPE_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("EXPENSE_TYPE", "Y");
        }

        #endregion

        private void IDA_ASSET_SALE_LIST_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Set_GRID_STATUS("N", "N");
                return;
            }
            Set_GRID_STATUS(pBindingManager.DataRow["SELECTED_YN"], pBindingManager.DataRow["MODIFY_YN"]); 
        }

        #region ----- Adapter Event -----
          
        #endregion


    }
}