using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.IO;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0320
{
    public partial class FCMF0320 : Office2007Form
    {       
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0320()
        {
            InitializeComponent();
        }

        public FCMF0320(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void Search_DB()
        {            
            string vASSET_CIP_CODE = iConv.ISNull(IGR_ASSET_CIP_MASTER.GetCellValue("ASSET_CIP_CODE"));
            int vCOL_IDX = IGR_ASSET_CIP_MASTER.GetColumnToIndex("ASSET_CIP_CODE");

            IDA_ASSET_CIP_MASTER.Fill();
            if (iConv.ISNull(vASSET_CIP_CODE) != string.Empty)
            {
                for (int i = 0; i < IGR_ASSET_CIP_MASTER.RowCount; i++)
                {
                    if (vASSET_CIP_CODE == iConv.ISNull(IGR_ASSET_CIP_MASTER.GetCellValue(i, vCOL_IDX)))
                    {
                        IGR_ASSET_CIP_MASTER.CurrentCellMoveTo(i, vCOL_IDX);
                        IGR_ASSET_CIP_MASTER.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
            IGR_ASSET_CIP_MASTER.Focus();
        }

        private void Insert_Asset_CIP_Master()
        {
            IGR_ASSET_CIP_MASTER.SetCellValue("QTY", 1);
            IGR_ASSET_CIP_MASTER.SetCellValue("ASSET_IF_FLAG", "N");
            IGR_ASSET_CIP_MASTER.Focus();
        }

        private void Delete_Asset_CIP_Master(object pASSET_CIP_ID)
        {
            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            IDC_DELETE_ASSET_CIP_MASTER.SetCommandParamValue("W_ASSET_CIP_ID", pASSET_CIP_ID);
            IDC_DELETE_ASSET_CIP_MASTER.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_DELETE_ASSET_CIP_MASTER.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_DELETE_ASSET_CIP_MASTER.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                if(vMESSAGE!= string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            Search_DB();
        }

        private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }
           

        private void Get_DPR_Rate(object pDPR_TYPE, object pUSEFUL_TYPE, object pUSEFUL_LIFE)
        {
            IDC_DPR_RATE.SetCommandParamValue("W_DPR_TYPE", pDPR_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_TYPE", pUSEFUL_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_LIFE", pUSEFUL_LIFE);
            IDC_DPR_RATE.ExecuteNonQuery();
            DPR_RATE.EditValue = IDC_DPR_RATE.GetCommandParamValue("O_DPR_RATE"); 
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
                    if (IDA_ASSET_CIP_MASTER.IsFocused)
                    {
                        IDA_ASSET_CIP_MASTER.AddOver();
                        Insert_Asset_CIP_Master();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ASSET_CIP_MASTER.IsFocused)
                    {
                        IDA_ASSET_CIP_MASTER.AddUnder();
                        Insert_Asset_CIP_Master();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ASSET_CIP_MASTER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_CIP_MASTER.IsFocused)
                    {
                        IDA_ASSET_CIP_DTL.Cancel();
                        IDA_ASSET_CIP_MASTER.Cancel();
                    }
                    else if (IDA_ASSET_CIP_DTL.IsFocused)
                    {
                        IDA_ASSET_CIP_DTL.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ASSET_CIP_MASTER.IsFocused)
                    {
                        if(IDA_ASSET_CIP_MASTER.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_ASSET_CIP_MASTER.Delete();
                        }
                        else
                        {
                            Delete_Asset_CIP_Master(IDA_ASSET_CIP_MASTER.CurrentRow["ASSET_CIP_ID"]);
                        }
                        
                    } 
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        
        private void FCMF0320_Load(object sender, EventArgs e)
        {
            W_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_ASSET_IF_FLAG.EditValue = W_ALL.RadioButtonString;
            GB_IF_FLAG.BringToFront();

            IDA_ASSET_CIP_MASTER.FillSchema(); 
        }

        private void FCMF0320_Shown(object sender, EventArgs e)
        {
             
        }

        private void USEFUL_LIFE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Get_DPR_Rate(DPR_TYPE.EditValue, USEFUL_TYPE.EditValue, USEFUL_LIFE.EditValue);
        }

        private void W_CIP_ALL_Click(object sender, EventArgs e)
        {
            if (W_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_ASSET_IF_FLAG.EditValue = W_ALL.RadioButtonString;
            }
        }

        private void W_CIP_NO_Click(object sender, EventArgs e)
        {
            if (W_ASSET_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_ASSET_IF_FLAG.EditValue = W_ASSET_NO.RadioButtonString;
            }
        }

        private void W_CIP_YES_Click(object sender, EventArgs e)
        {
            if (W_ASSET_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_ASSET_IF_FLAG.EditValue = W_ASSET_YES.RadioButtonString;
            }
        }
         
        #endregion
        
        #region ----- Lookup Event -----

        private void ILA_ASSET_CATEGORY_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY_W.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ILD_ASSET_CATEGORY_W.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY_W.SetLookupParamValue("W_ENABLED_YN", "N");
        }
         
        private void ILA_ASSET_CATEGORY_INPUT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY_INPUT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ASSET_CATEGORY_INPUT_SelectedRowData(object pSender)
        {
            IDC_BOOK_ASSET_CATE_DPR_RATE_P.SetCommandParamValue("W_AST_CATEGORY_ID", IGR_ASSET_CIP_MASTER.GetCellValue("AST_CATEGORY_ID"));
            IDC_BOOK_ASSET_CATE_DPR_RATE_P.ExecuteNonQuery();
            DPR_TYPE.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_DPR_TYPE");
            DPR_TYPE_DESC.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_DPR_TYPE_NAME");
            DPR_BEGIN_MONTH.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_DPR_BEGIN_AMONT");
            USEFUL_TYPE.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_USEFUL_TYPE");
            USEFUL_TYPE_DESC.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_USEFUL_TYPE_NAME");
            USEFUL_LIFE.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_USEFUL_LIFE");
            DPR_RATE.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_DPR_RATE");
            LAST_BOOK_RATE.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_LAST_BOOK_RATE");
            LAST_BOOK_AMOUNT.EditValue = IDC_BOOK_ASSET_CATE_DPR_RATE_P.GetCommandParamValue("O_LAST_BOOK_AMOUNT"); 
        }

        private void ILA_COSTCENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_EXPENSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("EXPENSE_TYPE", "Y");
        }

        private void ILA_MANAGE_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_DEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_USE_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_DEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
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

        private void ILA_USEFUL_TYPE_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("USEFUL_TYPE", "Y");
        }

        private void ILA_DPR_TYPE_SelectedRowData(object pSender)
        {
            Get_DPR_Rate(DPR_TYPE.EditValue, USEFUL_TYPE.EditValue, USEFUL_LIFE.EditValue);
        }

        private void ILA_USEFUL_TYPE_SelectedRowData(object pSender)
        {
            Get_DPR_Rate(DPR_TYPE.EditValue, USEFUL_TYPE.EditValue, USEFUL_LIFE.EditValue);
        }

        #endregion

        #region ----- Adapeter Event -----

        private void IDA_ASSET_CIP_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ASSET_CIP_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10201"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; 
                return;
            } 
            if (iConv.ISNull(e.Row["AST_CATEGORY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10093"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; 
                return;
            }
            if (iConv.ISNull(e.Row["QTY"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10203"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; 
                return;
            } 
        }

        private void IDA_ASSET_CIP_MASTER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                if (iConv.ISNull(e.Row["ASSET_CIP_CODE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10209"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        #endregion

    }
}