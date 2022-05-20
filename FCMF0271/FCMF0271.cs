using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;

using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
 

namespace FCMF0271
{
    public partial class FCMF0271 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mCurrency_Code;
        object mBudget_Control_YN; 

        #endregion;
        
        #region ----- Constructor -----

        public FCMF0271(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void GetAccountBook()
        {
            IDC_ACCOUNT_BOOK.ExecuteNonQuery();
            mSession_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_SESSION_ID");
            mAccount_Book_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = iString.ISNull(IDC_ACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private void Search()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_PREPAID_EXP_LIST.TabIndex)
            {
                IDA_PREPAID_EXPENSE_LIST.Fill();
                IGR_PREPAID_EXPENSE_LIST.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_PREPAID_EXP_MONTH.TabIndex)
            {
                if (iString.ISNull(W2_PERIOD_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_PERIOD_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W2_PERIOD_FR.Focus();
                    return;
                }
                if (iString.ISNull(W2_PERIOD_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_PERIOD_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W2_PERIOD_TO.Focus();
                    return;
                }
                IDA_PREPAID_EXP_MONTH.Fill();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        //조회된 자료에서 더블클릭하면 전표팝업 띄워준다.
        private void Show_Slip_Detail(int pSLIP_HEADER_ID)
        {
            if (pSLIP_HEADER_ID != 0)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, pSLIP_HEADER_ID);
                vFCMF0204.Show();

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                Application.UseWaitCursor = false;
            }
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

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)        //검색
                {
                    Search();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)  //위에 새레코드 추가
                {
                                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder) //아래에 새레코드 추가
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)   //저장
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)   //취소
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)   //삭제
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)    //인쇄
                {
                    //XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                {
                    //XLPrinting("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0271_Load(object sender, EventArgs e)
        {
            IDC_GET_DEFAULT_VALUE_GROUP.SetCommandParamValue("W_GROUP_CODE", "PREPAID_EXP_STATUS");
            IDC_GET_DEFAULT_VALUE_GROUP.ExecuteNonQuery();
            W_PREPAID_EXPENSE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE");
            W_PREPAID_EXPENSE_STATUS_DESC.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE_NAME");

            W2_PREPAID_EXPENSE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE");
            W2_PREPAID_EXPENSE_STATUS_DESC.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE_NAME");          
        }

        private void FCMF0271_Shown(object sender, EventArgs e)
        {
            GetAccountBook();

            W2_PERIOD_FR.EditValue = iDate.ISYearMonth(string.Format("{0}-01", iDate.ISYear(DateTime.Today)));
            W2_PERIOD_TO.EditValue = iDate.ISYearMonth(DateTime.Today);
        }
         
        private void IGR_PREPAID_EXP_HISTORY_CellDoubleClick(object pSender)
        {
            if (IGR_PREPAID_EXP_HISTORY.RowIndex > -1)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_PREPAID_EXP_HISTORY.GetCellValue("SLIP_HEADER_ID"));
                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        #endregion


        #region ----- Lookup Event -----

        private void ILA_COSTCENTER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_W2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_PREPAID_EXP_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_STATUS", "Y");
        }

        private void ILA_PREPAID_EXP_STATUS_W2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_STATUS", "Y");
        }

        private void ILA_ENTERED_METHOD_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("ENTERED_METHOD", "Y");
        }

        private void ILA_ENTERED_METHOD_W2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("ENTERED_METHOD", "Y");
        }

        private void ILA_PREPAID_EXP_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_TYPE", "Y");
        }

        private void ILA_PREPAID_EXP_TYPE_W2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_TYPE", "Y");
        }

        private void ILA_PREPAID_EXP_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_TYPE", "Y");
        }

        private void ILA_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_ENTRY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_CLASS.SetLookupParamValue("W_ACCOUNT_CLASS_TYPE", "PREPAID");
            ILD_ACCOUNT_CONTROL_CLASS.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_REPLACE_ACCOUNT_CONTROL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_EXP_SPREAD_METHOD_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("EXP_SPREAD_METHOD", "Y");
        }

        private void ILA_PREPAID_EXP_STATUS_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_STATUS", "Y");
        }
         
        private void ILA_PERIOD_FR_2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", null);
        }

        private void ILA_PERIOD_FR_2_SelectedRowData(object pSender)
        {
            W2_PERIOD_TO.EditValue = W2_PERIOD_FR.EditValue;
        }

        private void ILA_PERIOD_TO_2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", W2_PERIOD_FR.EditValue);
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddYears(6));
        } 

        #endregion

        #region ----- Adapter Lookup Event -----
                    
        #endregion

        

    }
}