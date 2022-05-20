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
 

namespace FCMF0272
{
    public partial class FCMF0272 : Office2007Form
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
        object mCurrency_Code;
        object mBudget_Control_YN; 

        #endregion;
        
        #region ----- Constructor -----

        public FCMF0272(Form pMainForm, ISAppInterface pAppInterface)
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
            mCurrency_Code = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE");
            mBudget_Control_YN = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private void SearchDB()
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iString.ISNull(W_EXP_SPREAD_METHOD_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_EXP_SPREAD_METHOD_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_EXP_SPREAD_METHOD_CODE.Focus();
                return;
            }

            IDA_PREPAID_EXP_SLIP.Fill();   
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }
        
        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10300"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_EXP_SPREAD_METHOD_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_EXP_SPREAD_METHOD_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_EXP_SPREAD_METHOD_CODE.Focus();
                return;
            }
            if (iString.ISNull(V_GL_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_GL_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE.Focus();
                return;
            }

            //전표생성여부 묻기.
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10303"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
              
            // 처리대상선택//
            int mCHECK_COUNT = 0;
            int mIDX_SELECT_FLAG = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("SELECT_FLAG");
            int mIDX_PREPAID_EXPENSE_ID = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("PREPAID_EXPENSE_ID");
            int mIDX_PERIOD_NAME = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("PERIOD_NAME");
            int mIDX_EXP_SPREAD_METHOD_CODE = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("EXP_SPREAD_METHOD_CODE");

            string mSYSDATE = iDate.ISGetDate().ToLongDateString();
            string mSTATUS = string.Empty;
            string mMESSAGE = string.Empty;

            for (int nRow = 0; nRow < IGR_PREPAID_EXP_SLIP.RowCount; nRow++)
            {
                if (iString.ISNull(IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_SELECT_FLAG)) == "Y")
                {
                    mCHECK_COUNT = mCHECK_COUNT + 1;

                    // 대상 UPDATE //
                    IDC_UPDATE_PREPAID_EXP_SLIP.SetCommandParamValue("P_SELECT_FLAG", IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_SELECT_FLAG));
                    IDC_UPDATE_PREPAID_EXP_SLIP.SetCommandParamValue("P_PREPAID_EXPENSE_ID", IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_PREPAID_EXPENSE_ID));
                    IDC_UPDATE_PREPAID_EXP_SLIP.SetCommandParamValue("P_PERIOD_NAME", IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_PERIOD_NAME));
                    IDC_UPDATE_PREPAID_EXP_SLIP.SetCommandParamValue("P_EXP_SPREAD_METHOD_CODE", IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_EXP_SPREAD_METHOD_CODE));
                    IDC_UPDATE_PREPAID_EXP_SLIP.SetCommandParamValue("P_SYSDATE", mSYSDATE);
                    IDC_UPDATE_PREPAID_EXP_SLIP.ExecuteNonQuery();

                    mSTATUS = IDC_UPDATE_PREPAID_EXP_SLIP.GetCommandParamValue("O_STATUS").ToString();
                    if (IDC_UPDATE_PREPAID_EXP_SLIP.ExcuteError || mSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        mMESSAGE = iString.ISNull(IDC_UPDATE_PREPAID_EXP_SLIP.GetCommandParamValue("O_MESSAGE"));
                        if (mMESSAGE != string.Empty)
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();

                            MessageBoxAdv.Show(mMESSAGE, "1.Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            } 
            
            //선택된 처리대상에 대해 전표 전송//
            IDC_CREATE_PREPAID_EXP_SLIP.SetCommandParamValue("P_SYSDATE", mSYSDATE);
            IDC_CREATE_PREPAID_EXP_SLIP.ExecuteNonQuery();
            mSTATUS = IDC_CREATE_PREPAID_EXP_SLIP.GetCommandParamValue("O_STATUS").ToString();
            if (IDC_CREATE_PREPAID_EXP_SLIP.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                mMESSAGE = iString.ISNull(IDC_CREATE_PREPAID_EXP_SLIP.GetCommandParamValue("O_MESSAGE"));
                if (mMESSAGE != string.Empty)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(mMESSAGE, "2.Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            } 
            isDataTransaction1.Commit();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            SearchDB();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10300"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_EXP_SPREAD_METHOD_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_EXP_SPREAD_METHOD_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_EXP_SPREAD_METHOD_CODE.Focus();
                return;
            }
            
            //전표삭제여부 묻기.
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = string.Empty;
            string mMESSAGE = string.Empty;

            int mIDX_SELECT_FLAG = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("SELECT_FLAG");
            int mIDX_SLIP_HEADER_ID = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("SLIP_HEADER_ID"); 

            //선택된 처리대상에 대해 전표 전송 취소//
            for (int nRow = 0; nRow < IGR_PREPAID_EXP_SLIP.RowCount; nRow++)
            {
                if (iString.ISNull(IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_SELECT_FLAG)) == "Y" && 
                    iString.ISNull(IGR_PREPAID_EXP_SLIP.GetCellValue("SLIP_HEADER_ID")) != string.Empty)
                { 
                    // 대상 UPDATE //
                    IDC_CANCEL_PREPAID_EXP_SLIP.SetCommandParamValue("P_SELECT_FLAG", IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_SELECT_FLAG));
                    IDC_CANCEL_PREPAID_EXP_SLIP.SetCommandParamValue("P_SLIP_HEADER_ID", IGR_PREPAID_EXP_SLIP.GetCellValue(nRow, mIDX_SLIP_HEADER_ID));
                    IDC_CANCEL_PREPAID_EXP_SLIP.ExecuteNonQuery();
                    mSTATUS = IDC_CANCEL_PREPAID_EXP_SLIP.GetCommandParamValue("O_STATUS").ToString();
                    if (IDC_CANCEL_PREPAID_EXP_SLIP.ExcuteError || mSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        mMESSAGE = iString.ISNull(IDC_CANCEL_PREPAID_EXP_SLIP.GetCommandParamValue("O_MESSAGE"));
                        if (mMESSAGE != string.Empty)
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();

                            MessageBoxAdv.Show(mMESSAGE, "1.Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    } 
                }
            } 

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            SearchDB();
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
                    SearchDB();
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

        private void FCMF0272_Load(object sender, EventArgs e)
        {
             
        }

        private void FCMF0272_Shown(object sender, EventArgs e)
        {
            GetAccountBook();

            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);
            V_GL_DATE.EditValue = iDate.ISMonth_Last(iDate.ISGetDate(string.Format("{0}-01", W_PERIOD_NAME.EditValue)));

            W_SLIP_FLAG_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_SLIP_IF_FLAG.EditValue = W_SLIP_FLAG_N.RadioCheckedString;
        }

        private void IGR_PREPAID_EXP_SLIP_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_PREPAID_EXP_SLIP.GetColumnToIndex("SELECT_FLAG"))
            {
                IGR_PREPAID_EXP_SLIP.LastConfirmChanges();
                IDA_PREPAID_EXP_SLIP.OraSelectData.AcceptChanges();
                IDA_PREPAID_EXP_SLIP.Refillable = true;
            }
        }

        private void IGR_PREPAID_EXP_SLIP_CellDoubleClick(object pSender)
        {
            if (IGR_PREPAID_EXP_SLIP.RowIndex < 0)
            {
                return;
            }
            int mIDX_SLIP_HEADER_ID = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("SLIP_HEADER_ID");
            if (IGR_PREPAID_EXP_SLIP.ColIndex == mIDX_SLIP_HEADER_ID)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_PREPAID_EXP_SLIP.GetCellValue("SLIP_HEADER_ID"));
                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        private void V_CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            string vCHECK_FLAG = V_CHECK_YN.CheckBoxString;

            int vIDX_SELECT_FLAG = IGR_PREPAID_EXP_SLIP.GetColumnToIndex("SELECT_FLAG");
            for (int i = 0; i < IGR_PREPAID_EXP_SLIP.RowCount; i++)
            {
                IGR_PREPAID_EXP_SLIP.SetCellValue(i, vIDX_SELECT_FLAG, vCHECK_FLAG);
            }

            IGR_PREPAID_EXP_SLIP.LastConfirmChanges();
            IDA_PREPAID_EXP_SLIP.OraSelectData.AcceptChanges();
            IDA_PREPAID_EXP_SLIP.Refillable = true;
        }

        private void W_SLIP_FLAG_N_Click(object sender, EventArgs e)
        {
            if (W_SLIP_FLAG_N.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SLIP_IF_FLAG.EditValue = W_SLIP_FLAG_N.RadioCheckedString;
                BTN_SET_SLIP.Enabled = true;
                BTN_CANCEL_SLIP.Enabled = false;
            }
        }

        private void W_SLIP_FLAG_Y_Click(object sender, EventArgs e)
        {
            if (W_SLIP_FLAG_Y.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SLIP_IF_FLAG.EditValue = W_SLIP_FLAG_Y.RadioCheckedString;
                BTN_SET_SLIP.Enabled = false;
                BTN_CANCEL_SLIP.Enabled = true;
            }
        }

        private void W_SLIP_FLAG_A_Click(object sender, EventArgs e)
        {
            if (W_SLIP_FLAG_A.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SLIP_IF_FLAG.EditValue = W_SLIP_FLAG_A.RadioCheckedString;
                BTN_SET_SLIP.Enabled = false;
                BTN_CANCEL_SLIP.Enabled = false;
            }
        }
        
        #endregion 

        #region ----- Lookup Event -----

        private void ILA_CUSTOMER_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_PREPAID_EXP_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_TYPE", "Y");
        }
         
        private void ILA_EXP_SPREAD_METHOD_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("EXP_SPREAD_METHOD", "Y");
        }
         
        private void ILA_PERIOD_NAME_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {            
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddMonths(3));
        }

        private void ILA_PERIOD_NAME_SelectedRowData(object pSender)
        {
            V_GL_DATE.EditValue = iDate.ISMonth_Last(iDate.ISGetDate(string.Format("{0}-01", W_PERIOD_NAME.EditValue)));
        }

        #endregion

        #region ----- Adapter Lookup Event -----

        #endregion



    }
}