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
using System.IO;

namespace FCMF0580
{
    public partial class FCMF0580 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        int mAccount_Book_ID;
        string mCurrency_Code;

        #endregion;

        #region ----- Constructor -----

        public FCMF0580()
        {
            InitializeComponent();
        }

        public FCMF0580(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            IGR_LOAN_LIST.LastConfirmChanges();
            IDA_LOAN_MASTER.OraSelectData.AcceptChanges();
            IDA_LOAN_MASTER.Refillable = true;

            IGR_LOAN_PLAN.LastConfirmChanges();
            IDA_LOAN_PLAN.OraSelectData.AcceptChanges();
            IDA_LOAN_PLAN.Refillable = true;

            IDA_LOAN_MASTER.Fill();
            IGR_LOAN_LIST.Focus();
        }

        private void Insert_Loan_Master()
        {            
            LOAN_NUM.Focus();
        }

        private void Insert_Loan_Plan()
        {
            IGR_LOAN_PLAN.SetCellValue("LOAN_NUM", LOAN_NUM.EditValue);
          
            IGR_LOAN_PLAN.Focus();
        }

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Init_Item(object pLoan_Kind)
        {
            if (iConv.ISNull(pLoan_Kind) == "2".ToString() || iConv.ISNull(pLoan_Kind) == "3".ToString())
            {
            

                ISSUE_DATE.ReadOnly = true;
                ISSUE_DATE.Insertable = false;
                ISSUE_DATE.Updatable = false;
                ISSUE_DATE.TabStop = false;

                DUE_DATE.ReadOnly = true;
                DUE_DATE.Insertable = false;
                DUE_DATE.Updatable = false;
                DUE_DATE.TabStop = false;

                CURRENCY_CODE.ReadOnly = true;
                CURRENCY_CODE.Insertable = false;
                CURRENCY_CODE.Updatable = false;
                CURRENCY_CODE.TabStop = false;

                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;
                EXCHANGE_RATE.TabStop = false;

                LOAN_CURR_AMOUNT.ReadOnly = true;
                LOAN_CURR_AMOUNT.Insertable = false;
                LOAN_CURR_AMOUNT.Updatable = false;
                LOAN_CURR_AMOUNT.TabStop = false;

                LOAN_AMOUNT.ReadOnly = true;
                LOAN_AMOUNT.Insertable = false;
                LOAN_AMOUNT.Updatable = false;
                LOAN_AMOUNT.TabStop = false;

                V_EXEC_NUM.ReadOnly = true;
                V_EXEC_NUM.Insertable = false;
                V_EXEC_NUM.Updatable = false;
                V_EXEC_NUM.TabStop = false;

                V_EXEC_CODE.ReadOnly = true;
            }
            else
            {
           
                ISSUE_DATE.ReadOnly = false;
                ISSUE_DATE.Insertable = true;
                ISSUE_DATE.Updatable = true;
                ISSUE_DATE.TabStop = true;

                DUE_DATE.ReadOnly = false;
                DUE_DATE.Insertable = true;
                DUE_DATE.Updatable = true;
                DUE_DATE.TabStop = true;

                CURRENCY_CODE.ReadOnly = false;
                CURRENCY_CODE.Insertable = true;
                CURRENCY_CODE.Updatable = true;
                CURRENCY_CODE.TabStop = true;

                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;
                EXCHANGE_RATE.TabStop = true;

                LOAN_CURR_AMOUNT.ReadOnly = false;
                LOAN_CURR_AMOUNT.Insertable = true;
                LOAN_CURR_AMOUNT.Updatable = true;
                LOAN_CURR_AMOUNT.TabStop = true;

                LOAN_AMOUNT.ReadOnly = false;
                LOAN_AMOUNT.Insertable = true;
                LOAN_AMOUNT.Updatable = true;
                LOAN_AMOUNT.TabStop = true;

                V_EXEC_NUM.ReadOnly = false;
                V_EXEC_NUM.Insertable = true;
                V_EXEC_NUM.Updatable = true;
                V_EXEC_NUM.TabStop = true;

                V_EXEC_CODE.ReadOnly = true;

                Init_Currency_Amount();
            }
        }


        

        private void Init_Currency_Amount()
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) == string.Empty || CURRENCY_CODE.EditValue.ToString() == mCurrency_Code)
            {
                EXCHANGE_RATE.EditValue = DBNull.Value;

                if (iConv.ISDecimaltoZero(LOAN_CURR_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    LOAN_CURR_AMOUNT.EditValue = 0;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                LOAN_CURR_AMOUNT.ReadOnly = true;
                LOAN_CURR_AMOUNT.Insertable = false;
                LOAN_CURR_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                LOAN_CURR_AMOUNT.TabStop = false;
            }
            else
            {
                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                LOAN_CURR_AMOUNT.ReadOnly = false;
                LOAN_CURR_AMOUNT.Insertable = true;
                LOAN_CURR_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                LOAN_CURR_AMOUNT.TabStop = true;

                V_EXEC_NUM.ReadOnly = false;
            }
        }

        private void Init_Loan_Amount()
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) != mCurrency_Code)
            {
                LOAN_AMOUNT.EditValue = iConv.ISDecimaltoZero(LOAN_CURR_AMOUNT.EditValue) * iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue);
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, string pCol_Name)
        {
            int vIDX_COL = pGrid.GetColumnToIndex(pCol_Name);
            int mCol_Count = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[vIDX_COL].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
        }

        private object Get_Grid_Prompt_0(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[0].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[0].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[0].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[0].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[0].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[0].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion

        #region ----- Excel Upload -----

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //기존 데이터 존재 확인//
            IDC_GET_LOAN_PLAN_COUNT.SetCommandParamValue("W_LOAN_NUM", LOAN_NUM.EditValue);
            IDC_GET_LOAN_PLAN_COUNT.SetCommandParamValue("W_LOAN_PLAN_TYPE", W_LOAN_PLAN_TYPE.EditValue);
            IDC_GET_LOAN_PLAN_COUNT.SetCommandParamValue("W_EXEC_NUM", W_EXEC_NUM.EditValue);
            IDC_GET_LOAN_PLAN_COUNT.ExecuteNonQuery();
            decimal vLOAN_PLAN_COUNT = iConv.ISDecimaltoZero(IDC_GET_LOAN_PLAN_COUNT.GetCommandParamValue("O_LOAN_PLAN_COUNT"));
            if(vLOAN_PLAN_COUNT > 0)
            {
                if(MessageBoxAdv.Show(string.Format("Loan Num [{0}] :: {1} \r{2}", LOAN_NUM.EditValue, isMessageAdapter1.ReturnText("FCM_10082"), isMessageAdapter1.ReturnText("FCM_10323")), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                    return;
                }

                //delete
                IDC_DELETE_LOAN_PLAN_ALL.SetCommandParamValue("P_LOAN_NUM", LOAN_NUM.EditValue);
                IDC_DELETE_LOAN_PLAN_ALL.SetCommandParamValue("W_EXEC_NUM", W_EXEC_NUM.EditValue);
                IDC_DELETE_LOAN_PLAN_ALL.ExecuteNonQuery();
                string vSTATUS = iConv.ISNull(IDC_DELETE_LOAN_PLAN_ALL.GetCommandParamValue("O_STATUS"));
                string vMESSAGE = iConv.ISNull(IDC_DELETE_LOAN_PLAN_ALL.GetCommandParamValue("O_MESSAGE"));
                if(vSTATUS == "F")
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                    return;
                }

            }

            if (Excel_Upload() == true)
            {
                IDA_LOAN_PLAN.Fill();
            }
            UPLOAD_FILE_PATH.EditValue = string.Empty;
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void Select_Excel_File()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    UPLOAD_FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private bool Excel_Upload()
        {
            bool vResult = false;

            if (iConv.ISNull(UPLOAD_FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPLOAD_FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            bool vXL_Load_OK = false;
            string vOPenFileName = UPLOAD_FILE_PATH.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);
            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return vResult;
            }

            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL(IDC_INSERT_LOAN_PLAN_UPLOAD, 2);
                    if (vXL_Load_OK == false)
                    { 
                        vResult = false;
                    }
                    else
                    { 
                        vResult = true;
                    }
                }
            }
            catch (Exception ex)
            { 
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                vResult = false;
                return vResult;
            }
            vXL_Upload.DisposeXL();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            return vResult;
        }

        #endregion


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
                    if (IDA_LOAN_MASTER.IsFocused)
                    {
                        IDA_LOAN_MASTER.AddOver();
                        Insert_Loan_Master();
                    }
                    else if(IDA_LOAN_PLAN.IsFocused)
                    {
                        IDA_LOAN_PLAN.AddOver();
                        Insert_Loan_Plan();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_LOAN_MASTER.IsFocused)
                    {
                        IDA_LOAN_MASTER.AddUnder();
                        Insert_Loan_Master();
                    }
                    else if (IDA_LOAN_PLAN.IsFocused)
                    {
                        IDA_LOAN_PLAN.AddUnder();
                        Insert_Loan_Plan();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_LOAN_MASTER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_LOAN_MASTER.IsFocused)
                    {
                        IDA_LOAN_MASTER.Cancel();
                    }
                    else if (IDA_LOAN_PLAN.IsFocused)
                    {
                        IDA_LOAN_PLAN.Cancel(); 
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_LOAN_MASTER.IsFocused)
                    {
                        IDA_LOAN_MASTER.Delete();
                    }
                    else if (IDA_LOAN_PLAN.IsFocused)
                    {
                        IDA_LOAN_PLAN.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (IDA_LOAN_MASTER.IsFocused)
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (IDA_LOAN_MASTER.IsFocused)
                    {
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        
        private void FCMF0580_Load(object sender, EventArgs e)
        {
            
        }
        
        private void FCMF0580_Shown(object sender, EventArgs e)
        {
            IDC_DV_ACCOUNT_BOOK.ExecuteNonQuery();
            mAccount_Book_ID = iConv.ISNumtoZero(IDC_DV_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID"));
            mCurrency_Code = iConv.ISNull(IDC_DV_ACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));

            IDA_LOAN_MASTER.FillSchema();
            IDA_LOAN_PLAN.FillSchema();
        }

        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Loan_Amount();
        }

        private void LOAN_CURR_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Loan_Amount();
        }
        
        
        private void IGR_LOAN_PLAN_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vIDX_PLAN_DATE = IGR_LOAN_PLAN.GetColumnToIndex("PLAN_DATE");
            if(e.ColIndex == vIDX_PLAN_DATE)
            {
                IGR_LOAN_PLAN.SetCellValue("PLAN_DATE_FR", iDate.ISMonth_1st(e.NewValue));
                IGR_LOAN_PLAN.SetCellValue("PLAN_DATE_TO", iDate.ISMonth_Last(e.NewValue));
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaLOAN_KIND_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("LOAN_KIND", "Y");
        }

        private void ilaLOAN_KIND_SelectedRowData(object pSender)
        {
            Init_Item(LOAN_KIND.EditValue);
            if (iConv.ISNull(LOAN_KIND.EditValue) == "2".ToString() || iConv.ISNull(LOAN_KIND.EditValue) == "3".ToString())
            {

            }
            else
            {
                if (iConv.ISNull(ISSUE_DATE.EditValue) == string.Empty)
                {
                    ISSUE_DATE.EditValue = DateTime.Today;
                }
                if (iConv.ISNull(DUE_DATE.EditValue) == string.Empty)
                {
                    DUE_DATE.EditValue = DateTime.Today;
                }
                if (iConv.ISNull(CURRENCY_CODE.EditValue) == string.Empty)
                {
                    CURRENCY_CODE.EditValue = mCurrency_Code;
                }         
            }
        }

        private void ilaLOAN_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("LOAN_TYPE", "Y");
        }

        private void ilaLOAN_USE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("LOAN_USE", "Y");
        }

        private void ilaLOAN_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaL_CURRENCY_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_LOAN_NUM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        
        private void ilaCURRENCY_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_LOAN_NUM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_CODE_SelectedRowData(object pSender)
        {
            Init_Currency_Amount();
        }

        private void ilaENSURE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ENSURE_TYPE", "Y");
        }

        private void ilaREPAY_CONDITION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("REPAY_CONDITION", "Y");
        }

        private void ilaRCV_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaRCV_ACCOUNT_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaINTEREST_ACCOUNT_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaICOMMISSION_ACCOUNT_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaINTEREST_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("INTEREST_TYPE", "Y");
        }
        
        private void ilaINTEREST_PAYMENT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("INTEREST_PAYMENT_TYPE", "Y");
        }

        private void ilaLOAN_BANK_ACCOUNT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCOST_CENTER_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COST_CENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCOST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ILD_COMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        }

        private void ILA_LOAN_PLAN_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("LOAN_PLAN_TYPE", "Y");
        }

        private void ILA_LOAN_PLAN_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("LOAN_PLAN_TYPE", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaLONE_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["LOAN_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10192"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                LOAN_NUM.Focus();
                return;
            }
            //if (iConv.ISNull(e.Row["CHANGE_DATE"]) == string.Empty)
            //{

            //    e.Cancel = true;
            //    return;
            //}

            if (iConv.ISNull(e.Row["EXEC_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SKFCM_10639"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                V_EXEC_NUM.Focus();
                return;
            }

            if (iConv.ISNull(e.Row["LOAN_KIND"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10193"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                LOAN_KIND_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["LOAN_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10194"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                LOAN_TYPE_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["LOAN_USE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10195"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                LOAN_USE_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["LOAN_KIND"]) == "2".ToString() || iConv.ISNull(e.Row["LOAN_KIND"]) == "3".ToString())
            {
                //if (iConv.ISNull(e.Row["L_ISSUE_DATE"]) == string.Empty)
                //{
                //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10196"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    e.Cancel = true;
                //    return;
                //}
                //if (iConv.ISNull(e.Row["L_DUE_DATE"]) == string.Empty)
                //{
                //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10145"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    e.Cancel = true;
                //    return;
                //}
                //if (iConv.ISNull(e.Row["L_CURRENCY_CODE"]) == string.Empty)
                //{
                //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    e.Cancel = true;
                //    return;
                //}
                //if (mCurrency_Code != iConv.ISNull(e.Row["L_CURRENCY_CODE"]))
                //{
                //    if (iConv.ISNull(e.Row["L_EXCHANGE_RATE"]) == string.Empty)
                //    {
                //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10125"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //        e.Cancel = true;
                //        return;
                //    }
                //    if (iConv.ISNull(e.Row["LIMIT_CURR_AMOUNT"]) == string.Empty)
                //    {
                //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //        e.Cancel = true;
                //        return;
                //    }
                //}
                //if (iConv.ISNull(e.Row["LIMIT_AMOUNT"]) == string.Empty)
                //{
                //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10197"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    e.Cancel = true;
                //    return;
                //}
            }
            else
            {
                if (iConv.ISNull(e.Row["ISSUE_DATE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10196"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    ISSUE_DATE.Focus();
                    return;
                }
                if (iConv.ISNull(e.Row["DUE_DATE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10145"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    DUE_DATE.Focus();
                    return;
                }
                if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    CURRENCY_CODE.Focus();
                    return;
                }
                if (mCurrency_Code != iConv.ISNull(e.Row["CURRENCY_CODE"]))
                {
                    if (iConv.ISNull(e.Row["EXCHANGE_RATE"]) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10125"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        EXCHANGE_RATE.Focus();
                        return;
                    }
                    if (iConv.ISNull(e.Row["LOAN_CURR_AMOUNT"]) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        LOAN_CURR_AMOUNT.Focus();
                        return;
                    }

                }
                if (iConv.ISNull(e.Row["LOAN_AMOUNT"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10197"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    LOAN_AMOUNT.Focus();
                    return;
                }
            }
            
            //if (iString.ISNull(e.Row["LOAN_ACCOUNT_CONTROL_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    LOAN_ACCOUNT_CONTROL_NAME.Focus();
            //    return;
            //}
            if (iConv.ISNull(e.Row["LOAN_BANK_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10200"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                LOAN_BANK_NAME.Focus();
                return;
            }
        }

        private void idaLONE_MASTER_PreDelete(ISPreDeleteEventArgs e)
        {
            
        }

        private void IDA_LOAN_PLAN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["LOAN_NUM"]) == string.Empty)
            { 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "LOAN_NUM"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PLAN_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "PLAN_DATE"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["LOAN_PLAN_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "LOAN_PLAN_TYPE"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "CURRENCY_CODE"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "AMOUNT"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PLAN_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "PLAN_DATE_FR"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PLAN_DATE_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_LOAN_PLAN, "PLAN_DATE_TO"))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaLONE_MASTER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Item(LOAN_KIND.EditValue);
        }


        #endregion

        private void ILA_LOAN_NUM_SelectedRowData(object pSender)
        {
            IDC_LOAN_MASTER.ExecuteNonQuery();
        }
    }
}

