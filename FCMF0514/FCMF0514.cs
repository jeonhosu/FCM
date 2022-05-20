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
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace FCMF0514
{
    public partial class FCMF0514 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        //int mAccount_Book_ID;
        string mCurrency_Code = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public FCMF0514()
        {
            InitializeComponent();
        }

        public FCMF0514(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void DefaultValues()
        {
            IDC_BASE_CURRENCY.ExecuteNonQuery();
            mCurrency_Code = iConv.ISNull(IDC_BASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE"));
        }

        private void SearchDB()
        {
            IDA_DEPOSIT_MASTER.Fill();
            IGR_BANK_ACCOUNT_LIST.Focus();
        }

        private void Insert_Deposit_Master()
        {
            ISSUE_DATE.EditValue = iDate.ISGetDate(DateTime.Today);
            DUE_DATE.EditValue = ISSUE_DATE.EditValue;
            CURRENCY_CODE.EditValue = mCurrency_Code;
            BANK_CODE.Focus(); 
        }

        private void Insert_Deposit_Plan()
        {
            IGR_DEPOSIT_PLAN.SetCellValue("CURRENCY_CODE", CURRENCY_CODE.EditValue);
            IGR_DEPOSIT_PLAN.Focus();
        } 

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter2(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON2.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON2.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
                GridExcelConverterControl converter = new GridExcelConverterControl();
                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2007;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2007;
                IWorksheet sheet = workBook.Worksheets[0];
                //used to convert grid to excel 
                converter.GridToExcel(pGrid.BaseGrid, sheet, ConverterOptions.ColumnHeaders);
                //used to save the file
                workBook.SaveAs(saveFileDialog.FileName);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                        "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

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
                    if (IDA_DEPOSIT_MASTER.IsFocused)
                    {
                        IDA_DEPOSIT_MASTER.AddOver();
                        Insert_Deposit_Master();
                    }
                    else if(IDA_DEPOSIT_PLAN.IsFocused)
                    {
                        IDA_DEPOSIT_PLAN.AddOver();
                        Insert_Deposit_Plan();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_DEPOSIT_MASTER.IsFocused)
                    {                        
                        IDA_DEPOSIT_MASTER.AddUnder();
                        Insert_Deposit_Master();
                    }
                    else if (IDA_DEPOSIT_PLAN.IsFocused)
                    {
                        IDA_DEPOSIT_PLAN.AddUnder();
                        Insert_Deposit_Plan();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_DEPOSIT_MASTER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DEPOSIT_MASTER.IsFocused)
                    {
                        IDA_DEPOSIT_PLAN.Cancel();
                        IDA_DEPOSIT_MASTER.Cancel();
                    }
                    else if (IDA_DEPOSIT_PLAN.IsFocused)
                    {
                        IDA_DEPOSIT_PLAN.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DEPOSIT_MASTER.IsFocused)
                    {
                        IDA_DEPOSIT_MASTER.Delete();
                    }
                    else if (IDA_DEPOSIT_PLAN.IsFocused)
                    {
                        IDA_DEPOSIT_PLAN.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (IDA_DEPOSIT_MASTER.IsFocused)
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (IDA_DEPOSIT_MASTER.IsFocused)
                    {
                    }
                }
            }
        }

        #endregion;

        #region ----- Excel Upload -----

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            if (Excel_Upload() == true)
            {
                IDA_DEPOSIT_PLAN.Fill();
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
                    vXL_Load_OK = vXL_Upload.LoadXL(IDC_DEPOSIT_PLAN_UPLOAD, 2);
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

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = String.Empty;
            if (vResult == true)
            {
                //본테이블 이관전 체크//
                IDC_TRANS_DEPOSIT_PLAN_CHECK.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_TRANS_DEPOSIT_PLAN_CHECK.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_TRANS_DEPOSIT_PLAN_CHECK.GetCommandParamValue("O_MESSAGE"));
                if(vSTATUS == "DELETE")
                {
                    if(MessageBoxAdv.Show(vMESSAGE, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        vResult = false;
                    }
                    else
                    {
                        vResult = true;
                    }
                }
                else if(vSTATUS == "F")
                {
                    vResult = false;
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } 
                } 
            }

            vXL_Upload.DisposeXL();

            if(vResult == true)
            {
                IDC_TRANS_DEPOSIT_PLAN.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_TRANS_DEPOSIT_PLAN.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_TRANS_DEPOSIT_PLAN.GetCommandParamValue("O_MESSAGE"));
                if (vSTATUS == "F")
                {
                    vResult = false;
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
             
            return vResult;
        }

        #endregion

        #region ----- Form Event -----

        private void FCMF0514_Load(object sender, EventArgs e)
        {
            DefaultValues();
            IDA_DEPOSIT_MASTER.FillSchema();
        }

        //private void diposit_curr_amount()
        //{
        //    decimal vRate = 0;
        //    decimal vDeposit_Amount = 0;

        //    vRate = Convert.ToDecimal(INTER_1ST_DATE.EditValue.ToString());
        //    vDeposit_Amount = Convert.ToDecimal(DEPOSIT_AMOUNT.EditValue.ToString());

        //    if (vRate > 0)
        //    {
        //        DEPOSIT_CURR_AMOUNT.EditValue = vRate * vDeposit_Amount;
        //    }
        //    else
        //    {
        //        DEPOSIT_CURR_AMOUNT.EditValue = vDeposit_Amount;
        //    }
        //}

        private void Cancel_Amount()
        {
            decimal vCancel_Rate = 0;
            decimal vCancel_Amount = 0;
            decimal vCancel_Prin_Amount = 0;
            decimal vCancel_Inter_Amount = 0;
            decimal vFinal_Amount = 0;

            vCancel_Rate = iConv.ISDecimaltoZero(CANCEL_EXCHANGE_RATE.EditValue);
            vCancel_Amount = iConv.ISDecimaltoZero(CANCEL_AMOUNT.EditValue);
            vCancel_Prin_Amount = iConv.ISDecimaltoZero(CANCEL_PRIN_AMOUNT.EditValue);
            vCancel_Inter_Amount = iConv.ISDecimaltoZero(CANCEL_INTER_AMOUNT.EditValue);
            vFinal_Amount = iConv.ISDecimaltoZero(FINAL_AMOUNT.EditValue);

            if (vCancel_Rate > 0)
            {
                CANCEL_CURR_AMOUNT.EditValue = vCancel_Rate * vCancel_Amount; 
            }
            else
            {
                CANCEL_CURR_AMOUNT.EditValue = 0;
            }
            FINAL_AMOUNT.EditValue = vFinal_Amount + iConv.ISDecimaltoZero(CANCEL_CURR_AMOUNT.EditValue, 0);
        }


        #endregion

        #region ----- Lookup Event -----

        private void ILA_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_ACCOUNT_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_SITE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK_SITE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CURRENCY_CODE.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ILD_CURRENCY_CODE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_CODE_PL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CURRENCY_CODE.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ILD_CURRENCY_CODE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_COST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COST_CENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_DEPOSIT_INTER_KIND_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter2("DEPOSIT_INTER_KIND", "Y");
        }

        private void ILA_DEPOSIT_INTER_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter2("DEPOSIT_INTER_TYPE", "Y");
        }

        private void ILA_DEPOSIT_INTER_CAL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter2("DEPOSIT_INTER_CAL", "Y");
        }

        private void ILA_PAYMENT_PERIOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter2("PERIOD_CYCLE_TYPE", "Y");
        }

        private void ILA_INTEREST_PERIOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter2("PERIOD_CYCLE_TYPE", "Y");
        }

        private void ilaSTATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter2("TRANS_STATUS", "Y");
        }
         
        private void ilaACCOUNT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_TYPE", "Y");
        }

        private void ILA_DEPOSIT_PLAN_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DEPOSIT_PLAN_TYPE", "Y"); 
        }

        private void ILA_DEPOSIT_PLAN_TYPE_PL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DEPOSIT_PLAN_TYPE", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_DEPOSIT_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["BANK_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BANK_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                BANK_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["BANK_SITE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BANK_SITE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                BANK_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["ACCOUNT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                ACCOUNT_TYPE_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["BANK_ACCOUNT_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BANK_ACCOUNT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                BANK_ACCOUNT_CODE.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["BANK_ACCOUNT_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BANK_ACCOUNT_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                BANK_ACCOUNT_NUM.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["BANK_ACCOUNT_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BANK_ACCOUNT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                BANK_ACCOUNT_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["ISSUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ISSUE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                ISSUE_DATE.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["DUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(DUE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                DUE_DATE.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CURRENCY_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                CURRENCY_CODE.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["TRANS_STATUS"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(TRANS_STATUS_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                TRANS_STATUS.Focus();
                return;
            }
        }

        private void IDA_DEPOSIT_PLAN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            int vIDX_COL = 0;
            if (iConv.ISNull(e.Row["PLAN_DATE"]) == string.Empty)
            {
                vIDX_COL = IGR_DEPOSIT_PLAN.GetColumnToIndex("PLAN_DATE");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEPOSIT_PLAN, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; 
                return;
            }
            if (iConv.ISNull(e.Row["TRX_DATE"]) == string.Empty)
            {
                vIDX_COL = IGR_DEPOSIT_PLAN.GetColumnToIndex("TRX_DATE");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEPOSIT_PLAN, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; 
                return;
            }
            if (iConv.ISNull(e.Row["DEPOSIT_PLAN_TYPE"]) == string.Empty)
            {
                vIDX_COL = IGR_DEPOSIT_PLAN.GetColumnToIndex("DEPOSIT_PLAN_NAME");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEPOSIT_PLAN, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                vIDX_COL = IGR_DEPOSIT_PLAN.GetColumnToIndex("CURRENCY_CODE");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEPOSIT_PLAN, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                vIDX_COL = IGR_DEPOSIT_PLAN.GetColumnToIndex("AMOUNT");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEPOSIT_PLAN, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}