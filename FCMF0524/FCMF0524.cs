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
using Syncfusion.XlsIO;

namespace FCMF0524
{
    public partial class FCMF0524 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0524()
        {
            InitializeComponent();
        }

        public FCMF0524(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (iString.ISNull(PAYMENT_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAYMENT_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAYMENT_DATE_0.Focus();
                return;
            }

            IGR_PAYMENT_SELECTED.LastConfirmChanges();
            IDA_PAYMENT_SELECTED.OraSelectData.AcceptChanges();
            IDA_PAYMENT_SELECTED.Refillable = true;

            string mAccount_Code = iString.ISNull(IGR_PAYMENT_ACCOUNT.GetCellValue("ACCOUNT_CODE"));
            int mIDX_Account_Code = IGR_PAYMENT_ACCOUNT.GetColumnToIndex("ACCOUNT_CODE");
            IDA_BATCH_PAYMENT_ACCOUNT.Fill();
            IGR_PAYMENT_ACCOUNT.Focus();
        }

        private void SEARCH_DETAIL(Object pACCOUNT_CONTROL_ID)
        {
            IGR_PAYMENT_SELECTED.LastConfirmChanges();
            IDA_PAYMENT_SELECTED.OraSelectData.AcceptChanges();
            IDA_PAYMENT_SELECTED.Refillable = true;

            IDA_ITEM_PROMPT.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            INIT_MANAGEMENT_COLUMN();

            // 대량지급 내역 조회.
            IDA_PAYMENT_SELECTED.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_PAYMENT_SELECTED.Fill();
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            IDA_ITEM_PROMPT.Fill();
            if (IDA_ITEM_PROMPT.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            int mStart_Column = 6;
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = IDA_ITEM_PROMPT.CurrentRow[mENABLED_COLUMN];
                mCOLUMN_DESC = IDA_ITEM_PROMPT.CurrentRow[mIDX_Column];
                if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    IGR_PAYMENT_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_PAYMENT_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                    IGR_PAYMENT_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    IGR_PAYMENT_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }

            // 전표일자 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_DATE");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 적요.
            mIDX_Column = 0;
            mIDX_Column = IGR_PAYMENT_SELECTED.GetColumnToIndex("SLIP_REMARK");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 외화금액 - 통화관리 하는 경우 적용.
            mIDX_Column = 0;
            mIDX_Column = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_CURR_AMOUNT");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["CONTROL_CURRENCY_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Insertable = 1;
                IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_Column].Updatable = 1;
            }
            IGR_PAYMENT_SELECTED.ResetDraw = true;
        }
         

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            int vIDX_SUMMARY_FLAG = pGrid.GetColumnToIndex("SUMMARY_FLAG");

            for (int i = 0; i < pGrid.RowCount; i++)
            {
                if (iString.ISNull(IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_SUMMARY_FLAG)) == "N")
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG); 
                }
                else
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, "N");
                }
            }
            pGrid.LastConfirmChanges(); 
            IDA_PAYMENT_SELECTED.OraSelectData.AcceptChanges();
            IDA_PAYMENT_SELECTED.Refillable = true; 
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

        #region ----- Excel Export -----

        private void ExcelExport()
        {
            //XLPrinting("FILE");
            Xls_Export(); 
        }

        private void Xls_Export()
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Payment List";

            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.Filter = "Excel file(*.xlsx)|*.xlsx|Excel file(*.xlsx)|*.xlsx";
            saveFileDialog1.DefaultExt = "xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = saveFileDialog1.FileName;
                System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                try
                {
                    if (vFileName.Exists)
                    {
                        vFileName.Delete();
                    }
                }
                catch (Exception EX)
                {
                    MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            vMessageText = string.Format(" Writing Starting...");

            System.Windows.Forms.Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine ExcelEngine = new ExcelEngine();

            //Step 2 : Instantiate the excel application object.
            IApplication Exc_App = ExcelEngine.Excel;

            //set 2.1 : file Extension check =>xlsx, xls 
            if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLSX")
            {
                ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2010;   // EXCEL .XLSX 로 바꾼뒤 2007 -> 2010 변경
            }
            else
            {
                ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
            }

            //A new workbook is created.[Equivalent to creating a new workbook in MS Excel]
            //The new workbook will have 3 worksheets 
            IWorkbook Exc_WorkBook = Exc_App.Workbooks.Create(1);
            if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLSX")
            {
                Exc_WorkBook.Version = ExcelVersion.Excel2010;    // EXCEL .XLSX 로 바꾼뒤 2007 -> 2010 변경
            }
            else
            {
                Exc_WorkBook.Version = ExcelVersion.Excel97to2003;
            }
            IWorksheet sheet = Exc_WorkBook.Worksheets[0];

            IDA_PAYMENT_ACCOUNT.Fill();
            foreach(System.Data.DataRow vROW in IDA_PAYMENT_ACCOUNT.CurrentRows)
            {
                try
                {
                    //DATA 조회 
                    IDA_PRINT_BATCH_PAYMENT1.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", vROW["ACCOUNT_CONTROL_ID"]);
                    IDA_PRINT_BATCH_PAYMENT1.Fill();
                    int vCountRow = IDA_PRINT_BATCH_PAYMENT1.CurrentRows.Count;
                    if (vCountRow > 0)
                    { 
                        //The first worksheet object in the worksheets collection is accessed.
                        sheet = Exc_WorkBook.Worksheets.Create(iString.ISNull(vROW["ACCOUNT_CODE"]));
                        sheet.Activate();                        

                        //헤더 프롬프트 조회// 
                        IDA_MANAGEMENT_PROMPT.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", vROW["ACCOUNT_CONTROL_ID"]);
                        IDA_MANAGEMENT_PROMPT.Fill();

                        string vPayment_DATE = string.Format("Payment Date : {0}", iDate.ISGetDate(PAYMENT_DATE_0.EditValue).ToShortDateString());

                        //Export DataTable.
                        sheet.ImportDataTable(IDA_PRINT_BATCH_PAYMENT1.OraDataTable(), false, 1, 1, IDA_PRINT_BATCH_PAYMENT1.CurrentRows.Count, IDA_PRINT_BATCH_PAYMENT1.OraSelectData.Columns.Count);

                        //1.title insert
                        sheet.InsertRow(1);
                        sheet.MergeRanges(sheet.Range[1, 1], sheet.Range[1, 5]); 
                        sheet.Range[1, 1].Value = vPayment_DATE;
                        //2.prompt insert
                        sheet.InsertRow(2);
                        sheet.ImportDataTable(IDA_MANAGEMENT_PROMPT.OraDataTable(), false, 2, 1);
                        sheet.Range[1, IDA_PRINT_BATCH_PAYMENT1.CurrentRows.Count].AutofitColumns();
                    }

                }
                catch (System.Exception ex)
                {
                    vMessageText = ex.Message;
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            //Saving the workbook to disk.
            Exc_WorkBook.SaveAs(vSaveFileName);

            //Close the workbook.
            Exc_WorkBook.Close();

            //No exception will be thrown if there are unsaved workbooks.
            ExcelEngine.ThrowNotSavedOnDestroy = false;
            ExcelEngine.Dispose();

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start(vSaveFileName);
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            string vSaveFileName2 = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;
            object vPAYMENT_DATE = iDate.ISGetDate(PAYMENT_DATE_0.EditValue).ToShortDateString();

            if (iString.ISNull(vPAYMENT_DATE) == String.Empty)
            {//기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 데이터 조회.
            IDA_PRINT_BATCH_PAYMENT1.Fill();
            vCountRow = IDA_PRINT_BATCH_PAYMENT1.OraSelectData.Rows.Count;

            IDA_PRINT_BATCH_PAYMENT2.Fill();
            vCountRow = vCountRow + IDA_PRINT_BATCH_PAYMENT2.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("Payment_{0}", vPAYMENT_DATE);
                vSaveFileName2 = string.Format("Curr_{0}", vSaveFileName);

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xlsx)|*.xlsx";
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName2 = vFilePath.Replace(vSaveFileName, vSaveFileName2);
                    vSaveFileName = vFilePath;
                    
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    
                    vFileName = new System.IO.FileInfo(vSaveFileName2);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {   
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0524_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite1(vPAYMENT_DATE, IDA_PRINT_BATCH_PAYMENT1);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            //외화 인쇄//
            xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {                
                xlPrinting.OpenFileNameExcel = "FCMF0524_002.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(vPAYMENT_DATE, IDA_PRINT_BATCH_PAYMENT2);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName2);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }

                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

            
        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_PAYMENT_SELECTED.IsFocused)
                    {
                        IDA_PAYMENT_SELECTED.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_PAYMENT_SELECTED.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_PAYMENT_SELECTED.IsFocused)
                    {
                        if (iString.ISNull(IGR_PAYMENT_SELECTED.GetCellValue("SUMMARY_FLAG")) != "N")
                        {
                            return;
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    Xls_Export(); 
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void FCMF0524_Load(object sender, EventArgs e)
        {            
        }

        private void FCMF0524_Shown(object sender, EventArgs e)
        {
            PAYMENT_DATE_0.EditValue = DateTime.Today;
            RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            RB_VENDOR.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_SORT_TYPE.EditValue = RB_VENDOR.RadioCheckedString;

            IGB_STATUS.BringToFront();

            IDA_BATCH_PAYMENT_ACCOUNT.FillSchema();
            IDA_PAYMENT_SELECTED.FillSchema();
        }

        private void BTN_CREATE_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PAYMENT_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAYMENT_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAYMENT_DATE_0.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0524_SET vFCMF0524_SET = new FCMF0524_SET(isAppInterfaceAdv1.AppInterface, PAYMENT_DATE_0.EditValue);
            vRESULT = vFCMF0524_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SEARCH_DB();
            } 
            vFCMF0524_SET.Dispose();
        }

        private void BTN_DELETE_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PAYMENT_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAYMENT_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAYMENT_DATE_0.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IDA_PAYMENT_SELECTED.Update();

            int vIDX_CHECK_YN = IGR_PAYMENT_SELECTED.GetColumnToIndex("CHECK_YN");
            int vIDX_BALANCE_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("BALANCE_DATE");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int vIDX_CURRENCY_CODE = IGR_PAYMENT_SELECTED.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_ITEM_GROUP_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("ITEM_GROUP_ID");
            int vIDX_GL_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_DATE");
            int vIDX_BALANCE_STATEMENT_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("BALANCE_STATEMENT_ID");

            object vCHECK_YN = "N";
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            for (int i = 0; i < IGR_PAYMENT_SELECTED.RowCount; i++)
            {
                vCHECK_YN = IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_CHECK_YN);
                if (iString.ISNull(vCHECK_YN, "N") == "Y")
                {
                    IGR_PAYMENT_SELECTED.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    IGR_PAYMENT_SELECTED.CurrentCellActivate(i, vIDX_CHECK_YN);
 
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("P_CHECK_YN", vCHECK_YN);
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_BATCH_DATE", PAYMENT_DATE_0.EditValue);
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_BALANCE_DATE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_BALANCE_DATE));
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_CURRENCY_CODE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_CURRENCY_CODE));
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_ITEM_GROUP_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_ITEM_GROUP_ID));
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_GL_DATE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_GL_DATE));
                    IDC_DEL_PAYMENT_SELECT.SetCommandParamValue("W_BALANCE_STATEMENT_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_BALANCE_STATEMENT_ID));
                    IDC_DEL_PAYMENT_SELECT.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(IDC_DEL_PAYMENT_SELECT.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(IDC_DEL_PAYMENT_SELECT.GetCommandParamValue("O_MESSAGE"));
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if(vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            SEARCH_DETAIL(IGR_PAYMENT_ACCOUNT.GetCellValue("ACCOUNT_CONTROL_ID"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_CONFIRM_Y_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PAYMENT_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAYMENT_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAYMENT_DATE_0.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90014"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IDA_PAYMENT_SELECTED.Update();

            int vIDX_CHECK_YN = IGR_PAYMENT_SELECTED.GetColumnToIndex("CHECK_YN");
            int vIDX_BALANCE_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("BALANCE_DATE");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int vIDX_CURRENCY_CODE = IGR_PAYMENT_SELECTED.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_ITEM_GROUP_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("ITEM_GROUP_ID");
            int vIDX_GL_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_DATE");
            int vIDX_BALANCE_STATEMENT_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("BALANCE_STATEMENT_ID");

            object vCHECK_YN = "N";
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            for (int i = 0; i < IGR_PAYMENT_SELECTED.RowCount; i++)
            {
                vCHECK_YN = IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_CHECK_YN);
                if (iString.ISNull(vCHECK_YN, "N") == "Y")
                {
                    IGR_PAYMENT_SELECTED.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    IGR_PAYMENT_SELECTED.CurrentCellActivate(i, vIDX_CHECK_YN);

                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("P_CHECK_YN", vCHECK_YN);
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_CONFIRM_YN", "Y");
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_BATCH_DATE", PAYMENT_DATE_0.EditValue);
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_BALANCE_DATE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_BALANCE_DATE));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_CURRENCY_CODE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_CURRENCY_CODE));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_ITEM_GROUP_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_ITEM_GROUP_ID));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_GL_DATE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_GL_DATE));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_BALANCE_STATEMENT_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_BALANCE_STATEMENT_ID));
                    IDC_PAYMENT_CONFIRM.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(IDC_PAYMENT_CONFIRM.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(IDC_PAYMENT_CONFIRM.GetCommandParamValue("O_MESSAGE"));
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            SEARCH_DETAIL(IGR_PAYMENT_ACCOUNT.GetCellValue("ACCOUNT_CONTROL_ID"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_CONFIRM_N_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            if (iString.ISNull(PAYMENT_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAYMENT_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAYMENT_DATE_0.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90015"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IDA_PAYMENT_SELECTED.Update();

            int vIDX_CHECK_YN = IGR_PAYMENT_SELECTED.GetColumnToIndex("CHECK_YN");
            int vIDX_BALANCE_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("BALANCE_DATE");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int vIDX_CURRENCY_CODE = IGR_PAYMENT_SELECTED.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_ITEM_GROUP_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("ITEM_GROUP_ID");
            int vIDX_GL_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_DATE");
            int vIDX_BALANCE_STATEMENT_ID = IGR_PAYMENT_SELECTED.GetColumnToIndex("BALANCE_STATEMENT_ID");

            object vCHECK_YN = "N";
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            for (int i = 0; i < IGR_PAYMENT_SELECTED.RowCount; i++)
            {
                vCHECK_YN = IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_CHECK_YN);
                if (iString.ISNull(vCHECK_YN, "N") == "Y")
                {
                    IGR_PAYMENT_SELECTED.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    IGR_PAYMENT_SELECTED.CurrentCellActivate(i, vIDX_CHECK_YN);

                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("P_CHECK_YN", vCHECK_YN);
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_CONFIRM_YN", "N");
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_BATCH_DATE", PAYMENT_DATE_0.EditValue);
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_BALANCE_DATE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_BALANCE_DATE));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_CURRENCY_CODE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_CURRENCY_CODE));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_ITEM_GROUP_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_ITEM_GROUP_ID));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_GL_DATE", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_GL_DATE));
                    IDC_PAYMENT_CONFIRM.SetCommandParamValue("W_BALANCE_STATEMENT_ID", IGR_PAYMENT_SELECTED.GetCellValue(i, vIDX_BALANCE_STATEMENT_ID));
                    IDC_PAYMENT_CONFIRM.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(IDC_PAYMENT_CONFIRM.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(IDC_PAYMENT_CONFIRM.GetCommandParamValue("O_MESSAGE"));
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            SEARCH_DETAIL(IGR_PAYMENT_ACCOUNT.GetCellValue("ACCOUNT_CONTROL_ID"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }
        
        private void RB_ALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;
            if (vRadio.Checked == true)
            {
                CONFIRM_STATUS_0.EditValue = vRadio.RadioCheckedString;
            }
        }

        private void RB_VENDOR_Click(object sender, EventArgs e)
        {
            if (RB_VENDOR.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SORT_TYPE.EditValue = RB_VENDOR.RadioButtonString;
            }
        }

        private void RB_DUE_DATE_Click(object sender, EventArgs e)
        {
            if (RB_DUE_DATE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SORT_TYPE.EditValue = RB_DUE_DATE.RadioButtonString;
            }
        }

        private void RB_GL_DATE_Click(object sender, EventArgs e)
        {
            if (RB_GL_DATE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SORT_TYPE.EditValue = RB_GL_DATE.RadioButtonString;
            }
        }

        private void RB_AMOUNT_Click(object sender, EventArgs e)
        {
            if (RB_AMOUNT.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SORT_TYPE.EditValue = RB_AMOUNT.RadioButtonString;
            }
        }

        private void CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(IGR_PAYMENT_SELECTED, CHECK_YN.CheckBoxValue);
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        private void ilaVENDOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaVENDOR_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        #endregion

        #region ----- Adapter event -----
        
        private void IDA_PAYMENT_SELECTED_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(PAYMENT_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAYMENT_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_BATCH_PAYMENT_ACCOUNT_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            SEARCH_DETAIL(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
        }

        private void IDA_PAYMENT_SELECTED_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            int mCHECK_YN = 0;
            int mEdit_Flag = 0;
            int mIDX_CHECK_YN = IGR_PAYMENT_SELECTED.GetColumnToIndex("CHECK_YN");
            int mIDX_GL_CURR_AMOUNT = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_CURR_AMOUNT");
            int mIDX_GL_AMOUNT = IGR_PAYMENT_SELECTED.GetColumnToIndex("GL_AMOUNT");
            int mIDX_BILL_DUE_DATE = IGR_PAYMENT_SELECTED.GetColumnToIndex("BILL_DUE_DATE");
            if (iString.ISNull(pBindingManager.DataRow["SUMMARY_FLAG"]) == "N")
            {
                mCHECK_YN = 1;
                mEdit_Flag = 1;
            }
            else
            {
                mEdit_Flag = 0;
            }
            
            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_CHECK_YN].Insertable = mCHECK_YN;
            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_CHECK_YN].Updatable = mCHECK_YN;

            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_GL_CURR_AMOUNT].Insertable = mEdit_Flag;
            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_GL_CURR_AMOUNT].Updatable = mEdit_Flag;

            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_GL_AMOUNT].Insertable = mEdit_Flag;
            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_GL_AMOUNT].Updatable = mEdit_Flag;

            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_BILL_DUE_DATE].Insertable = mEdit_Flag;
            IGR_PAYMENT_SELECTED.GridAdvExColElement[mIDX_BILL_DUE_DATE].Updatable = mEdit_Flag;

            IGR_PAYMENT_SELECTED.Refresh();
        }

        #endregion

    }
}