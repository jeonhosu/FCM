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


namespace FCMF0585
{
    public partial class FCMF0585 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

       // string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0585()
        {
            InitializeComponent();
        }

        public FCMF0585(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {

            int vIDX_Col = IGR_AR_REMAIN_LIST.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            decimal vACCOUNT_CONTROL_ID = iString.ISDecimaltoZero(IGR_AR_REMAIN_LIST.GetCellValue("ACCOUNT_CONTROL_ID"), 0);

            IDA_AR_REMAIN_LIST.Fill();
            
            //focus 이동.
            if (vACCOUNT_CONTROL_ID <= 0)
            {
                //
            }
            for (int nRow = 0; nRow < IGR_AR_REMAIN_LIST.RowCount; nRow++)
            {
                if (vACCOUNT_CONTROL_ID == Convert.ToInt32(iString.ISDecimaltoZero(IGR_AR_REMAIN_LIST.GetCellValue(nRow, vIDX_Col), 0)))
                {
                    IGR_AR_REMAIN_LIST.CurrentCellMoveTo(nRow, 1);
                    IGR_AR_REMAIN_LIST.CurrentCellActivate(nRow, 1); 
                    return;
                }
            }
            IGR_AR_REMAIN_LIST.Focus();
        }

        private void SearchDB_DTL(object pACCOUNT_CONTROL_ID, object pCURRENCY_CODE, object pITEM_GROUP_ID)
        {
            IDA_APAR_REMAIN_DTL.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_APAR_REMAIN_DTL.SetSelectParamValue("W_CURRENCY_CODE", pCURRENCY_CODE);
            IDA_APAR_REMAIN_DTL.SetSelectParamValue("W_ITEM_GROUP_ID", pITEM_GROUP_ID);
            IDA_APAR_REMAIN_DTL.Fill();
            IGR_APAR_REMAIN_DTL.Focus();
        }

        private void Show_Slip_Detail()
        {
            int mSLIP_HEADER_ID = iString.ISNumtoZero(IGR_APAR_REMAIN_DTL.GetCellValue("SLIP_HEADER_ID"));
            if (mSLIP_HEADER_ID != Convert.ToInt32(0))
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
                vFCMF0204.Show();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.UseWaitCursor = false;
            }
        }
          
        #endregion;

        #region ----- Territory Get Methods ----

        private object GetTerritory()
        {
            
            object vTerritory = "Default";
            vTerritory = isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage;
            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            try
            {
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
            }
            catch
            {
            }
            return mPrompt;
        }

        #endregion;

        #region ----- XL Print 1 (계정잔액명세서) Method ----

        //private void XLPrinting_1(string pOutChoice)
        //{// pOutChoice : 출력구분.
        //    string vMessageText = string.Empty;
        //    string vSaveFileName = string.Empty;

        //    object vBALANCE_DATE = iDate.ISGetDate(W_GL_DATE_TO.EditValue).ToShortDateString();
        //    object vACCOUNT_CODE = W_ACCOUNT_CODE.EditValue;
        //    object vACCOUNT_DESC = W_ACCOUNT_DESC.EditValue;
        //    object vTerritory = string.Empty;
        //    object vGROUPING_OPTION = null;
            
        //    if (iString.ISNull(vBALANCE_DATE) == String.Empty)
        //    {//기준일자
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    if (iString.ISNull(vACCOUNT_CODE) == String.Empty)
        //    {//계정과목코드
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    int vCountRow = IGR_AR_REMAIN_LIST.RowCount;
        //    if (vCountRow < 1)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //    vSaveFileName = string.Format("Balance_{0}_{1}", vBALANCE_DATE, vACCOUNT_DESC);

        //    saveFileDialog1.Title = "Excel Save";
        //    saveFileDialog1.FileName = vSaveFileName;
        //    saveFileDialog1.Filter = "Excel file(*.xlsx)|*.xlsx";
        //    saveFileDialog1.DefaultExt = "xlsx";
        //    if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        //    {
        //        return;
        //    }
        //    else
        //    {
        //        vSaveFileName = saveFileDialog1.FileName;
        //        System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
        //        try
        //        {
        //            if (vFileName.Exists)
        //            {
        //                vFileName.Delete();
        //            }
        //        }
        //        catch (Exception EX)
        //        {
        //            MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    System.Windows.Forms.Application.DoEvents();

        //    int vPageNumber = 0;

        //    vMessageText = string.Format(" Printing Starting...");
        //    isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //    System.Windows.Forms.Application.DoEvents();

        //    vTerritory = GetTerritory();
        //    XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //    try
        //    {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                
        //        // open해야 할 파일명 지정.
        //        //-------------------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = "FCMF0585_001.xlsx";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            switch (iString.ISNull(vTerritory))
        //            {
        //                case "TL1_KR":
        //                    vGROUPING_OPTION = "내역서";
        //                    break;
        //                case "TL2_CN":
        //                    vGROUPING_OPTION = "Detailed Statement";
        //                    break;
        //                case "TL3_VN":
        //                    vGROUPING_OPTION = "Detailed Statement";
        //                    break;
        //                case "TL4_JP":
        //                    vGROUPING_OPTION = "Detailed Statement";
        //                    break;
        //                case "TL5_XAA":
        //                    vGROUPING_OPTION = "Detailed Statement";
        //                    break;
        //                default:                                
        //                    vGROUPING_OPTION = "Detailed Statement";
        //                    break;
        //            }
        //            //날짜형식 변경.
        //            IDC_DATE_YYYYMMDD.SetCommandParamValue("P_DATE", vBALANCE_DATE);
        //            IDC_DATE_YYYYMMDD.ExecuteNonQuery();
        //            vBALANCE_DATE = IDC_DATE_YYYYMMDD.GetCommandParamValue("O_DATE");
        //            xlPrinting.HeaderWrite(vACCOUNT_CODE, vACCOUNT_DESC, vBALANCE_DATE, vGROUPING_OPTION, iString.ISNull(vTerritory), IGR_AR_REMAIN_LIST);

        //            // 실제 인쇄
        //            //vPageNumber = xlPrinting.LineWrite(vBALANCE_DATE, iString.ISNull(vTerritory), pGRID);
        //            vPageNumber = xlPrinting.LineWrite(IGR_AR_REMAIN_LIST.RowCount);

        //            //출력구분에 따른 선택(인쇄 or file 저장)
        //            if (pOutChoice == "PRINT")
        //            {
        //                xlPrinting.Printing(1, vPageNumber);
        //            }
        //            else if (pOutChoice == "FILE")
        //            {

        //                xlPrinting.SAVE(vSaveFileName);
        //            }

        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.Dispose();
        //            //-------------------------------------------------------------------------------------

        //            vMessageText = "Printing End";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        else
        //        {
        //            vMessageText = "Excel File Open Error";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        //-------------------------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        xlPrinting.Dispose();

        //        vMessageText = ex.Message;
        //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Application.DoEvents();
        //}

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
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_AR_REMAIN_LIST.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_AR_REMAIN_LIST.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_AR_REMAIN_LIST.IsFocused)
                    {
                        IDA_AR_REMAIN_LIST.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_AR_REMAIN_LIST);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0585_Load(object sender, EventArgs e)
        {
        }

        private void FCMF0585_Shown(object sender, EventArgs e)
        {
            
            IDA_AR_REMAIN_LIST.FillSchema();  
        }

  
        private void IGR_APAR_REMAIN_DTL_CellDoubleClick(object pSender)
        {
            if (IGR_APAR_REMAIN_DTL.RowIndex < 0)
            {
                return;
            }
            Show_Slip_Detail();
        }

        private void BTN_BSD_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SearchDB_DTL(IGR_AR_REMAIN_LIST.GetCellValue("ACCOUNT_CONTROL_ID"), IGR_AR_REMAIN_LIST.GetCellValue("CURRENCY_CODE"), IGR_AR_REMAIN_LIST.GetCellValue("ITEM_GROUP_ID"));
        } 

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", null);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_OPERATION_DIVISION_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OPERATION_DIVISION");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "N");
        }
                   
        private void ILA_VENDOR_CODE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_CODE.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaMANAGEMENT_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PAYMENT_METHOD.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_AR_REMAIN_LIST_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            object vACCOUNT_CONTROL_ID = DBNull.Value;
            object vCURRENCY_CODE = DBNull.Value;
            object vITEM_GROUP_ID = DBNull.Value;
            if (pBindingManager.DataRow != null)
            {
                vACCOUNT_CONTROL_ID = pBindingManager.DataRow["ACCOUNT_CONTROL_ID"];
                vCURRENCY_CODE = pBindingManager.DataRow["CURRENCY_CODE"];
                vITEM_GROUP_ID= pBindingManager.DataRow["ITEM_GROUP_ID"];
            }

            SearchDB_DTL(vACCOUNT_CONTROL_ID, vCURRENCY_CODE, vITEM_GROUP_ID);
        }

        #endregion

    }
}