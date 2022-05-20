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

namespace FCMF0305
{
    public partial class FCMF0305 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0305()
        {
            InitializeComponent();
        }

        public FCMF0305(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(PERIOD_FR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PERIOD_FR_0.Focus();
                return;
            }
            if (iString.ISNull(PERIOD_TO_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10219"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PERIOD_TO_0.Focus();
                return;
            }
            IDA_DPR_MONTH.Fill();
            igrDPR_MONTH.Focus();
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
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

        //    object vBALANCE_DATE = iDate.ISGetDate(W_BALANCE_DATE.EditValue).ToShortDateString();
        //    object vACCOUNT_CODE = W_ACCOUNT_CODE.EditValue;
        //    object vACCOUNT_DESC = W_ACCOUNT_DESC_FR.EditValue;
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

        //    int vCountRow = IGR_BALANCE_APAR_DETAIL.RowCount;
        //    if (vCountRow < 1)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //    vSaveFileName = string.Format("Balance_{0}_{1}", vBALANCE_DATE, vACCOUNT_DESC);

        //    saveFileDialog1.Title = "Excel Save";
        //    saveFileDialog1.FileName = vSaveFileName;
        //    saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
        //    saveFileDialog1.DefaultExt = "xls";
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
        //        xlPrinting.OpenFileNameExcel = "FCMF0530_001.xls";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            //조회시 그룹핑 옵션.
        //            if (iString.ISNull(W_GROUPING_DUE_DATE.CheckBoxValue) == "Y")
        //            {
        //                switch (iString.ISNull(vTerritory))
        //                {
        //                    case "TL1_KR":
        //                        vGROUPING_OPTION = string.Format("내역서({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL1_KR);
        //                        break;
        //                    case "TL2_CN":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL2_CN);
        //                        break;
        //                    case "TL3_VN":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL3_VN);
        //                        break;
        //                    case "TL4_JP":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL4_JP);
        //                        break;
        //                    case "TL5_XAA":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL5_XAA);
        //                        break;
        //                    default:                                
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].Default);
        //                        break;
        //                }
        //            }
        //            else
        //            {
        //                switch (iString.ISNull(vTerritory))
        //                {
        //                    case "TL1_KR":
        //                        vGROUPING_OPTION = "내역서";
        //                        break;
        //                    case "TL2_CN":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    case "TL3_VN":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    case "TL4_JP":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    case "TL5_XAA":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    default:                                
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                }
        //            }
        //            //날짜형식 변경.
        //            IDC_DATE_YYYYMMDD.SetCommandParamValue("P_DATE", vBALANCE_DATE);
        //            IDC_DATE_YYYYMMDD.ExecuteNonQuery();
        //            vBALANCE_DATE = IDC_DATE_YYYYMMDD.GetCommandParamValue("O_DATE");
        //            xlPrinting.HeaderWrite(vACCOUNT_CODE, vACCOUNT_DESC, vBALANCE_DATE, vGROUPING_OPTION, iString.ISNull(vTerritory), IGR_BALANCE_APAR_DETAIL);

        //            // 실제 인쇄
        //            //vPageNumber = xlPrinting.LineWrite(vBALANCE_DATE, iString.ISNull(vTerritory), pGRID);
        //            vPageNumber = xlPrinting.LineWrite(IGR_BALANCE_APAR_DETAIL);

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

        #region ----- XLS Print Method ----

        private void ExcelExport(ISGridAdvEx vGrid, string pDefault_Filename)
        {
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.RestoreDirectory = true;
            vSaveFileDialog.FileName = pDefault_Filename;
            vSaveFileDialog.Filter = "Excel file(*.xls)|*.xls";
            vSaveFileDialog.DefaultExt = "xls";

            if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                vExport.GridToExcel(vGrid.BaseGrid, vSaveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = vSaveFileDialog.FileName;
                    vProc.Start();
                }
            }

        }

        #endregion;

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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(igrDPR_MONTH, "dpr month_");
                }
            }
        }

        #endregion;

        #region ------ Form Event -----

        private void FCMF0305_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0305_Shown(object sender, EventArgs e)
        {
            PERIOD_FR_0.EditValue = string.Format("{0}-{1}", iDate.ISYear(DateTime.Today), "01");
            PERIOD_TO_0.EditValue = iDate.ISYearMonth(DateTime.Today);

            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "DPR_TYPE");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            DPR_TYPE_NAME_0.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            DPR_TYPE_0.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
        }

        #endregion

        #region ----- Lookup Event ------
        
        private void ilaPERIOD_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD.SetLookupParamValue("W_START_YYYYMM", null);
        }

        private void ilaPERIOD_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD.SetLookupParamValue("W_START_YYYYMM", PERIOD_FR_0.EditValue);
            ildPERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddYears(6));
        }

        private void ILA_COSTCENTER_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaEXPENSE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EXPENSE_TYPE", "N");
        }

        private void ilaDPR_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DPR_TYPE", "N");
        }

        private void ilaASSET_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ASSET_TYPE", "N");
        }

        private void ilaASSET_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }
        
        private void ilaASSET_CODE_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CODE_FR_TO_0.SetLookupParamValue("W_ASSET_CODE", null);
        }

        private void ilaASSET_CODE_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CODE_FR_TO_0.SetLookupParamValue("W_ASSET_CODE", ASSET_CODE_FR_0.EditValue);
        }

        private void ilaASSET_CODE_FR_0_SelectedRowData(object pSender)
        {
            ASSET_CODE_TO_0.EditValue = ASSET_CODE_FR_0.EditValue;
            ASSET_DESC_TO_0.EditValue = ASSET_DESC_FR_0.EditValue;
        }

        #endregion

        #region ------ Adapter Event ------


        #endregion
        
    }
}