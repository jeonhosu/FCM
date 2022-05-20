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

namespace FCMF0531
{
    public partial class FCMF0531 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0531()
        {
            InitializeComponent();
        }

        public FCMF0531(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(W_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BALANCE_DATE.Focus();
                return;
            }
            IDA_APAR_AGING_SUM.Fill();
            IGR_APAR_AGING_SUM.Focus(); 
        }
          
        private void Show_Slip_Detail()
        {
            int mSLIP_HEADER_ID = iString.ISNumtoZero(IGR_APAR_AGING_DETAIL.GetCellValue("SLIP_HEADER_ID"));
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
        //        xlPrinting.OpenFileNameExcel = "FCMF0531_001.xls";
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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (IDA_APAR_AGING_SUM.IsFocused)
                    {
                        ExcelExport(IGR_APAR_AGING_SUM, "apar_aging_sum");
                    }
                    else if (IDA_APAR_AGING_DETAIL.IsFocused)
                    {
                        ExcelExport(IGR_APAR_AGING_DETAIL, "apar_aging_detail");
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0531_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_CONFIRM_STATUS.EditValue = V_RB_CONFIRM_ALL.RadioCheckedString;

            V_RB_AP.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_BATCH_TYPE.EditValue = V_RB_AP.RadioCheckedString;
        }

        private void FCMF0531_Shown(object sender, EventArgs e)
        {
            W_BALANCE_DATE.EditValue = DateTime.Today; 
        }
         
        private void V_RB_AP_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_AP.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_BATCH_TYPE.EditValue = V_RB_AP.RadioCheckedString;
            }
        }

        private void V_RB_AR_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_AR.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_BATCH_TYPE.EditValue = V_RB_AR.RadioCheckedString;
            }
        }

        private void IGR_BALANCE_APAR_DETAIL_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_BATCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_BATCH.SetLookupParamValue("W_ACCOUNT_CODE", null);
            ILD_ACCOUNT_CONTROL_BATCH.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_SUPP_CUST_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_SUPP_CUST.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
          
        private void ILA_BIZ_LOCATION_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "BIZ_LOCATION");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        #endregion

        #region ----- Adapter Event -----
         
        #endregion


    }
}