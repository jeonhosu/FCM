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

using System.IO;
using Syncfusion.GridExcelConverter;


namespace FCMF0503
{
    public partial class FCMF0503 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCONFIRM_CHECK = "N";
        
        #endregion;

        #region ----- Constructor -----

        public FCMF0503(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            object vObject1 = V_GL_DATE_FR.EditValue;
            if (iString.ISNull(vObject1) == string.Empty)
            {
                //시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object vObject2 = V_GL_DATE_TO.EditValue;
            if (iString.ISNull(vObject2) == string.Empty)
            {
                //종료일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Convert.ToDateTime(V_GL_DATE_FR.EditValue) > Convert.ToDateTime(V_GL_DATE_TO.EditValue))
            {
                //종료일은 시작일 이후이어야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10345"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_FR.Focus();
                return;
            }

            object vObject3 = V_MANAGEMENT_NAME.EditValue;
            if (iString.ISNull(vObject3) == string.Empty)
            {
                //관리항목은 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10417"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object vObject4 = V_ACCOUNT_CODE_FR.EditValue;
            object vObject5 = V_ACCOUNT_CODE_TO.EditValue;
            if (iString.ISNull(vObject4) == string.Empty || iString.ISNull(vObject5) == string.Empty)
            {
                //계정과목은 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int vACCOUNT_CODE_FR_0 = ConvertInteger(vObject4);
            int vACCOUNT_CODE_TO_0 = ConvertInteger(vObject5);
            if (vACCOUNT_CODE_FR_0 > vACCOUNT_CODE_TO_0)
            {
                //종료계정은 시작계정 이후의 계정이어야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10414"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_ACCOUNT_CODE_FR.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
            {
                IDA_ACCOUNT_LIST.Fill();
                IGR_ACCOUNT_LIST.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_ALL.TabIndex)
            {
                IDA_ACC_MANAGEMENT_ALL.Fill();
                IGR_ALL_MANAGEMENT_LEDGER.Focus();
            }
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

        #region ----- Convert decimal  Method ----

        private int ConvertInteger(object pObject)
        {
            bool vIsConvert = false;
            int vConvertInteger = 0;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        string vString = pObject as string;
                        vConvertInteger = int.Parse(vString);
                    }
                }

            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertInteger;
        }

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx vGrid)
        {
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.RestoreDirectory = true;
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

        #endregion

        #region ----- XL Print 1 Method -----

        private void XLPrinting(string pOutChoice)
        {
            object vPRINT_TYPE = string.Empty;
            DialogResult dlgResult;
            FCMF0503_PRINT vFCMF0503_PRINT = new FCMF0503_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgResult = vFCMF0503_PRINT.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                object vACCOUNT_LEVEL = string.Empty;
                object vACCOUNT_CONTROL_ID = string.Empty;
                object vCUSTOMER_CODE = string.Empty;

                //if (TB_MAIN.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
                //{
                //    vACCOUNT_LEVEL = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_LEVEL");
                //    vACCOUNT_CONTROL_ID = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_CONTROL_ID");
                //    vCUSTOMER_CODE = V_VENDOR_CODE.EditValue;
                //}
                //else if (TB_MAIN.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
                //{
                //    vACCOUNT_LEVEL = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_LEVEL");
                //    vACCOUNT_CONTROL_ID = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_CONTROL_ID");
                //    vCUSTOMER_CODE = V_VENDOR_CODE.EditValue;
                //}

                vPRINT_TYPE = vFCMF0503_PRINT.Get_Print_Type;
                if (iString.ISNull(vPRINT_TYPE) == "H")
                {
                    //합계 인쇄
                    XLPrinting_H(pOutChoice);
                }
                else if (iString.ISNull(vPRINT_TYPE) == "D")
                {
                    IDA_PRINT_ACC_MANAGEMENT_DETAIL.Fill();

                    //상세 인쇄
                    XLPrinting_D(pOutChoice);
                }
            }
            vFCMF0503_PRINT.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_H(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = IGR_ACC_MANAGEMENT_SUM.RowCount;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Management_ledger_remain_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
                saveFileDialog1.DefaultExt = "xls";
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
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0503_001.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_FR.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    object vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    object vPeriod = vDATE_FORMAT;

                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_TO.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    vPeriod = String.Format("Period : {0} ~ {1}", vPeriod, vDATE_FORMAT);

                    object vACCOUNT_CODE = IGR_ACC_MANAGEMENT_SUM.GetCellValue("ACCOUNT_CODE");
                    object vACCOUNT_DESC = IGR_ACC_MANAGEMENT_SUM.GetCellValue("ACCOUNT_DESC");
                    vACCOUNT_DESC = string.Format("({0}){1}", vACCOUNT_CODE, vACCOUNT_DESC);

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_1(vPeriod, vACCOUNT_DESC);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(IGR_ACC_MANAGEMENT_SUM);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_D(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = IDA_PRINT_ACC_MANAGEMENT_DETAIL.CurrentRows.Count;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Management_ledger_Detail";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
                saveFileDialog1.DefaultExt = "xls";
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
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0503_002.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_FR.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    object vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    object vPeriod = vDATE_FORMAT;

                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_TO.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    vPeriod = String.Format("Period : {0} ~ {1}", vPeriod, vDATE_FORMAT);

                    object vACCOUNT_CODE = IGR_ACC_MANAGEMENT_SUM.GetCellValue("ACCOUNT_CODE");
                    object vACCOUNT_DESC = IGR_ACC_MANAGEMENT_SUM.GetCellValue("ACCOUNT_DESC");
                    vACCOUNT_DESC = string.Format("({0}){1}", vACCOUNT_CODE, vACCOUNT_DESC);

                    object vMANAGEMENT_VALUE = IGR_ACC_MANAGEMENT_SUM.GetCellValue("MANAGEMENT_VALUE");
                    object vMANAGEMENT_DESC = IGR_ACC_MANAGEMENT_SUM.GetCellValue("MANAGEMENT_DESC");
                    vMANAGEMENT_DESC = string.Format("({0}){1}", vMANAGEMENT_VALUE, vMANAGEMENT_DESC);

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_2(vPeriod, vACCOUNT_DESC, vMANAGEMENT_DESC);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_2(IDA_PRINT_ACC_MANAGEMENT_DETAIL);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- MDi ToolBar Button Event -----

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
                    if (IDA_ACCOUNT_LIST.IsFocused == true)
                    {
                        IDA_ACCOUNT_LIST.Cancel();
                    }
                    else if (IDA_ACC_MANAGEMENT_SUM.IsFocused == true)
                    {
                        IDA_ACC_MANAGEMENT_SUM.Cancel();
                    }
                    else if (IDA_ACC_MANAGEMENT_DETAIL.IsFocused == true)
                    {
                        IDA_ACC_MANAGEMENT_DETAIL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (IDA_ACC_MANAGEMENT_SUM.IsFocused)
                    {
                        ExcelExport(IGR_ACC_MANAGEMENT_SUM);
                    }
                    else if (IDA_ACC_MANAGEMENT_DETAIL.IsFocused)
                    {
                        ExcelExport(IGR_ACC_MANAGEMENT_DETAIL);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0503_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;

            int vIDX_ACC_CONFIRM_FLAG = IGR_ACC_MANAGEMENT_DETAIL.GetColumnToIndex("CONFIRM_FLAG"); 
            if (mCONFIRM_CHECK == "Y")
            {
                GB_CONFIRM_STATUS.Visible = true;

                IGR_ACC_MANAGEMENT_DETAIL.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 1; 
            }
            else
            {
                GB_CONFIRM_STATUS.Visible = false;

                IGR_ACC_MANAGEMENT_DETAIL.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 0; 
            }
            IGR_ACC_MANAGEMENT_DETAIL.ResetDraw = true; 
        }
        
        private void FCMF0503_Shown(object sender, EventArgs e)
        {
            V_GL_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            V_GL_DATE_TO.EditValue = System.DateTime.Today;

            V_ACC_DTL_GL_DATE_FR.EditValue = V_GL_DATE_FR.EditValue;
            V_ACC_DTL_GL_DATE_TO.EditValue = V_GL_DATE_TO.EditValue;
             
            IDC_GET_ACCOUNT_LEVEL.ExecuteNonQuery();
            V_ACCOUNT_LEVEL_DESC.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE_NAME");
            V_ACCOUNT_LEVEL.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE");
        }
         
        private void V_ACCOUNT_CODE_FR_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(V_ACCOUNT_CODE_FR.EditValue) == string.Empty)
            {
                V_ACCOUNT_CODE_TO.EditValue = null;
                V_ACCOUNT_DESC_TO.EditValue = null;
            }
        }

        private void V_RB_CONFIRM_ALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;

            V_CONFIRM_STATUS.EditValue = iStatus.RadioCheckedString;
        }

        private void V_GL_DATE_FR_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            V_ACC_DTL_GL_DATE_FR.EditValue = e.EditValue;
            V_GL_DATE_TO.EditValue = e.EditValue;
        }

        private void V_GL_DATE_TO_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            V_ACC_DTL_GL_DATE_TO.EditValue = e.EditValue; 
        }

        private void TAB_MAIN_Click(object sender, EventArgs e)
        {
            SearchDB(); 
        }


        private void BTN_ACC_CUST_DETAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ACC_MANAGEMENT_DETAIL.Fill();
        }

        #endregion

        #region ----- Adapter Event -----

        #endregion

        #region ----- Grid Event -----

        private void IGR_ACC_MANAGEMENT_DETAIL_CellDoubleClick(object pSender)
        {
            if (IGR_ACC_MANAGEMENT_DETAIL.RowIndex > -1)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_ACC_MANAGEMENT_DETAIL.GetCellValue("SLIP_HEADER_ID"));
                if (vSLIP_HEADER_ID > Convert.ToInt32(0))
                {
                    System.Windows.Forms.Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    
                    FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, vSLIP_HEADER_ID);
                    vFCMF0204.Show();

                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.UseWaitCursor = false;
                }
            }
        }

        private void IGR_ALL_MANAGEMENT_LEDGER_CellDoubleClick(object pSender)
        {
            //if (IGR_ALL_MANAGEMENT_LEDGER.RowIndex > -1)
            //{
            //    int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_ALL_MANAGEMENT_LEDGER.GetCellValue("SLIP_HEADER_ID"));
            //    if (vSLIP_HEADER_ID > Convert.ToInt32(0))
            //    {
            //        System.Windows.Forms.Application.UseWaitCursor = true;
            //        this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            //        FCMF0205.FCMF0205 vFCMF0205 = new FCMF0205.FCMF0205(this.MdiParent, isAppInterfaceAdv1.AppInterface, vSLIP_HEADER_ID);
            //        vFCMF0205.Show();

            //        this.Cursor = System.Windows.Forms.Cursors.Default;
            //        System.Windows.Forms.Application.UseWaitCursor = false;
            //    }
            //}
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_LEVEL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_LEVEL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_LEVEL_SelectedRowData(object pSender)
        {
            V_ACCOUNT_CODE_FR.EditValue = String.Empty;
            V_ACCOUNT_DESC_FR.EditValue = String.Empty;

            V_ACCOUNT_CODE_TO.EditValue = String.Empty;
            V_ACCOUNT_DESC_TO.EditValue = String.Empty;
        }

        private void ILA_ACCOUNT_CODE_FR_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_0.SetLookupParamValue("P_ACCOUNT_CODE", null);  
            ILD_ACCOUNT_CONTROL_0.SetLookupParamValue("P_ENABLED_YN", "Y");
        }
         
        private void ILA_ACCOUNT_CODE_TO_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_0.SetLookupParamValue("P_ACCOUNT_CODE", V_ACCOUNT_CODE_FR.EditValue);
            ILD_ACCOUNT_CONTROL_0.SetLookupParamValue("P_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CODE_FR_0_SelectedRowData(object pSender)
        {
            V_ACCOUNT_CODE_TO.EditValue = V_ACCOUNT_CODE_FR.EditValue;
            V_ACCOUNT_DESC_TO.EditValue = V_ACCOUNT_DESC_FR.EditValue;
        }

        private void ILA_MANAGEMENT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_MANAGEMENT_0.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_MANAGEMENT_0_SelectedRowData(object pSender)
        {
            V_MANAGEMENT_VALUE.EditValue = null;
            V_MANAGEMENT_VALUE_DESC.EditValue = null;
        }

        private void ILA_MANAGEMENT_ITEM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_MANAGEMENT_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

    }
}