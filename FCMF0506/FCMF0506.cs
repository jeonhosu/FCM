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
 
namespace FCMF0506
{
    public partial class FCMF0506 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0506(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search()
        {
            //회계일자는 계정기준(1번째 탭)과 거래처기준(2번째 탭) 모두 필수이다.
            if (iString.ISNull(V_GL_DATE_FR.EditValue) == string.Empty)
            {
                //시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_FR.Focus();
                return;
            }

            if (iString.ISNull(V_GL_DATE_TO.EditValue) == string.Empty)
            {
                //종료일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(V_GL_DATE_FR.EditValue) > Convert.ToDateTime(V_GL_DATE_TO.EditValue))
            {
                //종료일은 시작일 이후이어야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10345"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_FR.Focus();
                return;
            }


            //조회조건의 필요한 값이 모두 채워졌을 경우 조건에 부합되는 자료를 조회한다.
            if (TB_MAIN.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
            {
                //계정과목은 계정기준으로 조회시에만 필수이다.
                if (iString.ISNull(V_ACCOUNT_LEVEL.EditValue) == string.Empty)
                {
                    //계정구분은 필수입니다
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10523"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    V_ACCOUNT_LEVEL_DESC.Focus();
                    return;
                }
 
                if (iString.ISNull(V_ACCOUNT_CODE_FR.EditValue) == string.Empty || iString.ISNull(V_ACCOUNT_CODE_TO.EditValue) == string.Empty)
                {
                    //계정과목은 필수입니다.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    V_ACCOUNT_CODE_FR.Focus();
                    return;
                }

                IDA_ACCOUNT_CUST_LIST.Fill();
                IGR_ACCOUNT_CUST_LIST.Focus();
            }
            else
            {
                //거래처기준(2번째 탭)으로 조회시는 시작 또는 종료 조건의 한쪽에만 값이 입력되서는 안된다.
                //시작계정과 종료계정 모두에 검색할 계정을 입력해야한다.
                if (iString.ISNull(V_VENDOR_CODE.EditValue) == string.Empty)
                {
                    //계정과목은 필수입니다.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    V_ACCOUNT_CODE_FR.Focus();
                    return;
                }       
                IDA_CUSTOMER_CUST_LIST.Fill();
                IGR_CUSTOMER_CUST_LIST.Focus();
            }  
        }


        //계정구분 조건에 값이 있는지 체크
        private void Check_ACCOUNT_LEVEL()
        {
            if (iString.ISNull(V_ACCOUNT_LEVEL.EditValue) == string.Empty)
            {
                //계정구분은 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10523"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_ACCOUNT_LEVEL_DESC.Focus();
                return;
            }
        }

        //조회된 자료에서 더블클릭하면 전표팝업 띄워준다.
        private void Show_Slip_Detail(Int32 pSLIP_HEADER_ID)
        {
            if (pSLIP_HEADER_ID != Convert.ToInt32(0))
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, pSLIP_HEADER_ID);
                vFCMF0204.Show();

                this.Cursor = System.Windows.Forms.Cursors.Default;
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

        #region ----- XL Print 1 Method -----

        private void XLPrinting(string pOutChoice)
        {
            object vPRINT_TYPE = string.Empty;
            DialogResult dlgResult;
            FCMF0506_PRINT vFCMF0506_PRINT = new FCMF0506_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgResult = vFCMF0506_PRINT.ShowDialog();
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
                 
                vPRINT_TYPE = vFCMF0506_PRINT.Get_Print_Type;
                if (iString.ISNull(vPRINT_TYPE) == "H")
                {
                    //IDA_PRINT_CUST_REMAIN.SetSelectParamValue("P_ACCOUNT_LEVEL", vACCOUNT_LEVEL);
                    //IDA_PRINT_CUST_REMAIN.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", vACCOUNT_CONTROL_ID);
                    //IDA_PRINT_CUST_REMAIN.SetSelectParamValue("P_CUSTOMER_CODE", vCUSTOMER_CODE);
                    //IDA_PRINT_CUST_REMAIN.Fill();

                    //합계 인쇄
                    XLPrinting_H(pOutChoice);
                }
                else if (iString.ISNull(vPRINT_TYPE) == "D")
                {                    
                    vACCOUNT_LEVEL = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_LEVEL");
                    vACCOUNT_CONTROL_ID = IGR_ACCOUNT_CUST_SUM.GetCellValue("ACCOUNT_CONTROL_ID");
                    vCUSTOMER_CODE = IGR_ACCOUNT_CUST_SUM.GetCellValue("SUPP_CUST_CODE");
                    
                    IDA_PRINT_CUST_DETAIL.SetSelectParamValue("P_ACCOUNT_LEVEL", vACCOUNT_LEVEL);
                    IDA_PRINT_CUST_DETAIL.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", vACCOUNT_CONTROL_ID);
                    IDA_PRINT_CUST_DETAIL.SetSelectParamValue("P_CUSTOMER_CODE", vCUSTOMER_CODE);
                    IDA_PRINT_CUST_DETAIL.Fill();

                    //상세 인쇄
                    XLPrinting_D(pOutChoice);
                }
            }
            vFCMF0506_PRINT.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_H(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = IGR_ACCOUNT_CUST_SUM.RowCount;
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
                vSaveFileName = "Customer_ledger_remain_";

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
                xlPrinting.OpenFileNameExcel = "FCMF0506_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    IDC_GL_TITLE_P.SetCommandParamValue("P_DATE_FR", V_GL_DATE_FR.EditValue);
                    IDC_GL_TITLE_P.SetCommandParamValue("P_DATE_TO", V_GL_DATE_TO.EditValue);
                    IDC_GL_TITLE_P.ExecuteNonQuery();
                    object vCOMPANY_NAME = IDC_GL_TITLE_P.GetCommandParamValue("O_COMPANY_NAME");
                    object vPERIOD_NAME = IDC_GL_TITLE_P.GetCommandParamValue("O_PERIOD_NAME"); 

                    object vACCOUNT_CODE = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_CODE");
                    object vACCOUNT_DESC = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_DESC");
                    vACCOUNT_DESC = string.Format("({0}){1}", vACCOUNT_CODE, vACCOUNT_DESC);

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_1(vCOMPANY_NAME, vPERIOD_NAME, vACCOUNT_DESC);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(IGR_ACCOUNT_CUST_SUM);

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

            int vCountRow = IDA_PRINT_CUST_DETAIL.CurrentRows.Count;
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
                vSaveFileName = "Customer_ledger_Detail";

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
                xlPrinting.OpenFileNameExcel = "FCMF0506_002.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    IDC_GL_TITLE_P.SetCommandParamValue("P_DATE_FR", V_GL_DATE_FR.EditValue);
                    IDC_GL_TITLE_P.SetCommandParamValue("P_DATE_TO", V_GL_DATE_TO.EditValue);
                    IDC_GL_TITLE_P.ExecuteNonQuery();
                    object vCOMPANY_NAME = IDC_GL_TITLE_P.GetCommandParamValue("O_COMPANY_NAME");
                    object vPERIOD_NAME = IDC_GL_TITLE_P.GetCommandParamValue("O_PERIOD_NAME");

                    object vACCOUNT_CODE = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_CODE");
                    object vACCOUNT_DESC = IGR_ACCOUNT_CUST_LIST.GetCellValue("ACCOUNT_DESC");
                    vACCOUNT_DESC = string.Format("({0}){1}", vACCOUNT_CODE, vACCOUNT_DESC);

                    object vCUSTOMER_CODE = IGR_ACCOUNT_CUST_SUM.GetCellValue("SUPP_CUST_CODE");
                    object vCUSTOMER_DESC = IGR_ACCOUNT_CUST_SUM.GetCellValue("SUPP_CUST_NAME");
                    vCUSTOMER_DESC = string.Format("({0}){1}", vCUSTOMER_CODE, vCUSTOMER_DESC);

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_2(vCOMPANY_NAME, vPERIOD_NAME, vACCOUNT_DESC, vCUSTOMER_DESC);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_2(IDA_PRINT_CUST_DETAIL);

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
                    XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                {
                    XLPrinting("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0506_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;

            int vIDX_ACC_CONFIRM_FLAG = IGR_ACCOUNT_CUST_DETAIL.GetColumnToIndex("CONFIRM_FLAG");
            int vIDX_CUST_CONFIRM_FLAG = IGR_CUSTOMER_ACC_DETAIL.GetColumnToIndex("CONFIRM_FLAG");
            if (mCONFIRM_CHECK == "Y")
            {
                GB_CONFIRM_STATUS.Visible = true;

                IGR_ACCOUNT_CUST_DETAIL.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 1;
                IGR_CUSTOMER_ACC_DETAIL.GridAdvExColElement[vIDX_CUST_CONFIRM_FLAG].Visible = 1;
            }
            else
            {
                GB_CONFIRM_STATUS.Visible = false;

                IGR_ACCOUNT_CUST_DETAIL.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 0;
                IGR_CUSTOMER_ACC_DETAIL.GridAdvExColElement[vIDX_CUST_CONFIRM_FLAG].Visible = 0;
            }

            IGR_ACCOUNT_CUST_DETAIL.ResetDraw = true;
            IGR_CUSTOMER_ACC_DETAIL.ResetDraw = true;
        }

        private void FCMF0506_Shown(object sender, EventArgs e)
        {
            V_GL_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            V_GL_DATE_TO.EditValue = System.DateTime.Today;

            V_ACC_DTL_GL_DATE_FR.EditValue = V_GL_DATE_FR.EditValue;
            V_ACC_DTL_GL_DATE_TO.EditValue = V_GL_DATE_TO.EditValue;

            V_CUST_DTL_GL_DATE_FR.EditValue = V_GL_DATE_FR.EditValue;
            V_CUST_DTL_GL_DATE_TO.EditValue = V_GL_DATE_TO.EditValue;
            
            IDC_GET_ACCOUNT_LEVEL.ExecuteNonQuery();
            V_ACCOUNT_LEVEL_DESC.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE_NAME");
            V_ACCOUNT_LEVEL.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE");
        }

        private void BTN_ACCOUNT_CUST_DETAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ACCOUNT_CUST_DETAIL.Fill();
        }

        private void BTN_CUST_ACC_DETAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CUSTOMER_ACC_DETAIL.Fill();
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
            V_CUST_DTL_GL_DATE_FR.EditValue = e.EditValue;
        }

        private void V_GL_DATE_TO_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            V_ACC_DTL_GL_DATE_TO.EditValue = e.EditValue;
            V_CUST_DTL_GL_DATE_TO.EditValue = e.EditValue;
        }

        #endregion

        #region ----- Grid Event -----

        private void IGR_ACCOUNT_CUST_DETAIL_CellDoubleClick(object pSender)
        {
            if (IGR_ACCOUNT_CUST_DETAIL.Row > 0)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_ACCOUNT_CUST_DETAIL.GetCellValue("SLIP_HEADER_ID"));

                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        private void IGR_CUSTOMER_ACC_DETAIL_CellDoubleClick(object pSender)
        {
            if (IGR_CUSTOMER_ACC_DETAIL.Row > 0)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_CUSTOMER_ACC_DETAIL.GetCellValue("SLIP_HEADER_ID"));

                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        #endregion

        #region ---- Lookup Event ----- 

        private void ILA_ACCOUNT_CODE_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Check_ACCOUNT_LEVEL();  //계정구분이 설정되어 있는지를 체크한다.

            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_LEVEL", V_ACCOUNT_LEVEL.EditValue);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_CODE", null);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ENABLED_YN", "N");
        }

        private void ILA_ACCOUNT_CODE_FR_SelectedRowData(object pSender)
        {
            V_ACCOUNT_CODE_TO.EditValue = V_ACCOUNT_CODE_FR.EditValue;
            V_ACCOUNT_DESC_TO.EditValue = V_ACCOUNT_DESC_FR.EditValue;
        }

        private void ILA_ACCOUNT_CODE_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Check_ACCOUNT_LEVEL();  //계정구분이 설정되어 있는지를 체크한다.

            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_LEVEL", V_ACCOUNT_LEVEL.EditValue);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_CODE", V_ACCOUNT_CODE_FR.EditValue);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ENABLED_YN", "N");
        }

        private void ILA_ACCOUNT_LEVEL_SelectedRowData(object pSender)
        {
            V_ACCOUNT_CODE_FR.EditValue = String.Empty;
            V_ACCOUNT_DESC_FR.EditValue = String.Empty;

            V_ACCOUNT_CODE_TO.EditValue = String.Empty;
            V_ACCOUNT_DESC_TO.EditValue = String.Empty;
        }

        private void ILA_ACCOUNT_LEVEL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_LEVEL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion


    }
}