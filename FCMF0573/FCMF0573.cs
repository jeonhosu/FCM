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


namespace FCMF0573
{
    public partial class FCMF0573 : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

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

        public FCMF0573()
        {
            InitializeComponent();
        }

        public FCMF0573(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void GetAccountBook()
        {
            idcACCOUNT_BOOK.ExecuteNonQuery();
            mSession_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_SESSION_ID");
            mAccount_Book_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = idcACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = idcACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE");
            mBudget_Control_YN = idcACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private void Search_DB()
        {
            if (iString.ISNull(W_GL_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_GL_DATE_FR.Focus();
                return;
            }

            if (iString.ISNull(W_GL_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_GL_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(W_GL_DATE_FR.EditValue) > Convert.ToDateTime(W_GL_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_GL_DATE_FR.Focus();
                return;
            }

            if (iString.ISNull(W_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
            {// 예산부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_ACCOUNT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ACCOUNT_CODE.Focus();
                return;
            }

            IDA_CASHBOOK.Fill();
            IGR_CASHBOOK.Focus();
        }

        //조회된 자료에서 더블클릭하면 전표팝업 띄워준다.
        private void Show_Slip_Detail(Int32 pSLIP_HEADER_ID)
        {
            //if (pSLIP_HEADER_ID != Convert.ToInt32(0))
            //{
            //    Application.UseWaitCursor = true;
            //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            //    FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, pSLIP_HEADER_ID);
            //    vFCMF0204.Show();

            //    this.Cursor = System.Windows.Forms.Cursors.Default;
            //    Application.UseWaitCursor = false;
            //}
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

        #region ----- XLS Print Method ----

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

                vExport.ExportStyle = false;
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
                    Search_DB();
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
                        XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                        XLPrinting1("EXCEL");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0573_Load(object sender, EventArgs e)
        {
            W_GL_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_GL_DATE_TO.EditValue = iDate.ISGetDate();

            IDC_GET_ACCOUNT_LEVEL.ExecuteNonQuery();
            V_ACCOUNT_LEVEL_DESC.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE_NAME");
            V_ACCOUNT_LEVEL.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE");

            IDC_GET_MULTI_LANG_P.ExecuteNonQuery();
            string vMULTI_LANG_FLAG = iString.ISNull(IDC_GET_MULTI_LANG_P.GetCommandParamValue("O_MULTI_LANG_FLAG"));
            if (vMULTI_LANG_FLAG == "Y")
            {
                V_LANG_DESC.Visible = true;
                V_LANG_DESC.BringToFront();
                IDC_GET_LANG_CODE.ExecuteNonQuery();
                V_LANG_DESC.EditValue = IDC_GET_LANG_CODE.GetCommandParamValue("O_LANG_DESC");
                V_LANG_CODE.EditValue = IDC_GET_LANG_CODE.GetCommandParamValue("O_LANG_CODE");
            }
            else
            {
                V_LANG_DESC.Visible = false;
                V_LANG_CODE.EditValue = null;
            }
            V_LANG_DESC.BringToFront(); 

            // 회계장부 정보 설정.
            GetAccountBook();
        }

        private void igrCASHBOOK_CellDoubleClick(object pSender)
        {
            if (IGR_CASHBOOK.Row > 0)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_CASHBOOK.GetCellValue("SLIP_HEADER_ID"));

                Show_Slip_Detail(vSLIP_HEADER_ID);
            } 
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CODE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_W.SetLookupParamValue("P_ACCOUNT_CLASS_CODE", "802");
            ILD_ACCOUNT_CONTROL_W.SetLookupParamValue("P_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_LEVEL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_LEVEL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_LEVEL_SelectedRowData(object pSender)
        {
            W_ACCOUNT_CODE.EditValue = null;
            W_ACCOUNT_DESC.EditValue = null;
            W_ACCOUNT_CONTROL_ID.EditValue = null;
            V_CODE.EditValue = null;
            V_DESC.EditValue = null;
        }

        #endregion

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1(string pOutput_Type)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            int vCountRowGrid = IGR_CASHBOOK.RowCount;
            //if ((itbSLIP.SelectedIndex == 0 && vCountRowGrid > 0) ||
            //    (itbSLIP.SelectedIndex == 1 && iString.ISNull(H_SLIP_HEADER_ID.EditValue) != string.Empty))
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting", vPageTotal);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0573_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        int vCountRow = 0;
                        int vRow = IGR_CASHBOOK.RowIndex;


                        //인쇄일자 
                        IDC_GET_DATE.ExecuteNonQuery();
                        object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");


                        // 계정코드 분류에 따른 값 가져오기. Start ////////////////////////////////
                        object vACCOUNT_CODE = W_ACCOUNT_CODE.EditValue;
                        object vACCOUNT_DESC = W_ACCOUNT_DESC.EditValue;

                        object vACCOUNT_CODE_R = V_CODE.EditValue;
                        object vACCOUNT_DESC_R = V_DESC.EditValue;

                        object vAccount_default = V_ACCOUNT_LEVEL.EditValue;


                        if (vAccount_default.ToString() == "10")
                        {
                            xlPrinting.HeaderWrite(vLOCAL_DATE, vACCOUNT_CODE, vACCOUNT_DESC);
                        }
                        else
                        {
                            xlPrinting.HeaderWrite(vLOCAL_DATE, vACCOUNT_CODE_R, vACCOUNT_DESC_R);
                        }
                        // 계정코드 분류에 따른 값 가져오기. End ///////////////////////////////////


                        vCountRow = IDA_CASHBOOK.CurrentRows.Count;
                        if (vCountRow > 0)
                        {
                            vPageNumber = xlPrinting.LineWrite(IDA_CASHBOOK);
                        }

                        if (pOutput_Type == "PRINT")
                        {//[PRINT]
                            ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                            xlPrinting.PreView(1, vPageNumber);

                        }
                        else if (pOutput_Type == "EXCEL")
                        {
                            ////[SAVE]
                            xlPrinting.Save("SLIP_"); //저장 파일명
                        }

                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();

                    System.Windows.Forms.Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();

                    return;
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

    }
}