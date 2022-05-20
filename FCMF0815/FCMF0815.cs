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

namespace FCMF0815
{
    public partial class FCMF0815 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();
        Object mSESSION_ID; 
        string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0815(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void Set_Default_Value()
        {
            //세금계산서 발행기간.
            DateTime vGetDateTime = GetDate();

            //사업장 구분.
            IDC_GET_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
            IDC_GET_DEFAULT_VALUE.ExecuteNonQuery();
            //W_TAX_CODE_NAME.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            //W_TAX_CODE.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE");

            //조회기간 기준
            IDC_GET_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "VAT_SELECT_PERIOD");
            IDC_GET_DEFAULT_VALUE.ExecuteNonQuery();
            W_VAT_SELECT_PERIOD_DESC.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_VAT_SELECT_PERIOD.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE");

            W_VAT_DATE_FR.EditValue = iDate.ISMonth_1st(vGetDateTime);
            W_VAT_DATE_TO.EditValue = iDate.ISMonth_Last(vGetDateTime);
        }


        private void SearchDB()
        {             
            if (iString.ISNull(W_VAT_SELECT_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_VAT_ISSUED_CATEGORY))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_SELECT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(W_VAT_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(W_VAT_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_DATE_FR.Focus();
                return;
            }
            if (Convert.ToDateTime(W_VAT_DATE_FR.EditValue) > Convert.ToDateTime(W_VAT_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_DATE_FR.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_NTS.TabIndex)
            {
                IDA_VAT_NTS.Fill();
                IGR_VAT_NTS.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_PURCHASE.TabIndex)
            {
                IDA_VAT_CHECK_PURCHASE.Fill();
                IGR_VAT_CHECK_PURCHASE.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALES.TabIndex)
            {
                IDA_VAT_CHECK_SALES.Fill();
                IGR_VAT_CHECK_SALES.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_ADDITIONAL_TAX.TabIndex)
            {
                IDA_VAT_ADDITIONAL_TAX.Fill();
                IDA_VAT_ADDITIONAL_TAX_DTL.Fill();
            }
        }
         
        private void Show_Slip_Interface_Detail()
        {
            //System.Windows.Forms.DialogResult vdlgResultValue;
            //int mHEADER_INTERFACE_ID = iString.ISNumtoZero(igrNOT_CONFIRM_VAT.GetCellValue("HEADER_INTERFACE_ID"));
            //if (mHEADER_INTERFACE_ID != Convert.ToInt32(0))
            //{
            //    Application.UseWaitCursor = true;
            //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            //    Application.DoEvents();

            //    Form vSLIP_IF_DETAIL = new SLIP_IF_DETAIL(isAppInterfaceAdv1.AppInterface, mHEADER_INTERFACE_ID);
            //    vdlgResultValue = vSLIP_IF_DETAIL.ShowDialog();
            //    vSLIP_IF_DETAIL.Dispose();

            //    Application.DoEvents();
            //    this.Cursor = System.Windows.Forms.Cursors.Default;
            //    Application.UseWaitCursor = false;
            //}
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Show_Slip_Detail(object pSLIP_HEADER_ID)
        {
            if (iString.ISDecimaltoZero(pSLIP_HEADER_ID) != 0)
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

        #region ----- XL Export Methods ----

        private void ExportXL()
        {
            int vCountRow = IDA_VAT_NTS.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xls";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
            //if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    string vsSaveExcelFileName = saveFileDialog1.FileName;
            //    XL.XLPrint xlExport = new XL.XLPrint();
            //    bool vXLSaveOK = xlExport.XLExport(idaVAT_MASTER.OraSelectData, vsSaveExcelFileName, vsSheetName);
            //    if (vXLSaveOK == true)
            //    {
            //        vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
            //        MessageBoxAdv.Show(vsMessage);
            //    }
            //    else
            //    {
            //        vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
            //        MessageBoxAdv.Show(vsMessage);
            //    }
            //    xlExport.XLClose();
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

        #region ----- XL Print 1 Methods ----

        //private void XLPrinting1()
        //{
        //    string vMessageText = string.Empty;

        //    XLPrinting xlPrinting = new XLPrinting();

        //    try
        //    {
        //        //mMainForm
        //        //string vPathReport = string.Empty;
        //        //object vObject = mMainForm.Tag;
        //        //if (vObject != null)
        //        //{
        //        //    vPathReport = mMainForm.Tag;
        //        //}
        //        //-------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = @"K:\00_5_XL_Print\Ex_XL_Print\XL_Print_02.xls";
        //        xlPrinting.XLFileOpen();

        //        int vTerritory = GetTerritory(igrVAT_MASTER.TerritoryLanguage);
        //        string vPeriodFrom = W_ISSUE_DATE_FR.EditText;
        //        string vPeriodTo = W_ISSUE_DATE_FR.EditText;
        //        int vPageNumber = xlPrinting.XLWirte(igrVAT_MASTER, vTerritory, vPeriodFrom, vPeriodTo);

        //        //xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
        //        ////xlPrinting.Printing(3, 4);


        //        xlPrinting.Save("t_XL_"); //저장 파일명
        //        //vMessageText = string.Format("Err : {0}", xlPrinting.ErrorMessage);
        //        //MessageGrid(vMessageText);

        //        //xlPrinting.PreView();

        //        xlPrinting.Dispose();
        //        //-------------------------------------------------------------------------

        //        vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
        //        MessageBoxAdv.Show(vMessageText);
        //    }
        //    catch (System.Exception ex)
        //    {
        //        string vMessage = ex.Message;
        //        xlPrinting.Dispose();
        //    }
        //}

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
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0815_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));


            int vIDX_CONFIRM_FLAG_P = IGR_VAT_CHECK_PURCHASE.GetColumnToIndex("CONFIRM_YN");
            int vIDX_CONFIRM_FLAG_S = IGR_VAT_CHECK_SALES.GetColumnToIndex("CONFIRM_YN");
            if (mCONFIRM_CHECK == "Y")
            {
                IGR_VAT_CHECK_PURCHASE.GridAdvExColElement[vIDX_CONFIRM_FLAG_P].Visible = 1;
                IGR_VAT_CHECK_SALES.GridAdvExColElement[vIDX_CONFIRM_FLAG_S].Visible = 1;
            }
            else
            {
                IGR_VAT_CHECK_PURCHASE.GridAdvExColElement[vIDX_CONFIRM_FLAG_P].Visible = 0;
                IGR_VAT_CHECK_SALES.GridAdvExColElement[vIDX_CONFIRM_FLAG_S].Visible = 0;
            }
            IGR_VAT_CHECK_PURCHASE.ResetDraw = true;
            IGR_VAT_CHECK_SALES.ResetDraw = true;
        }
        
        private void FCMF0815_Shown(object sender, EventArgs e)
        {
            Set_Default_Value();              
        }

        private void BTN_EXCEL_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //FCMF0815_UPLOAD vFCMF0815_UPLOAD = new FCMF0815_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            //vFCMF0815_UPLOAD.ShowDialog();
            //vFCMF0815_UPLOAD.Dispose();

            DialogResult vdlgResult;
            FCMF0815_IMPORT vFCMF0815_IMPORT = new FCMF0815_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSESSION_ID, 
                                                                    W_VAT_SELECT_PERIOD.EditValue, W_VAT_SELECT_PERIOD_DESC.EditValue,
                                                                    W_VAT_DATE_FR.EditValue, W_VAT_DATE_TO.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0815_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vFCMF0815_IMPORT.ShowDialog();
            vFCMF0815_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                SearchDB();
            }
        }
         
        private void IGR_VAT_NTS_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail(IGR_VAT_NTS.GetCellValue("SLIP_HEADER_ID"));
        }
        
        private void IGR_VAT_CHECK_PURCHASE_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail(IGR_VAT_CHECK_PURCHASE.GetCellValue("SLIP_HEADER_ID"));
        }

        private void IGR_VAT_CHECK_SALES_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail(IGR_VAT_CHECK_SALES.GetCellValue("SLIP_HEADER_ID"));
        }

        private void BTN_RESET_ADDITIONAL_TAX_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_TAX_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TAX_CODE_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_VAT_SELECT_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_VAT_ISSUED_CATEGORY))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_SELECT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(W_VAT_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(W_VAT_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_DATE_FR.Focus();
                return;
            }
            if (Convert.ToDateTime(W_VAT_DATE_FR.EditValue) > Convert.ToDateTime(W_VAT_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_DATE_FR.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_VAT_ADDITIONAL_TAX.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SET_VAT_ADDITIONAL_TAX.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_VAT_ADDITIONAL_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_VAT_ADDITIONAL_TAX.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
                return;
            }
            SearchDB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaTAX_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        }

        private void ILA_VAT_SELECT_PERIOD_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_SELECT_PERIOD", "Y");
        }

        private void ilaVAT_CLASS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_GUBUN", "Y");
        }

        private void ILA_TAX_ELECTRO_PUR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_ELECTRO", "Y");
        }

        private void ILA_TAX_ELECTRO_SAL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_ELECTRO", "Y");
        }
        
        private void ilaCUSTOMER_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaVAT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_TYPE", "Y");
        }

        private void ILA_VAT_ISSUED_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_ISSUED_CATEGORY", "Y");
        }

        #endregion

    }
}