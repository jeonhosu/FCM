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

namespace FCMF0575
{
    public partial class FCMF0575 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0575()
        {
            InitializeComponent();
        }

        public FCMF0575(Form pMainForm, ISAppInterface pAppInterface)
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

            //Set_Tab_Focus();
            if (TB_DUE_DATE.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
            {
                IDA_DUE_DATE_ACCOUNT.Fill();
                IGR_DUE_DATE_ACCOUNT.Focus();
            }
            else if (TB_DUE_DATE.SelectedTab.TabIndex == TP_VENDOR.TabIndex)
            {
                IDA_DUE_DATE_VENDOR.Fill();
                IGR_DUE_DATE_VENDOR.Focus();
            }
        }
            
        private void Show_Slip_Detail(object pSLIP_HEADER_ID)
        {
            int mSLIP_HEADER_ID = iString.ISNumtoZero(pSLIP_HEADER_ID);
            if (mSLIP_HEADER_ID != Convert.ToInt32(0))
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

        private void XLPrinting_1(string pOutChoice, string pSort_Type, ISGridAdvEx pGrid)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            string vBALANCE_DATE = iDate.ISGetDate(W_BALANCE_DATE.EditValue).ToShortDateString();            
            
            if (iString.ISNull(vBALANCE_DATE) == String.Empty)
            {//기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int vCountRow = pGrid.RowCount;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            vSaveFileName = string.Format("Due list_{0}_{1}", pSort_Type, vBALANCE_DATE);
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                
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
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();
           
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0575_001.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    //날짜형식 변경.
                    IDC_DATE_YYYYMMDD.SetCommandParamValue("P_DATE", vBALANCE_DATE);
                    IDC_DATE_YYYYMMDD.ExecuteNonQuery();

                    vBALANCE_DATE = iString.ISNull(IDC_DATE_YYYYMMDD.GetCommandParamValue("O_DATE"));
                    xlPrinting.HeaderWrite(pSort_Type, vBALANCE_DATE, pGrid, vLOCAL_DATE);

                    // 실제 인쇄
                    if (pSort_Type == "ACCOUNT")
                    {
                        vPageNumber = xlPrinting.LineWrite_Account(pGrid);
                    }
                    else
                    {
                        vPageNumber = xlPrinting.LineWrite_Vendor(pGrid);
                    }

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

                    vMessageText = "Printing End";
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
                    IDA_DUE_DATE_ACCOUNT.Cancel();
                    IDA_DUE_DATE_VENDOR.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (TB_DUE_DATE.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
                    {
                        XLPrinting_1("PRINT", "ACCOUNT", IGR_DUE_DATE_ACCOUNT);
                    }
                    else if (TB_DUE_DATE.SelectedTab.TabIndex == TP_VENDOR.TabIndex)
                    {
                        XLPrinting_1("PRINT", "VENDOR", IGR_DUE_DATE_VENDOR);
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (TB_DUE_DATE.SelectedTab.TabIndex == TP_ACCOUNT.TabIndex)
                    {
                        XLPrinting_1("FILE", "ACCOUNT", IGR_DUE_DATE_ACCOUNT);
                    }
                    else if (TB_DUE_DATE.SelectedTab.TabIndex == TP_VENDOR.TabIndex)
                    {
                        XLPrinting_1("FILE", "VENDOR", IGR_DUE_DATE_VENDOR);
                    } 
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0575_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;

            int vIDX_ACC_CONFIRM_FLAG = IGR_DUE_DATE_ACCOUNT.GetColumnToIndex("CONFIRM_FLAG");
            int vIDX_VEN_CONFIRM_FLAG = IGR_DUE_DATE_VENDOR.GetColumnToIndex("CONFIRM_FLAG"); 
            
            if (mCONFIRM_CHECK == "Y")
            {
                GB_CONFIRM_STATUS.Visible = true;

                IGR_DUE_DATE_ACCOUNT.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 1;
                IGR_DUE_DATE_VENDOR.GridAdvExColElement[vIDX_VEN_CONFIRM_FLAG].Visible = 1; 
            }
            else
            {
                GB_CONFIRM_STATUS.Visible = false;

                IGR_DUE_DATE_ACCOUNT.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 0;
                IGR_DUE_DATE_VENDOR.GridAdvExColElement[vIDX_VEN_CONFIRM_FLAG].Visible = 0; 
            }

            IGR_DUE_DATE_ACCOUNT.ResetDraw = true;
            IGR_DUE_DATE_VENDOR.ResetDraw = true;
        }

        private void FCMF0575_Shown(object sender, EventArgs e)
        {
            W_BALANCE_DATE.EditValue = DateTime.Today;

            IDA_DUE_DATE_ACCOUNT.FillSchema();
            IDA_DUE_DATE_VENDOR.FillSchema();
        }

        private void V_RB_CONFIRM_ALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
         
            V_CONFIRM_STATUS.EditValue = iStatus.RadioCheckedString;
        }

        private void V_AP_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_BATCH_TYPE.EditValue = iStatus.RadioCheckedString;
        }

        private void IGR_DUE_DATE_ACCOUNT_CellDoubleClick(object pSender)
        {
            if (IGR_DUE_DATE_ACCOUNT.RowIndex < 0)
            {
                return;
            }
            object vSLIP_HEADER_ID = IGR_DUE_DATE_ACCOUNT.GetCellValue("SLIP_HEADER_ID");
            Show_Slip_Detail(vSLIP_HEADER_ID);
        }

        private void IGR_DUE_DATE_VENDOR_CellDoubleClick(object pSender)
        {
            if (IGR_DUE_DATE_ACCOUNT.RowIndex < 0)
            {
                return;
            }
            object vSLIP_HEADER_ID = IGR_DUE_DATE_ACCOUNT.GetCellValue("SLIP_HEADER_ID");
            Show_Slip_Detail(vSLIP_HEADER_ID);
        }
          
        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_TO_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", W_ACCOUNT_CODE_FR.EditValue);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_ACCOUNT_CODE_FR_SelectedRowData(object pSender)
        {
            W_ACCOUNT_CODE_TO.EditValue = W_ACCOUNT_CODE_FR.EditValue;
            W_ACCOUNT_DESC_TO.EditValue = W_ACCOUNT_DESC_FR.EditValue;
        }

        private void ilaMANAGEMENT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y");
        } 

        #endregion

        #region ----- Adapter Event -----
         
        #endregion


    }
}