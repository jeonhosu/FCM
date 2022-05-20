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

namespace FCMF0212
{
    public partial class FCMF0212 : Office2007Form
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
        string mCurrency_Code;
        object mBudget_Control_YN;

        string mPrintOptionFlag; 

        #endregion;

        #region ----- Constructor -----

        public FCMF0212()
        {
            InitializeComponent();
        }

        public FCMF0212(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
             
            mPrintOptionFlag = "N";

            //HEADER_ID.EditValue = 21925;
            //SLIP_DATE.EditValue = "2017-04-12";
            //SLIP_NUM.EditValue = "AP-201704-0011";
            //mPrintOptionFlag = "Y";
        }

        public FCMF0212(Form pMainForm, ISAppInterface pAppInterface, 
                        object pSlip_Header_ID, object pSlip_Date, object pSlip_Num, 
                        string pPrintOptionFlag)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            HEADER_ID.EditValue = pSlip_Header_ID;
            SLIP_DATE.EditValue = pSlip_Date;
            SLIP_NUM.EditValue = pSlip_Num;

            mPrintOptionFlag = pPrintOptionFlag;
        }
          
        #endregion;

        #region ----- Private Methods -----

        private void GetAccountBook()
        {
            idcACCOUNT_BOOK.ExecuteNonQuery();
            mSession_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_SESSION_ID");
            mAccount_Book_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = idcACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = idcACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private void Search_DB()
        {
             
        }
                    
        #endregion; 

        #region ----- XL Export Methods ----

        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.CurrentRows.Count;
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
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(pAdapter.OraSelectData, vsSaveExcelFileName, vsSheetName);
                if (vXLSaveOK == true)
                {
                    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                else
                {
                    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                xlExport.XLClose();
            }
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

        private void XLPrinting_Main(string pOutput_Type)
        {
            string vSaveFileName = string.Empty;
            if (pOutput_Type == "EXCEL")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "xls file(*.xls)|*.xls|(*.xlsx)|*.xlsx";
                vSaveFileDialog.DefaultExt = "xlsx";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }
            else if (pOutput_Type == "PDF")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "pdf file(*.pdf)|*.pdf";
                vSaveFileDialog.DefaultExt = "pdf";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            idaSLIP_HEADER.Fill();
            if(idaSLIP_HEADER.CurrentRows.Count == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.Cancel;
                this.Close();
                return;
            }

            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", SLIP_DATE.EditValue);
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0212");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            if (vREPORT_TYPE.ToUpper() == "BSK")
            {
                XLPrinting_BSK(pOutput_Type, vSaveFileName);
            }
            else if (vREPORT_TYPE.ToUpper() == "SEK")
            {
                XLPrinting_SEK(pOutput_Type, vSaveFileName);
            }
            else
            {
                XLPrinting(pOutput_Type, vSaveFileName);
            }
        }

        private void XLPrinting(string pOutput_Type, string pSaveFileName)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0; 
            
            vMessageText = string.Format("Printing Starting", vPageTotal);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0212_001.xls";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    object vObject;
                    int vCountRow = 0;
                    
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");
                    
                     
                    xlPrinting.HeaderWrite(idaSLIP_HEADER, vLOCAL_DATE);
                    vObject = HEADER_ID.EditValue;

                    idaPRINT_SLIP_LINE.SetSelectParamValue("W_HEADER_ID", vObject);
                    idaPRINT_SLIP_LINE.Fill();

                    vCountRow = idaPRINT_SLIP_LINE.CurrentRows.Count;
                    if (vCountRow > 0)
                    {
                        vPageNumber = xlPrinting.LineWrite(idaPRINT_SLIP_LINE);
                    }

                    if (pOutput_Type == "PREVIEW")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PreView(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PRINT")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.Printing(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PDF")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PDF(pSaveFileName);

                    }
                    else if (pOutput_Type == "EXCEL")
                    {
                        ////[SAVE]
                        xlPrinting.Save(pSaveFileName); //저장 파일명
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

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_BSK(string pOutput_Type, string pSaveFileName)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;
         
            vMessageText = string.Format("Printing Starting", vPageTotal);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0212_011.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                { 
                    int vCountRow = 0;
                     
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    //회계법인명.
                    IDC_GET_COMPANY_NAME_P.ExecuteNonQuery();
                    object vSOB_DESC = IDC_GET_COMPANY_NAME_P.GetCommandParamValue("O_SOB_DESC");

                    xlPrinting.HeaderWrite_BSK(idaSLIP_HEADER, vSOB_DESC, vLOCAL_DATE);

                    idaPRINT_SLIP_LINE.SetSelectParamValue("W_HEADER_ID", HEADER_ID.EditValue);
                    idaPRINT_SLIP_LINE.Fill();

                    vCountRow = idaPRINT_SLIP_LINE.CurrentRows.Count;
                    if (vCountRow > 0)
                    {
                        vPageNumber = xlPrinting.LineWrite_BSK(idaPRINT_SLIP_LINE);
                    }

                    if (pOutput_Type == "PREVIEW")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PreView(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PRINTER")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.Printing(1, vPageNumber); 
                    }
                    else if (pOutput_Type == "PDF")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PDF(pSaveFileName);

                    }
                    else if (pOutput_Type == "EXCEL")
                    {
                        ////[SAVE]
                        xlPrinting.Save(pSaveFileName); //저장 파일명
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

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_SEK(string pOutput_Type, string pSaveFileName)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            vMessageText = string.Format("Printing Starting", vPageTotal);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0212_021.xls";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    object vObject;
                    int vCountRow = 0;
                     
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    //회계법인명.
                    IDC_GET_COMPANY_NAME_P.ExecuteNonQuery();
                    object vSOB_DESC = IDC_GET_COMPANY_NAME_P.GetCommandParamValue("O_SOB_DESC");

                    xlPrinting.HeaderWrite_SEK(idaSLIP_HEADER, vLOCAL_DATE);
                    vObject = HEADER_ID.EditValue;

                    idaPRINT_SLIP_LINE.SetSelectParamValue("W_HEADER_ID", vObject);
                    idaPRINT_SLIP_LINE.Fill();

                    vCountRow = idaPRINT_SLIP_LINE.CurrentRows.Count;
                    if (vCountRow > 0)
                    {
                        vPageNumber = xlPrinting.LineWrite_SEK(idaPRINT_SLIP_LINE);
                    }

                    if (pOutput_Type == "PREVIEW")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PreView(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PRINTER")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.Printing(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PDF")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PDF(pSaveFileName);

                    }
                    else if (pOutput_Type == "EXCEL")
                    {
                        ////[SAVE]
                        xlPrinting.Save(pSaveFileName); //저장 파일명
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

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

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
                     
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0212_Load(object sender, EventArgs e)
        {             
            // 회계장부 정보 설정.
            GetAccountBook();
        }

        private void FCMF0212_Shown(object sender, EventArgs e)
        {
            RB_PRINTER.CheckedState = ISUtil.Enum.CheckedState.Checked;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            BTN_OK.Enabled = true;
            BTN_CLOSED.Enabled = true;
            if (mPrintOptionFlag == "Y")
            {
                this.Width = 310;
                this.Height = 220;

                GB_PRINT_OPTION.Visible = true;
                BTN_OK.Visible = true;
                BTN_CLOSED.Visible = true;
            }
            else
            { 
                this.Width = 310;
                this.Height = 105;

                GB_PRINT_OPTION.Visible = false;
                BTN_OK.Visible = false;
                BTN_CLOSED.Visible = false;

                XLPrinting_Main(iString.ISNull(V_PRINT_TYPE.EditValue));

                this.DialogResult = DialogResult.OK;
                this.Close();
                return;
            } 
        }

        private void RB_PRINT_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;
            V_PRINT_TYPE.EditValue = vRadio.RadioCheckedString; 
        }
             
        #endregion

        private void BTN_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_PRINT_TYPE.EditValue) == string.Empty)
            { 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10327"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            BTN_OK.Enabled = false;
            BTN_CLOSED.Enabled = false;

            XLPrinting_Main(iString.ISNull(V_PRINT_TYPE.EditValue));

            BTN_CLOSED.Enabled = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #region ----- Lookup Event ----- 
         
        #endregion       

        #region ----- Adapter Event -----
         
        #endregion

    }
}