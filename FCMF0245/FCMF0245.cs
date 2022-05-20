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

namespace FCMF0245
{
    public partial class FCMF0245 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private bool mSave_Appr_Status = false;
        bool mSUB_SHOW_FLAG = false;
        private string vAPPROVAL_PERSON_YN = "N";

        bool mIsClickInquiryDetail = false;
        int mInquiryDetailPreX, mInquiryDetailPreY; //마우스 이동 제어.

        #endregion;

        #region ----- Constructor -----

        public FCMF0245()
        {
            InitializeComponent();
        }

        public FCMF0245(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public FCMF0245(Form pMainForm, ISAppInterface pAppInterface, object pJOB_NO)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            
            

        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            IDA_CLOSE_SLIP_SUMMARY.Fill();
            IDA_CLOSE_SLIP_ACCOUNT.Fill();
            IDA_CLOSE_SLIP_MONTHLY.Fill();
            IDA_CLOSE_SLIP_LIST.Fill();


            IGR_APPROVAL_PERIOD.LastConfirmChanges();
            IDA_APPROVAL_PERIOD.OraSelectData.AcceptChanges();
            IDA_APPROVAL_PERIOD.Refillable = true;
            IDA_APPROVAL_PERIOD.Fill();
        }

        private void Init_Sub_Panel(bool pShow_Flag, string pSub_Panel)
        {
            if (mSUB_SHOW_FLAG == true && pShow_Flag == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pShow_Flag == true)
            {
                try
                {   
                    if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Left = 65;
                        GB_APPR.Top = 115;

                        GB_APPR.Width = 900;
                        GB_APPR.Height = 240;

                        GB_APPR.Border3DStyle = Border3DStyle.Bump;
                        GB_APPR.BorderStyle = BorderStyle.Fixed3D;

                        //GroupBox 이동//
                        GB_APPR.Controls[0].MouseDown += GB_APPR_MouseDown;
                        GB_APPR.Controls[0].MouseMove += GB_APPR_MouseMove;
                        GB_APPR.Controls[0].MouseUp += GB_APPR_MouseUp;
                        GB_APPR.Controls[1].MouseDown += GB_APPR_MouseDown;
                        GB_APPR.Controls[1].MouseMove += GB_APPR_MouseMove;
                        GB_APPR.Controls[1].MouseUp += GB_APPR_MouseUp;

                        GB_APPR.BringToFront();
                        GB_APPR.Visible = true;
                    }
                    else if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Left = 278;
                        GB_RETURN.Top = 89;

                        GB_RETURN.Width = 600;
                        GB_RETURN.Height = 200;

                        GB_RETURN.Border3DStyle = Border3DStyle.Bump;
                        GB_RETURN.BorderStyle = BorderStyle.Fixed3D;

                        //GroupBox 이동// 
                        GB_RETURN.Controls[0].MouseDown += GB_RETURN_MouseDown;
                        GB_RETURN.Controls[0].MouseMove += GB_RETURN_MouseMove;
                        GB_RETURN.Controls[0].MouseUp += GB_RETURN_MouseUp;
                        GB_RETURN.Controls[1].MouseDown += GB_RETURN_MouseDown;
                        GB_RETURN.Controls[1].MouseMove += GB_RETURN_MouseMove;
                        GB_RETURN.Controls[1].MouseUp += GB_RETURN_MouseUp;

                        //값 초기화.
                        V_RETURN_REMARK.EditValue = string.Empty;
                        GB_RETURN.BringToFront();
                        GB_RETURN.Visible = true;
                    }
                    else if (pSub_Panel == "APPROVAL")
                    {
                        GB_APPROVAL.Left = 278;
                        GB_APPROVAL.Top = 89;

                        GB_APPROVAL.Width = 600;
                        GB_APPROVAL.Height = 200;

                        GB_APPROVAL.Border3DStyle = Border3DStyle.Bump;
                        GB_APPROVAL.BorderStyle = BorderStyle.Fixed3D;

                        //GroupBox 이동// 
                        GB_APPROVAL.Controls[0].MouseDown += GB_APPROVAL_MouseDown;
                        GB_APPROVAL.Controls[0].MouseMove += GB_APPROVAL_MouseMove;
                        GB_APPROVAL.Controls[0].MouseUp += GB_APPROVAL_MouseUp;
                        GB_APPROVAL.Controls[1].MouseDown += GB_APPROVAL_MouseDown;
                        GB_APPROVAL.Controls[1].MouseMove += GB_APPROVAL_MouseMove;
                        GB_APPROVAL.Controls[1].MouseUp += GB_APPROVAL_MouseUp;

                        //값 초기화.
                        V_APPROVAL_DESCRIPTION.EditValue = string.Empty;
                        GB_APPROVAL.BringToFront();
                        GB_APPROVAL.Visible = true;
                    }
                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                } 

                TB_MAIN.Enabled = false;
                IGB_INQUIRY_CONDITION.Enabled = false; 
                GB_APPR_STEP.Enabled = false;
            }
            else
            {
                try
                {
                    if (pSub_Panel == "ALL")
                    { 
                        GB_APPR_STEP.Enabled = false;
                        GB_APPR.Visible = false;
                        GB_RETURN.Visible = false;
                        GB_APPROVAL.Visible = false;
                    } 
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Visible = false;
                    } 
                    else if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Visible = false;
                    }
                    else if (pSub_Panel == "APPROVAL")
                    {
                        GB_APPROVAL.Visible = false;
                    }
                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                TB_MAIN.Enabled = true;
                IGB_INQUIRY_CONDITION.Enabled = true;
                GB_APPR_STEP.Enabled = true;
                GB_APPR_STEP.Enabled = true;
            }
        }

        private void Init_Approval_Person()
        {
            IDC_GET_APPROVAL_PERSON_YN.ExecuteNonQuery();
            vAPPROVAL_PERSON_YN = iConv.ISNull(IDC_GET_APPROVAL_PERSON_YN.GetCommandParamValue("O_APPROVAL_PERSON_YN"), "N");

            if (vAPPROVAL_PERSON_YN.Equals("Y"))
            {                
                igbCONFIRM_STATUS.Visible = true;
                btnCONFIRM_YES.Visible = true;
                btnCONFIRM_CANCEL.Visible = true;
                btnCONFIRM_RETURN.Visible = true; 
            }
            else
            {
                igbCONFIRM_STATUS.Visible = false;

                btnCONFIRM_YES.Visible = false;
                btnCONFIRM_CANCEL.Visible = false;
                btnCONFIRM_RETURN.Visible = false; 
            } 
        }


        private void Init_Approval_BTN()
        {            
            if (iConv.ISNull(V_CONFIRM_STATUS.EditValue).Equals("Y"))
            {
                BTN_INIT_APPR_STEP.Enabled = true;
                btnCONFIRM_YES.Enabled = false;
                btnCONFIRM_CANCEL.Enabled = true;
                btnCONFIRM_RETURN.Enabled = true;
            }
            else if (iConv.ISNull(V_CONFIRM_STATUS.EditValue).Equals("N"))
            {
                BTN_INIT_APPR_STEP.Enabled = true;
                btnCONFIRM_YES.Enabled = true;
                btnCONFIRM_CANCEL.Enabled = false;
                btnCONFIRM_RETURN.Enabled = true;
            } 
            else
            {
                btnCONFIRM_YES.Enabled = true;
                btnCONFIRM_CANCEL.Enabled = false;
                btnCONFIRM_RETURN.Enabled = false;
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


        #region ---- Doc Att / Appr Step ----

        private void Init_SLIP_APPR()
        {
            S_PERIOD_NAME.EditValue = W_PERIOD_NAME.EditValue;
            S_DEPT_NAME.EditValue = W_DEPT_NAME.EditValue;
            S_DEPT_ID.EditValue = W_DEPT_ID.EditValue;

            Init_Sub_Panel(true, "APPR_STEP");

            IDA_APPROVAL_PERIOD.Cancel();
            IDA_APPROVAL_PERIOD.Fill();
            if (IDA_APPROVAL_PERIOD.CurrentRows.Count > 0)
            { 
                return;
            }

            IDA_INIT_APPROVAL_LINE.SetSelectParamValue("W_PERIOD_NAME", W_PERIOD_NAME.EditValue);
            IDA_INIT_APPROVAL_LINE.SetSelectParamValue("W_DEPT_ID", W_DEPT_ID.EditValue);
            IDA_INIT_APPROVAL_LINE.Fill();
            foreach (DataRow row in IDA_INIT_APPROVAL_LINE.SelectRows)
            {
                IDA_APPROVAL_PERIOD.AddUnder();

                IGR_APPROVAL_PERIOD.SetCellValue("APPROVAL_PERIOD_ID", row["APPROVAL_PERIOD_ID"]);
                IGR_APPROVAL_PERIOD.SetCellValue("PERIOD_NAME", row["PERIOD_NAME"]);
                IGR_APPROVAL_PERIOD.SetCellValue("DEPT_ID", row["DEPT_ID"]);
                IGR_APPROVAL_PERIOD.SetCellValue("APPROVAL_STEP_SEQ", row["APPROVAL_STEP_SEQ"]);
                IGR_APPROVAL_PERIOD.SetCellValue("APPROVAL_STEP_ID", row["APPROVAL_STEP_ID"]);
                IGR_APPROVAL_PERIOD.SetCellValue("APPROVAL_STEP", row["APPROVAL_STEP"]);
                IGR_APPROVAL_PERIOD.SetCellValue("APPROVAL_STEP_NAME", row["APPROVAL_STEP_NAME"]);
                IGR_APPROVAL_PERIOD.SetCellValue("PERSON_ID", row["PERSON_ID"]);
                IGR_APPROVAL_PERIOD.SetCellValue("PERSON_NUM", row["PERSON_NUM"]);
                IGR_APPROVAL_PERIOD.SetCellValue("PERSON_NAME", row["PERSON_NAME"]);
                IGR_APPROVAL_PERIOD.SetCellValue("EMAIL", row["EMAIL"]);
                IGR_APPROVAL_PERIOD.SetCellValue("DESCRIPTION", row["DESCRIPTION"]);
                IGR_APPROVAL_PERIOD.SetCellValue("REQUIRED_FLAG", row["REQUIRED_FLAG"]);
                IGR_APPROVAL_PERIOD.SetCellValue("APPR_FLAG", row["APPR_FLAG"]);
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting("PRINT", ISG_SLIP_SUMMARY, ISG_SLIP_ACCOUNT, ISG_SLIP_MONTHLY, ISG_SLIP_LIST);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting("EXCEL", ISG_SLIP_SUMMARY, ISG_SLIP_ACCOUNT, ISG_SLIP_MONTHLY, ISG_SLIP_LIST);
                }
            }
        }

        #endregion;

        #region ----- Excel Export -----
        private void ExcelExport(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            GridExcelConverterControl vExport = new GridExcelConverterControl();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog.DefaultExt = ".xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ////데이터 테이블을 이용한 export
                //Syncfusion.XlsIO.ExcelEngine vEng = new Syncfusion.XlsIO.ExcelEngine();
                //Syncfusion.XlsIO.IApplication vApp = vEng.Excel;
                //string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();
                //if (vFileExtension == "XLSX")
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel2007;
                //}
                //else
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel97to2003;
                //}
                //Syncfusion.XlsIO.IWorkbook vWorkbook = vApp.Workbooks.Create(1);
                //Syncfusion.XlsIO.IWorksheet vSheet = vWorkbook.Worksheets[0];
                //foreach(System.Data.DataRow vRow in IDA_MATERIAL_LIST_ALL.CurrentRows)
                //{
                //    vSheet.ImportDataTable(vRow.Table, true, 1, 1, -1, -1);
                //}
                //vWorkbook.SaveAs(saveFileDialog.FileName);
                vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);
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

        #region ----- Printing -----

        private void XLPrinting(string pOutput_Type
                                       , InfoSummit.Win.ControlAdv.ISGridAdvEx pCLOSE_SLIP_SUMMARY
                                       , InfoSummit.Win.ControlAdv.ISGridAdvEx pCLOSE_SLIP_ACCOUNT
                                       , InfoSummit.Win.ControlAdv.ISGridAdvEx pCLOSE_SLIP_MONTHLY
                                       , InfoSummit.Win.ControlAdv.ISGridAdvEx pCLOSE_SLIP_LIST
                                       )
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();
            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;
            int vCountRowGrid = pCLOSE_SLIP_SUMMARY.RowCount;
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
                    xlPrinting.OpenFileNameExcel = "FCMF0245_001.xlsx";
                    //-------------------------------------------------------------------------------------
                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------
                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        int vCountRow = 0;
                        string vPerson_Name = string.Empty;
                        string vDepartment = string.Empty;
                        string vPeriod = string.Empty;
                        DateTime vCurrent_date = DateTime.Now;


                        IDA_PRINT_APPROVAL_PERSON.Fill(); 

                        vCountRow = IDA_CLOSE_SLIP_SUMMARY.CurrentRows.Count;
                        vPerson_Name = iConv.ISNull(DISPLAY_PERSON.EditValue);
                        vDepartment = iConv.ISNull(W_DEPT_NAME.EditValue);
                        vPeriod = iConv.ISNull(W_PERIOD_NAME.EditValue);


                        if (vCountRow > 0)
                        {
                            vPageNumber = xlPrinting.MainWrite(IDA_CLOSE_SLIP_SUMMARY, IDA_CLOSE_SLIP_ACCOUNT
                                                             , IDA_CLOSE_SLIP_MONTHLY, IDA_CLOSE_SLIP_LIST
                                                             , IDA_PRINT_APPROVAL_PERSON
                                                             , vPerson_Name 
                                                             , vDepartment
                                                             , vCurrent_date
                                                             , vPeriod);
                        }

                        if (pOutput_Type == "PRINT")
                        {
                            //[PRINT]
                            //시작 페이지 번호, 종료 페이지 번호 
                            xlPrinting.PrintPreview(1, vPageNumber);

                            //xlPrinting.Printing(1, vPageNumber);
                        }

                        else if (pOutput_Type == "EXCEL")
                        {
                            ////[SAVE]
                            xlPrinting.Save("Dept_Monthly_Closed_"); //저장 파일명
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
            vMessageText = string.Format("Print End ^.^ [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion
        
        private void FCMF0245_Load(object sender, EventArgs e)
        {
            object vFCM_YN = string.Empty;
             
            IDC_GET_PERSON_DEPT_P.ExecuteNonQuery(); 
            W_DEPT_NAME.EditValue = IDC_GET_PERSON_DEPT_P.GetCommandParamValue("O_DEPT_NAME");
            W_DEPT_ID.EditValue = IDC_GET_PERSON_DEPT_P.GetCommandParamValue("O_DEPT_ID"); 
            DISPLAY_PERSON.EditValue = isAppInterfaceAdv1.AppInterface.DisplayName;

            igbCONFIRM_STATUS.BringToFront();
            APPROVAL_STEP_SEQ.BringToFront();

            irbCONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_CONFIRM_STATUS.EditValue = irbCONFIRM_ALL.RadioCheckedString;

            Init_Approval_Person();

            //IDC_FCM_DEPT.ExecuteNonQuery();
            //vFCM_YN = IDC_FCM_DEPT.GetCommandParamValue("O_FCM_DEPT_YN");

            //if (iConv.ISNull(vFCM_YN) == "Y")
            //    W_DEPT_NAME.ReadOnly = false;    
            //else
            //    W_DEPT_NAME.ReadOnly = true; 

            //서브판넬 
            Init_Sub_Panel(false, "ALL");

            APPROVAL_STEP_SEQ.BringToFront(); 
        }

        private void FCMF0245_Shown(object sender, EventArgs e)
        {
            IDA_APPROVAL_PERIOD.FillSchema();
        }

        private void BTN_APPR_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            S_PERIOD_NAME.EditValue = W_PERIOD_NAME.EditValue;
            S_DEPT_NAME.EditValue = W_DEPT_NAME.EditValue;
            S_DEPT_ID.EditValue = W_DEPT_ID.EditValue;

            IDA_APPROVAL_PERIOD.Fill();  
            Init_Sub_Panel(true, "APPR_STEP");
        }
        private void irbCONFIRM_Status_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;

            V_CONFIRM_STATUS.EditValue = iStatus.RadioCheckedString;

            Init_Approval_BTN();
        }


        private void IDA_APPROVAL_PERSON_UpdateCompleted(object pSender)
        {
            if (IDA_APPROVAL_PERIOD.UpdateModifiedRowCount != 0)
            {
                mSave_Appr_Status = true;
            }
        }

        private void BTN_INSERT_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERIOD.AddUnder();
            IGR_APPROVAL_PERIOD.SetCellValue("PERIOD_NAME", S_PERIOD_NAME.EditValue);
            IGR_APPROVAL_PERIOD.SetCellValue("DEPT_ID", S_DEPT_ID.EditValue);

            IGR_APPROVAL_PERIOD.CurrentCellMoveTo(IGR_APPROVAL_PERIOD.RowIndex, IGR_APPROVAL_PERIOD.GetColumnToIndex("APPROVAL_STEP_SEQ"));
            IGR_APPROVAL_PERIOD.CurrentCellActivate(IGR_APPROVAL_PERIOD.RowIndex, IGR_APPROVAL_PERIOD.GetColumnToIndex("APPROVAL_STEP_SEQ"));
            IGR_APPROVAL_PERIOD.Focus();
        }

        private void BTN_CANCEL_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERIOD.Cancel();
        }

        private void BTN_DELETE_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERIOD.Delete();
        }

        private void BTN_CLOSED_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            mSave_Appr_Status = true;
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) != string.Empty)
            {
                foreach (DataRow vRow in IDA_APPROVAL_PERIOD.CurrentRows)
                {
                    if (vRow.RowState != DataRowState.Unchanged)
                    {
                        mSave_Appr_Status = false;
                    }
                }

                if (mSave_Appr_Status == false)
                {
                    try
                    {
                        IDA_APPROVAL_PERIOD.Update();
                    }
                    catch
                    {
                        return;
                    }
                    Init_Sub_Panel(false, "APPR_STEP");
                }
                else
                {
                    Init_Sub_Panel(false, "APPR_STEP");
                }
            }
            Init_Sub_Panel(false, "APPR_STEP");
        }


        #region ----- Form Event ------

        #endregion

        #region ----- Lookup Event ------

        #endregion

        private void ILA_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_APPROVAL_PERSON_YN", vAPPROVAL_PERSON_YN);
        }

        private void GB_APPR_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_APPR_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_APPR_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_APPR.Location;
                I.Offset(gx, gy);
                GB_APPR.Location = I;
            }
        }

        private void GB_RETURN_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_RETURN_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_RETURN_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_RETURN.Location;
                I.Offset(gx, gy);
                GB_RETURN.Location = I;
            }
        }

        private void GB_APPROVAL_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_APPROVAL_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_APPROVAL.Location;
                I.Offset(gx, gy);
                GB_APPROVAL.Location = I;
            }
        }

        private void GB_APPROVAL_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }


        private void BTN_INIT_APPR_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                return;
            }

            //승인단계 설정. 
            Init_SLIP_APPR();
        }

        private void btnCONFIRM_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IDA_CLOSE_SLIP_SUMMARY.CurrentRows.Count <= 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10054"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
             
            R_PERIOD_NAME.EditValue = W_PERIOD_NAME.EditValue;
            R_DEPT_NAME.EditValue = W_DEPT_NAME.EditValue;
            R_DEPT_ID.EditValue = W_DEPT_ID.EditValue;

            IDC_GET_APPROVAL_PERIOD_SEQ.ExecuteNonQuery();

            //서브판넬 
            Init_Sub_Panel(true, "RETURN");
        }

        private void btnCONFIRM_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IDA_CLOSE_SLIP_SUMMARY.CurrentRows.Count <= 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10054"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_GET_APPROVAL_PERIOD_SEQ.ExecuteNonQuery();
             
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_APPROVAL_PERIOD.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CANCEL_APPROVAL_PERIOD.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CANCEL_APPROVAL_PERIOD.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_APPROVAL_PERIOD.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            IGR_APPROVAL_PERIOD.LastConfirmChanges();
            IDA_APPROVAL_PERIOD.OraSelectData.AcceptChanges();
            IDA_APPROVAL_PERIOD.Refillable = true;

            IDA_APPROVAL_PERIOD.Fill();
        }

        private void C_BTN_EXEC_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_RETURN_APPROVAL_PERIOD.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_RETURN_APPROVAL_PERIOD.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_RETURN_APPROVAL_PERIOD.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_RETURN_APPROVAL_PERIOD.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //서브판넬 
            Init_Sub_Panel(false, "RETURN");

            IGR_APPROVAL_PERIOD.LastConfirmChanges();
            IDA_APPROVAL_PERIOD.OraSelectData.AcceptChanges();
            IDA_APPROVAL_PERIOD.Refillable = true;

            IDA_APPROVAL_PERIOD.Fill();
        }

        private void C_BTN_RETURN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "RETURN");
        }

        private void C_BTN_EXEC_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (iConv.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_APPROVAL_PERIOD.SetCommandParamValue("P_APPROVAL_DESCRIPTION", V_APPROVAL_DESCRIPTION.EditValue);
            IDC_EXEC_APPROVAL_PERIOD.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_EXEC_APPROVAL_PERIOD.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_EXEC_APPROVAL_PERIOD.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_EXEC_APPROVAL_PERIOD.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Init_Sub_Panel(false, "APPROVAL");

            IGR_APPROVAL_PERIOD.LastConfirmChanges();
            IDA_APPROVAL_PERIOD.OraSelectData.AcceptChanges();
            IDA_APPROVAL_PERIOD.Refillable = true;

            IDA_APPROVAL_PERIOD.Fill();
        }

        private void C_BTN_EXEC_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "APPROVAL");
        }

        private void ILA_PERIOD_SelectedRowData(object pSender)
        {
            Init_Approval_Person();
        }

        private void ILA_DEPT_SelectedRowData(object pSender)
        {
            Init_Approval_Person();
        }

        private void btnCONFIRM_YES_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IDA_CLOSE_SLIP_SUMMARY.CurrentRows.Count <= 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10054"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            A_PERIOD_NAME.EditValue = W_PERIOD_NAME.EditValue;
            A_DEPT_NAME.EditValue = W_DEPT_NAME.EditValue;
            A_DEPT_ID.EditValue = W_DEPT_ID.EditValue;

            IDC_GET_APPROVAL_PERIOD_SEQ.ExecuteNonQuery();
              
            //서브판넬 
            Init_Sub_Panel(true, "APPROVAL");

        }


    }
}