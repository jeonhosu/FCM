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

namespace FCMF0629
{
    public partial class FCMF0629 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ----- Constructor -----

        public FCMF0629()
        {
            InitializeComponent();
        }

        public FCMF0629(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Set_Default_Value()
        {
            // Budget Select Type.
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "BUDGET_CAPACITY");
            idcDEFAULT_VALUE.ExecuteNonQuery();

            //APPROVE_STATUS_0.EditValue =idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
            //APPROVE_STATUS_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
        }

        private void SearchDB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_DETAIL.TabIndex)
            {
                SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
            }
            else
            {
                if (iString.ISNull(W_BUDGET_PERIOD_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_BUDGET_PERIOD_FR.Focus();
                    return;
                }
                IDA_BUDGET_APPLY_HEADER.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", -1);
                IDA_BUDGET_APPLY_HEADER.Fill();

                IDA_BUDGET_APPLY_APPR_LIST.Fill();
                Set_Total_Amount(); 
            }
        }

        private void SearchDB_DTL(object pBUDGET_APPLY_HEADER_ID)
        {
            if (iString.ISNull(pBUDGET_APPLY_HEADER_ID) != string.Empty)
            {
                TB_MAIN.SelectedIndex = 1;
                TB_MAIN.SelectedTab.Focus();

                IDA_BUDGET_APPLY_HEADER.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", pBUDGET_APPLY_HEADER_ID); 
                try
                {
                    IDA_BUDGET_APPLY_HEADER.Fill();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }
                BUDGET_PERIOD.Focus();    
            }
        }
          
        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ILD_COMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter_W(object pGroupCode, object pWhere, object pEnabled_YN)
        {
            ILD_COMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ILD_COMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ILD_COMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Total_Amount()
        {
            decimal vBudget_Amount = 0;
            decimal vApply_Sum_Amount = 0;
            decimal vApply_Amount = 0;
            decimal vRemain_Amount = 0;
            object vAmount;
            int vIDX_BUDGET_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("BUDGET_AMOUNT");
            int vIDX_APPLY_SUM_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPLY_SUM_AMOUNT");
            int vIDX_APPLY_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPLY_AMOUNT");
            int vIDX_REMAIN_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("REMAIN_AMOUNT");

            for (int r = 0; r < IGR_BUDGET_APPLY_LINE.RowCount; r++)
            {
                //기초예산액.
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_BUDGET_AMOUNT);
                vBudget_Amount = vBudget_Amount + iString.ISDecimaltoZero(vAmount);

                //누적금액
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_APPLY_SUM_AMOUNT);
                vApply_Sum_Amount = vApply_Sum_Amount + iString.ISDecimaltoZero(vAmount);

                //신청금액
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_APPLY_AMOUNT);
                vApply_Amount = vApply_Amount + iString.ISDecimaltoZero(vAmount);

                //잔액
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_REMAIN_AMOUNT);
                vRemain_Amount = vRemain_Amount + iString.ISDecimaltoZero(vAmount);
            }
            V_BUDGET_AMOUNT.EditValue = vBudget_Amount;
            V_APPLY_SUM_AMOUNT.EditValue = vApply_Sum_Amount;
            V_APPLY_AMOUNT.EditValue = vApply_Amount;
            V_REMAIN_AMOUNT.EditValue = vRemain_Amount;
        } 

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            int vIDX_APPROVE_STATUS = pGrid.GetColumnToIndex("APPROVE_STATUS");
            object vAPPROVE_STATUS = string.Empty;
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                vAPPROVE_STATUS = pGrid.GetCellValue(i, vIDX_APPROVE_STATUS);
                if (iString.ISNull(W_APPROVE_STATUS.EditValue) != string.Empty)
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
                }
                else
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, "N");
                }
            }

            pGrid.LastConfirmChanges();
            IDA_BUDGET_APPLY_APPR_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_APPR_LIST.Refillable = true;
        }

        private bool Check_Added()
        {
            Boolean Row_Added_Status = false;
            
            //헤더 체크 
            for (int r = 0; r < IDA_BUDGET_APPLY_HEADER.SelectRows.Count; r++)
            {
                if (IDA_BUDGET_APPLY_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_BUDGET_APPLY_HEADER.SelectRows[r].RowState == DataRowState.Modified)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10169"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            //헤더 변경없으면 라인 체크 
            if (Row_Added_Status == false)
            {
                for (int r = 0; r < IDA_BUDGET_APPLY_LINE.SelectRows.Count; r++)
                {
                    if (IDA_BUDGET_APPLY_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_BUDGET_APPLY_LINE.SelectRows[r].RowState == DataRowState.Modified)
                    {
                        Row_Added_Status = true;
                    }
                }
                if (Row_Added_Status == true)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10169"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            return (Row_Added_Status);
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
                    if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Left = 278;
                        GB_RETURN.Top = 89;

                        GB_RETURN.Width = 600;
                        GB_RETURN.Height = 200;

                        GB_RETURN.Border3DStyle = Border3DStyle.Bump;
                        GB_RETURN.BorderStyle = BorderStyle.Fixed3D;

                        //값 초기화.
                        RETURN_REMARK.EditValue = string.Empty;

                        GB_RETURN.Visible = true;
                    }

                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                TB_MAIN.Enabled = false;
            }
            else
            {
                try
                {
                    if (pSub_Panel == "ALL")
                    {
                        GB_RETURN.Visible = false;
                    }
                    else if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Visible = false;
                    }

                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }

                TB_MAIN.Enabled = true;
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void Show_Detail(object pPERIOD_NAME, object pSLIP_NUM
                                , object pBUDGET_DEPT_NAME, object pBUDGET_DEPT_ID)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            FCMF0629_DETAIL vFCMF0629_DETAIL = new FCMF0629_DETAIL(isAppInterfaceAdv1.AppInterface, pPERIOD_NAME, pSLIP_NUM
                                                                , pBUDGET_DEPT_NAME, pBUDGET_DEPT_ID);

            dlgRESULT = vFCMF0629_DETAIL.ShowDialog();
            vFCMF0629_DETAIL.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
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


        #region ----- XL Print Methods ----

        private void XLPrinting_Main(string pOutput_Type)
        {
            object vBUDGET_PERIOD;
            if(TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                vBUDGET_PERIOD = IGR_BUDGET_APPLY_LIST.GetCellValue("BUDGET_PERIOD");
            }
            else
            {
                vBUDGET_PERIOD = BUDGET_PERIOD.EditValue;
            }
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(vBUDGET_PERIOD));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0629");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            XLPrinting(pOutput_Type);             
        }

        private void XLPrinting(string pOutput_Type)
        {
            //string vDefaultPrinter = GetDefaultPrinter();
            //PD.PrinterSettings = PS;
            //if (PD.ShowDialog() == DialogResult.OK)
            //{
            //    SetDefaultPrinter(PD.PrinterSettings.PrinterName);
            //} 

            string vMessageText = string.Empty;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            object vBUDGET_APPLY_HEADER_ID = BUDGET_APPLY_HEADER_ID.EditValue;
            if (TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                vBUDGET_APPLY_HEADER_ID = IGR_BUDGET_APPLY_LIST.GetCellValue("BUDGET_APPLY_HEADER_ID");
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1); 

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0629_001.xlsx"; 
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    //회계법인명.
                    IDC_GET_COMPANY_NAME_P.ExecuteNonQuery();
                    object vSOB_DESC = IDC_GET_COMPANY_NAME_P.GetCommandParamValue("O_SOB_DESC");

                    IDA_PRINT_BUDGET_APPLY_HEADER.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", vBUDGET_APPLY_HEADER_ID);
                    IDA_PRINT_BUDGET_APPLY_HEADER.Fill();

                    IDA_PRINT_BUDGET_APPLY_LINE.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", IDA_PRINT_BUDGET_APPLY_HEADER.CurrentRow["BUDGET_APPLY_HEADER_ID"]);
                    IDA_PRINT_BUDGET_APPLY_LINE.Fill();

                    IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_TYPE", IDA_PRINT_BUDGET_APPLY_HEADER.CurrentRow["BUDGET_TYPE"]);
                    IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_HEADER_ID", IDA_PRINT_BUDGET_APPLY_HEADER.CurrentRow["BUDGET_APPLY_HEADER_ID"]);
                    IDA_PRINT_APPROVAL_STEP_PERSON.Fill();

                    vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_BUDGET_APPLY_HEADER, IDA_PRINT_BUDGET_APPLY_LINE, IDA_PRINT_APPROVAL_STEP_PERSON, vSOB_DESC, vLOCAL_DATE);

                    if (pOutput_Type == "PRINT")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.Printing(1, vPageNumber);

                    }
                    else if (pOutput_Type == "FILE")
                    {
                        ////[SAVE]
                        xlPrinting.SAVE("Budget_Req_"); //저장 파일명
                    }
                     
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }

                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                //SetDefaultPrinter(vDefaultPrinter);

                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            //SetDefaultPrinter(vDefaultPrinter);
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
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
                    if (IDA_BUDGET_APPLY_LINE.IsFocused)
                    {
                        IDA_BUDGET_APPLY_LINE.Cancel();
                    }
                    else
                    {
                        IDA_BUDGET_APPLY_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_APPLY_LINE.CurrentRow.RowState == DataRowState.Added)
                    {
                        IDA_BUDGET_APPLY_LINE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0629_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_APPLY_APPR_LIST.FillSchema();
            IDA_BUDGET_APPLY_HEADER.FillSchema();
            IDA_BUDGET_APPLY_LINE.FillSchema();

            //서브판넬 
            Init_Sub_Panel(false, "ALL");
        }

        private void FCMF0629_Shown(object sender, EventArgs e)
        {
            W_BUDGET_PERIOD_FR.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_BUDGET_PERIOD_TO.EditValue = iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 1)); 
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;

            BTN_CHG_APPROVAL_STEP.BringToFront();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        private void irbALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRADIO = sender as ISRadioButtonAdv;
            W_APPROVE_STATUS.EditValue = vRADIO.RadioButtonValue;

            //버튼제어 및 체크박스 제어.
            if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "N")
            {
                BTN_APPROVAL.Enabled = true;
                BTN_CHG_APPROVAL_STEP.Enabled = true;
                BTN_CANCEL_APPROVAL.Enabled = false;
                BTN_RETURN.Enabled = true;
            }
            else if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "Y")
            {
                BTN_APPROVAL.Enabled = false;
                BTN_CHG_APPROVAL_STEP.Enabled = false;
                BTN_CANCEL_APPROVAL.Enabled = true;
                BTN_RETURN.Enabled = false;
            }
            else
            {
                BTN_APPROVAL.Enabled = false;
                BTN_CANCEL_APPROVAL.Enabled = false;
                BTN_CHG_APPROVAL_STEP.Enabled = false;
                BTN_RETURN.Enabled = false;
            }
            SearchDB();
        }
         
        private void IGR_BUDGET_ADD_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_BUDGET_APPLY_LIST.RowCount > 0)
            {
                SearchDB_DTL(IGR_BUDGET_APPLY_LIST.GetCellValue("BUDGET_APPLY_HEADER_ID"));
            }
        }
         
        private void BTN_REQ_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue,0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_BUDGET_APPLY_APPR.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_APPLY_APPR.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_APPLY_APPR.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            
            if (IDC_EXEC_BUDGET_APPLY_APPR.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        private void BTN_CANCEL_REQ_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_BUDGET_APPLY_APPR.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_BUDGET_APPLY_APPR.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_BUDGET_APPLY_APPR.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_BUDGET_APPLY_APPR.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        private void BTN_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }
            
            //서브판넬 
            Init_Sub_Panel(true, "RETURN");
        }

        private void C_BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "RETURN");
        }

        private void C_BTN_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(RETURN_REMARK.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(RETURN_REMARK))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_BUDGET_APPLY_RETURN.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_APPLY_RETURN.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_APPLY_RETURN.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_EXEC_BUDGET_APPLY_RETURN.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        private void BTN_CHG_APPROVAL_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            DialogResult dlgResult = DialogResult.None;
            FCMF0629_APPR_STEP vFCMF0629_APPR_STEP = new FCMF0629_APPR_STEP(isAppInterfaceAdv1.AppInterface,
                                                                            BUDGET_APPLY_NUM.EditValue, APPROVAL_STEP_SEQ.EditValue,
                                                                            BUDGET_PERIOD.EditValue, BUDGET_APPLY_HEADER_ID.EditValue,
                                                                            BUDGET_TYPE_NAME.EditValue, BUDGET_TYPE.EditValue, 
                                                                            BUDGET_DEPT_NAME.EditValue, BUDGET_DEPT_CODE.EditValue, BUDGET_DEPT_ID.EditValue);
            dlgResult = vFCMF0629_APPR_STEP.ShowDialog();
            Application.DoEvents();

            vFCMF0629_APPR_STEP.Dispose();
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        private void IGR_BUDGET_APPLY_SLIP_CellDoubleClick(object pSender)
        {
            if (IGR_BUDGET_APPLY_SLIP.Row > 0)
            {
                Show_Detail(BUDGET_PERIOD.EditValue, IGR_BUDGET_APPLY_SLIP.GetCellValue("SLIP_NUM")
                            , BUDGET_DEPT_NAME.EditValue, BUDGET_DEPT_ID.EditValue);
            }
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ILA_PERIOD_FR_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_PERIOD_TO_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", W_BUDGET_PERIOD_FR.EditValue);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_PERIOD_NAME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_PERIOD_NAME_COPY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ilaPERIOD_NAME_0_SelectedRowData(object pSender)
        {
            SearchDB();
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(W_BUDGET_PERIOD_FR.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(W_BUDGET_PERIOD_TO.EditValue));
        }

        private void ilaBUDGET_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'ADD'", "Y");
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUDGET_CAPACITY", DBNull.Value, "Y");
        }

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_CONTROL_YN", "N");
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBUDGET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'ADD'", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ILA_DEPT_COPY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }


        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_CONTROL_YN", "N");
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCAUSE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_CAUSE", "Value1 = 'ADD'", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_BUDGET_ADD_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager == null)
            {
                V_BUDGET_AMOUNT.EditValue = 0;
                V_APPLY_SUM_AMOUNT.EditValue = 0;
                V_APPLY_AMOUNT.EditValue = 0;
                V_REMAIN_AMOUNT.EditValue = 0;
            }
            Set_Total_Amount();
        }

        private void IDA_BUDGET_ADD_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_PERIOD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BUDGET_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REMARK))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_BUDGET_ADD_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["AMOUNT"],0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10537"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        } 

        #endregion


    }
}