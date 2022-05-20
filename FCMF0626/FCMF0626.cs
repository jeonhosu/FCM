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

namespace FCMF0626
{
    public partial class FCMF0626 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ----- Constructor -----

        public FCMF0626()
        {
            InitializeComponent();
        }

        public FCMF0626(Form pMainForm, ISAppInterface pAppInterface)
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
                SearchDB_DTL(BUDGET_ADD_HEADER_ID.EditValue);
            }
            else
            {
                if (iString.ISNull(W_BUDGET_PERIOD_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_BUDGET_PERIOD_FR.Focus();
                    return;
                }
                IDA_BUDGET_ADD_HEADER.SetSelectParamValue("P_BUDGET_ADD_HEADER_ID", -1);
                IDA_BUDGET_ADD_HEADER.Fill();

                IDA_BUDGET_ADD_LIST.Fill();
                Set_Total_Amount();
                IGR_BUDGET_ADD_LINE.Focus();
            }
        }

        private void SearchDB_DTL(object pBUDGET_ADD_HEADER_ID)
        {
            if (iString.ISNull(pBUDGET_ADD_HEADER_ID) != string.Empty)
            {
                TB_MAIN.SelectedIndex = 1;
                TB_MAIN.SelectedTab.Focus();

                Set_Item_Status();
                Application.DoEvents();

                IDA_BUDGET_ADD_HEADER.SetSelectParamValue("P_BUDGET_ADD_HEADER_ID", pBUDGET_ADD_HEADER_ID); 
                try
                {
                    IDA_BUDGET_ADD_HEADER.Fill();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }
                BUDGET_PERIOD.Focus();    
            }
        }

        private void Budget_Add_Header_Insert()
        {
            IGR_BUDGET_ADD_LINE.SetCellValue("BUDGET_TYPE", W_BUDGET_TYPE.EditValue);
            IGR_BUDGET_ADD_LINE.SetCellValue("BUDGET_TYPE_NAME", W_BUDGET_TYPE_NAME.EditValue);
            IGR_BUDGET_ADD_LINE.SetCellValue("BUDGET_PERIOD", W_BUDGET_PERIOD_TO.EditValue);

            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus(); 

            BUDGET_PERIOD.Focus();
        }

        private void Budget_Add_Line_Insert()
        {
            //IGR_BUDGET_ADD_LINE.SetCellValue("EXPENDITURE_DATE", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
            IGR_BUDGET_ADD_LINE.SetCellValue("PL_AMOUNT", 0);
            IGR_BUDGET_ADD_LINE.SetCellValue("AMOUNT", 0);
            IGR_BUDGET_ADD_LINE.Focus();
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
            decimal vPlan_Amount = 0;
            decimal vThis_Amount = 0;
            decimal vGap_Amount = 0;
            object vAmount;
            int vIDX_PL_AMOUNT = IGR_BUDGET_ADD_LINE.GetColumnToIndex("PL_AMOUNT");
            int vIDX_AMOUNT = IGR_BUDGET_ADD_LINE.GetColumnToIndex("AMOUNT");
            int vIDX_GAP_AMOUNT = IGR_BUDGET_ADD_LINE.GetColumnToIndex("GAP_AMOUNT");
            for (int r = 0; r < IGR_BUDGET_ADD_LINE.RowCount; r++)
            {
                //사업계획.
                vAmount = 0;
                vAmount = IGR_BUDGET_ADD_LINE.GetCellValue(r, vIDX_PL_AMOUNT);
                vPlan_Amount = vPlan_Amount + iString.ISDecimaltoZero(vAmount);

                //당월
                vAmount = 0;
                vAmount = IGR_BUDGET_ADD_LINE.GetCellValue(r, vIDX_AMOUNT);
                vThis_Amount = vThis_Amount + iString.ISDecimaltoZero(vAmount);

                //차이금액
                vAmount = 0;
                vAmount = IGR_BUDGET_ADD_LINE.GetCellValue(r, vIDX_GAP_AMOUNT);
                vGap_Amount = vGap_Amount + iString.ISDecimaltoZero(vAmount);
            }
            V_PLAN_AMOUNT.EditValue = vPlan_Amount;
            V_THIS_AMOUNT.EditValue = vThis_Amount;
            V_GAP_AMOUNT.EditValue = vGap_Amount;
        }

        private void EXE_BUDGET_ADD_STATUS(object pPERIOD_NAME, object pAPPROVE_STATUS, object pAPPROVE_FLAG)
        {
            IDA_BUDGET_ADD_LIST.Update(); //수정사항 반영.

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_CHECK_YN = IGR_BUDGET_ADD_LINE.GetColumnToIndex("CHECK_YN");
            int vIDX_BUDGET_TYPE = IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_TYPE");
            int vIDX_BUDGET_PERIOD = IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_PERIOD");
            int vIDX_DEPT_ID = IGR_BUDGET_ADD_LINE.GetColumnToIndex("DEPT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_BUDGET_ADD_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_BUDGET_ADD_LINE.RowCount; i++)
            {
                if (iString.ISNull(IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_CHECK_YN), "N") == "Y")
                {
                    IGR_BUDGET_ADD_LINE.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    IGR_BUDGET_ADD_LINE.CurrentCellActivate(i, vIDX_CHECK_YN);

                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_BUDGET_TYPE));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_BUDGET_PERIOD));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_DEPT_ID));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("P_APPROVE_STATUS", pAPPROVE_STATUS);
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("P_APPROVE_FLAG", pAPPROVE_FLAG);
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("P_CHECK_YN", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_CHECK_YN));
                    idcBUDGET_ADD_STATUS.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(idcBUDGET_ADD_STATUS.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(idcBUDGET_ADD_STATUS.GetCommandParamValue("O_MESSAGE"));
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (idcBUDGET_ADD_STATUS.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }
            SearchDB();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void Set_Item_Status()
        {
            int mIDX_Col;

            if (C2_ALL_RECORD_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                REMARK.Insertable = false;
                REMARK.Updatable = false;

                //예상지급일.
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("EXPENDITURE_DATE");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 확정예산
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("PL_AMOUNT");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
                // 당월예산.
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("AMOUNT");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 비고.
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("DESCRIPTION");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            }
            else
            {
                REMARK.Insertable = true;
                REMARK.Updatable = true;

                //예상지급일.
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("EXPENDITURE_DATE");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 확정예산
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("PL_AMOUNT");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
                // 당월예산.
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("AMOUNT");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 비고.
                mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("DESCRIPTION");
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            }             
            IGR_BUDGET_ADD_LINE.ResetDraw = true;
        }

        private bool Check_Added()
        {
            Boolean Row_Added_Status = false;
            
            //헤더 체크 
            for (int r = 0; r < IDA_BUDGET_ADD_HEADER.SelectRows.Count; r++)
            {
                if (IDA_BUDGET_ADD_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_BUDGET_ADD_HEADER.SelectRows[r].RowState == DataRowState.Modified)
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
                for (int r = 0; r < IDA_BUDGET_ADD_LINE.SelectRows.Count; r++)
                {
                    if (IDA_BUDGET_ADD_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_BUDGET_ADD_LINE.SelectRows[r].RowState == DataRowState.Modified)
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
                    if (pSub_Panel == "COPY_BUDGET")
                    {
                        GB_COPY_DOCUMENT.Left = 180;
                        GB_COPY_DOCUMENT.Top = 95;

                        GB_COPY_DOCUMENT.Width = 550;
                        GB_COPY_DOCUMENT.Height = 195;

                        GB_COPY_DOCUMENT.Border3DStyle = Border3DStyle.Bump;
                        GB_COPY_DOCUMENT.BorderStyle = BorderStyle.Fixed3D;

                        GB_COPY_DOCUMENT.Visible = true;
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
                        GB_COPY_DOCUMENT.Visible = false;
                    }
                    else if (pSub_Panel == "COPY_BUDGET")
                    {
                        GB_COPY_DOCUMENT.Visible = false;
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
                vBUDGET_PERIOD = IGR_BUDGET_ADD_LIST.GetCellValue("BUDGET_PERIOD");
            }
            else
            {
                vBUDGET_PERIOD = BUDGET_PERIOD.EditValue;
            }
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(vBUDGET_PERIOD));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0626");
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

            object vBUDGET_ADD_HEADER_ID = BUDGET_ADD_HEADER_ID.EditValue;
            string vBUDGET_TYPE = iString.ISNull(BUDGET_TYPE.EditValue);
            if (TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                vBUDGET_ADD_HEADER_ID = IGR_BUDGET_ADD_LIST.GetCellValue("BUDGET_ADD_HEADER_ID");
                vBUDGET_TYPE = iString.ISNull(IGR_BUDGET_ADD_LIST.GetCellValue("BUDGET_TYPE"));
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //편성예산//
                if (vBUDGET_TYPE == "11")
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0626_001.xlsx";
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

                        IDA_PRINT_BUDGET_ADD_HEADER.SetSelectParamValue("P_BUDGET_ADD_HEADER_ID", vBUDGET_ADD_HEADER_ID);
                        IDA_PRINT_BUDGET_ADD_HEADER.Fill();

                        IDA_PRINT_BUDGET_ADD_LINE.SetSelectParamValue("P_BUDGET_ADD_HEADER_ID", IDA_PRINT_BUDGET_ADD_HEADER.CurrentRow["BUDGET_ADD_HEADER_ID"]);
                        IDA_PRINT_BUDGET_ADD_LINE.Fill();

                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_TYPE", IDA_PRINT_BUDGET_ADD_HEADER.CurrentRow["BUDGET_TYPE"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_HEADER_ID", IDA_PRINT_BUDGET_ADD_HEADER.CurrentRow["BUDGET_ADD_HEADER_ID"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.Fill();

                        vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_BUDGET_ADD_HEADER, IDA_PRINT_BUDGET_ADD_LINE, IDA_PRINT_APPROVAL_STEP_PERSON, vSOB_DESC, vLOCAL_DATE);

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
                else
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0626_002.xlsx";
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

                        IDA_PRINT_BUDGET_ADD_HEADER.SetSelectParamValue("P_BUDGET_ADD_HEADER_ID", vBUDGET_ADD_HEADER_ID);
                        IDA_PRINT_BUDGET_ADD_HEADER.Fill();

                        IDA_PRINT_BUDGET_ADD_LINE.SetSelectParamValue("P_BUDGET_ADD_HEADER_ID", IDA_PRINT_BUDGET_ADD_HEADER.CurrentRow["BUDGET_ADD_HEADER_ID"]);
                        IDA_PRINT_BUDGET_ADD_LINE.Fill();

                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_TYPE", IDA_PRINT_BUDGET_ADD_HEADER.CurrentRow["BUDGET_TYPE"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_HEADER_ID", IDA_PRINT_BUDGET_ADD_HEADER.CurrentRow["BUDGET_ADD_HEADER_ID"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.Fill();

                        vPageNumber = xlPrinting.ExcelWrite_Etc(IDA_PRINT_BUDGET_ADD_HEADER, IDA_PRINT_BUDGET_ADD_LINE, IDA_PRINT_APPROVAL_STEP_PERSON, vSOB_DESC, vLOCAL_DATE);

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
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.AddOver();
                        Budget_Add_Line_Insert();
                    }
                    else
                    {
                        if (Check_Added() == true)
                        {
                            return;
                        }

                        IDA_BUDGET_ADD_HEADER.AddOver();
                        Budget_Add_Header_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.AddUnder();
                        Budget_Add_Line_Insert();
                    }
                    else
                    {
                        if (Check_Added() == true)
                        {
                            return;
                        }

                        IDA_BUDGET_ADD_HEADER.AddUnder();
                        Budget_Add_Header_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_ADD_HEADER.Update();
                    }
                    catch
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.Cancel();
                    }
                    else
                    {
                        IDA_BUDGET_ADD_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.Delete();
                    }
                    else
                    {
                        IDA_BUDGET_ADD_HEADER.Delete();
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

        private void FCMF0626_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_ADD_LIST.FillSchema();
            IDA_BUDGET_ADD_HEADER.FillSchema();
            IDA_BUDGET_ADD_LINE.FillSchema();

            //서브판넬 
            Init_Sub_Panel(false, "ALL");
        }

        private void FCMF0626_Shown(object sender, EventArgs e)
        {
            W_BUDGET_PERIOD_FR.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_BUDGET_PERIOD_TO.EditValue = iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today,1)); 
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            C1_ALL_RECORD_FLAG.BringToFront();
            C2_ALL_RECORD_FLAG.BringToFront();
            BTN_CHG_APPROVAL_STEP.BringToFront();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        private void BTN_PRE_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //전표 작성중이면 저장후 작업해야 함
            if (iString.ISNull(BUDGET_ADD_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10128"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Check_Added() == true)
            {
                return;
            }

            //서브판넬 
            C_OLD_BUDGET_DEPT_CODE.EditValue = BUDGET_DEPT_CODE.EditValue;
            C_OLD_BUDGET_DEPT_ID.EditValue = BUDGET_DEPT_ID.EditValue;
            C_OLD_BUDGET_DEPT_NAME.EditValue = BUDGET_DEPT_NAME.EditValue;

            C_OLD_BUDGET_PERIOD.EditValue = BUDGET_PERIOD.EditValue;
            C_OLD_BUDGET_ADD_HEADER_ID.EditValue = BUDGET_ADD_HEADER_ID.EditValue;
            C_OLD_BUDGET_ADD_NUM.EditValue = BUDGET_ADD_NUM.EditValue;

            C_NEW_BUDGET_ADD_HEADER_ID.EditValue = DBNull.Value;
            C_NEW_BUDGET_ADD_NUM.EditValue = string.Empty;
            C_NEW_BUDGET_DEPT_CODE.EditValue = BUDGET_DEPT_CODE.EditValue;
            C_NEW_BUDGET_DEPT_ID.EditValue = BUDGET_DEPT_ID.EditValue;
            C_NEW_BUDGET_DEPT_NAME.EditValue = BUDGET_DEPT_NAME.EditValue;

            Init_Sub_Panel(true, "COPY_BUDGET");
        }

        private void C_BTN_SET_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iString.ISNull(C_NEW_BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(C_NEW_BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(C_NEW_BUDGET_DEPT_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(C_NEW_BUDGET_DEPT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_COPY_BUDGET_ADD_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_COPY_BUDGET_ADD_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_COPY_BUDGET_ADD_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_COPY_BUDGET_ADD_REQ.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            C_OLD_BUDGET_ADD_HEADER_ID.EditValue = IDC_COPY_BUDGET_ADD_REQ.GetCommandParamValue("O_NEW_BUDGET_ADD_HEADER_ID");
            C_NEW_BUDGET_ADD_NUM.EditValue = IDC_COPY_BUDGET_ADD_REQ.GetCommandParamValue("O_NEW_BUDGET_ADD_NUM"); 

            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        private void C_BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "COPY_BUDGET");

            if (CB_NEW_SEARCH_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                SearchDB_DTL(C_NEW_BUDGET_ADD_HEADER_ID.EditValue);

                BTN_REQ_APPROVAL.Enabled = true;
                BTN_CHG_APPROVAL_STEP.Enabled = true;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
            }            
        }

        private void irbALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRADIO = sender as ISRadioButtonAdv;
            W_APPROVE_STATUS.EditValue = vRADIO.RadioButtonValue;

            //버튼제어 및 체크박스 제어.
            if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "N")
            {
                BTN_REQ_APPROVAL.Enabled = true;
                BTN_CHG_APPROVAL_STEP.Enabled = true;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
            } 
            else if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "Y")
            {
                BTN_REQ_APPROVAL.Enabled = false;
                BTN_CHG_APPROVAL_STEP.Enabled = false;
                BTN_CANCEL_REQ_APPROVAL.Enabled = true;
            }
            else
            {
                BTN_REQ_APPROVAL.Enabled = false;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
                BTN_CHG_APPROVAL_STEP.Enabled = false;
            }
            SearchDB();
        }
         
        private void IGR_BUDGET_ADD_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_BUDGET_ADD_LIST.RowCount > 0)
            {
                C2_ALL_RECORD_FLAG.CheckedState = C1_ALL_RECORD_FLAG.CheckedState;

                SearchDB_DTL(IGR_BUDGET_ADD_LIST.GetCellValue("BUDGET_ADD_HEADER_ID"));
            }
        }

        private void ibtREQ_APPROVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_BUDGET_ADD_LIST.Update();

            object mValue;
            int mRowCount = IGR_BUDGET_ADD_LINE.RowCount;
            int mIDX_COL = IGR_BUDGET_ADD_LINE.GetColumnToIndex("APPROVE_STATUS");

            for (int R = 0; R < mRowCount; R++)
            {
                if (iString.ISNull(IGR_BUDGET_ADD_LINE.GetCellValue(R, mIDX_COL)) == "N".ToString())
                {// 승인미요청 건에 대해서 승인 처리.
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_TYPE")));
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_PERIOD")));
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("DEPT_ID")));
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID")));
                    IDC_APPROVE_REQUEST.ExecuteNonQuery();

                    mValue = DBNull.Value;
                    mValue = IDC_APPROVE_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                    IGR_BUDGET_ADD_LINE.SetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("APPROVE_STATUS"), mValue);

                    mValue = DBNull.Value;
                    mValue = IDC_APPROVE_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                    IGR_BUDGET_ADD_LINE.SetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                }
            }
            IDA_BUDGET_ADD_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_ADD_LIST.Refillable = true;
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
            if (iString.ISNull(BUDGET_ADD_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_ADD_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_ADD_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            IDA_BUDGET_ADD_HEADER.Update();

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_BUDGET_ADD_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_ADD_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_ADD_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            
            if (IDC_EXEC_BUDGET_ADD_REQ.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_ADD_HEADER_ID.EditValue);
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
            if (iString.ISNull(BUDGET_ADD_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_ADD_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_ADD_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_BUDGET_ADD_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_BUDGET_ADD_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_BUDGET_ADD_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_BUDGET_ADD_REQ.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_ADD_HEADER_ID.EditValue);
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
            if (iString.ISNull(BUDGET_ADD_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_ADD_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_ADD_NUM.Focus();
                return;
            }

            DialogResult dlgResult = DialogResult.None;
            FCMF0626_APPR_STEP vFCMF0626_APPR_STEP = new FCMF0626_APPR_STEP(isAppInterfaceAdv1.AppInterface,
                                                                            BUDGET_ADD_NUM.EditValue, APPROVAL_STEP_SEQ.EditValue,
                                                                            BUDGET_PERIOD.EditValue, BUDGET_ADD_HEADER_ID.EditValue,
                                                                            BUDGET_TYPE_NAME.EditValue, BUDGET_TYPE.EditValue, 
                                                                            BUDGET_DEPT_NAME.EditValue, BUDGET_DEPT_CODE.EditValue, BUDGET_DEPT_ID.EditValue);
            dlgResult = vFCMF0626_APPR_STEP.ShowDialog();
            Application.DoEvents();

            vFCMF0626_APPR_STEP.Dispose();
            SearchDB_DTL(BUDGET_ADD_HEADER_ID.EditValue);
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
            ILD_BUDGET_TYPE.SetLookupParamValue("W_LOOKUP_TYPE", "ADD");
            ILD_BUDGET_TYPE.SetLookupParamValue("W_ENABLED_YN", "N"); 
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
            ILD_BUDGET_TYPE.SetLookupParamValue("W_LOOKUP_TYPE", "ADD");
            ILD_BUDGET_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
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
                V_PLAN_AMOUNT.EditValue = 0;
                V_THIS_AMOUNT.EditValue = 0;
                V_GAP_AMOUNT.EditValue = 0;
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

        private void IDA_BUDGET_ADD_HEADER_UpdateCompleted(object pSender)
        {
            SearchDB_DTL(BUDGET_ADD_HEADER_ID.EditValue);
        }

        private void IDA_BUDGET_ADD_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (iString.ISNull(e.Row["EXPENDITURE_DATE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10242", string.Format("&&VALUE:=예상지급일(Expenditure Date)")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //} 
        }


        #endregion

    }
}