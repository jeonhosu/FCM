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

namespace FCMF0633
{
    public partial class FCMF0633 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ----- Constructor -----

        public FCMF0633()
        {
            InitializeComponent();
        }

        public FCMF0633(Form pMainForm, ISAppInterface pAppInterface)
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
                SearchDB_DTL(BUDGET_MOVE_HEADER_ID.EditValue);
            }
            else
            {
                if (iString.ISNull(W_BUDGET_PERIOD_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_BUDGET_PERIOD_FR.Focus();
                    return;
                }
                if (iString.ISNull(W_BUDGET_PERIOD_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10219"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_BUDGET_PERIOD_FR.Focus();
                    return;
                }

                IDA_BUDGET_MOVE_HEADER.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", -1);
                IDA_BUDGET_MOVE_HEADER.Fill();

                IDA_BUDGET_MOVE_LIST.Fill();
                Set_Total_Amount();
                IGR_BUDGET_MOVE_LINE.Focus();
            }
        }

        private void SearchDB_DTL(object pBUDGET_MOVE_HEADER_ID)
        {
            if (iString.ISNull(pBUDGET_MOVE_HEADER_ID) != string.Empty)
            {
                TB_MAIN.SelectedIndex = 1;
                TB_MAIN.SelectedTab.Focus();

                Set_Item_Status();
                Application.DoEvents();

                IDA_BUDGET_MOVE_HEADER.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", pBUDGET_MOVE_HEADER_ID); 
                try
                {
                    IDA_BUDGET_MOVE_HEADER.Fill();
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
            decimal vThis_Amount = 0; 
            object vAmount;
            int vIDX_AMOUNT = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("AMOUNT"); 
            for (int r = 0; r < IGR_BUDGET_MOVE_LINE.RowCount; r++)
            {
                //당월
                vAmount = 0;
                vAmount = IGR_BUDGET_MOVE_LINE.GetCellValue(r, vIDX_AMOUNT);
                vThis_Amount = vThis_Amount + iString.ISDecimaltoZero(vAmount); 
            }
            V_TOTAL_AMOUNT.EditValue = vThis_Amount;
        }

        //private void EXE_BUDGET_MOVE_STATUS(object pPERIOD_NAME, object pAPPROVE_STATUS, object pAPPROVE_FLAG)
        //{
        //    IDA_BUDGET_MOVE_LIST.Update(); //수정사항 반영.

        //    Application.UseWaitCursor = true;
        //    this.Cursor = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    int vIDX_CHECK_YN = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("CHECK_YN");
        //    int vIDX_BUDGET_TYPE = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("BUDGET_TYPE");
        //    int vIDX_BUDGET_PERIOD = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("BUDGET_PERIOD");
        //    int vIDX_DEPT_ID = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("DEPT_ID");
        //    int vIDX_ACCOUNT_CONTROL_ID = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            
        //    string vSTATUS = "F";
        //    string vMESSAGE = null;
        //    for (int i = 0; i < IGR_BUDGET_MOVE_LINE.RowCount; i++)
        //    {
        //        if (iString.ISNull(IGR_BUDGET_MOVE_LINE.GetCellValue(i, vIDX_CHECK_YN), "N") == "Y")
        //        {
        //            IGR_BUDGET_MOVE_LINE.CurrentCellMoveTo(i, vIDX_CHECK_YN);
        //            IGR_BUDGET_MOVE_LINE.CurrentCellActivate(i, vIDX_CHECK_YN);

        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_MOVE_LINE.GetCellValue(i, vIDX_BUDGET_TYPE));
        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_MOVE_LINE.GetCellValue(i, vIDX_BUDGET_PERIOD));
        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_MOVE_LINE.GetCellValue(i, vIDX_DEPT_ID));
        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_MOVE_LINE.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("P_APPROVE_STATUS", pAPPROVE_STATUS);
        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("P_APPROVE_FLAG", pAPPROVE_FLAG);
        //            idcBUDGET_MOVE_STATUS.SetCommandParamValue("P_CHECK_YN", IGR_BUDGET_MOVE_LINE.GetCellValue(i, vIDX_CHECK_YN));
        //            idcBUDGET_MOVE_STATUS.ExecuteNonQuery();
        //            vSTATUS = iString.ISNull(idcBUDGET_MOVE_STATUS.GetCommandParamValue("O_STATUS"));
        //            vMESSAGE = iString.ISNull(idcBUDGET_MOVE_STATUS.GetCommandParamValue("O_MESSAGE"));
        //            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //            Application.DoEvents();

        //            if (idcBUDGET_MOVE_STATUS.ExcuteError || vSTATUS == "F")
        //            {
        //                Application.UseWaitCursor = false;
        //                this.Cursor = System.Windows.Forms.Cursors.Default;
        //                Application.DoEvents();
        //                if (vMESSAGE != string.Empty)
        //                {
        //                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                }
        //                return;
        //            }
        //        }
        //    }
        //    SearchDB();
        //    Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    Application.DoEvents();
        //}

        private void Set_Item_Status()
        {
            int mIDX_Col;

            if (C2_ALL_RECORD_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                REMARK.Insertable = false;
                REMARK.Updatable = false;

                //전용금액.
                mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("AMOUNT");
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 사유
                mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("CAUSE_NAME");
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                //// 당월예산.
                //mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("AMOUNT");
                //IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                //IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                //IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 비고.
                mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("DESCRIPTION");
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            }
            else
            {
                REMARK.Insertable = true;
                REMARK.Updatable = true;

                //전용금액.
                mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("AMOUNT");
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 사유
                mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("CAUSE_NAME");
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                //// 당월예산.
                //mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("AMOUNT");
                //IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                //IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                //IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

                // 비고.
                mIDX_Col = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("DESCRIPTION");
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                IGR_BUDGET_MOVE_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            }
            IGR_BUDGET_MOVE_LINE.ResetDraw = true;
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
            IDA_BUDGET_MOVE_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_MOVE_LIST.Refillable = true;
        }

        private bool Check_Added()
        {
            Boolean Row_Added_Status = false;
            
            //헤더 체크 
            for (int r = 0; r < IDA_BUDGET_MOVE_HEADER.SelectRows.Count; r++)
            {
                if (IDA_BUDGET_MOVE_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_BUDGET_MOVE_HEADER.SelectRows[r].RowState == DataRowState.Modified)
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
                for (int r = 0; r < IDA_BUDGET_MOVE_LINE.SelectRows.Count; r++)
                {
                    if (IDA_BUDGET_MOVE_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_BUDGET_MOVE_LINE.SelectRows[r].RowState == DataRowState.Modified)
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
                        V_RETURN_REMARK.EditValue = string.Empty;

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

                        //값 초기화.
                        V_APPROVAL_DESCRIPTION.EditValue = string.Empty;

                        GB_APPROVAL.Visible = true;
                    }
                    else if (pSub_Panel == "VIEW_DESCRIPTION")
                    {
                        GB_VIEW_DESCRIPTION.Left = 278;
                        GB_VIEW_DESCRIPTION.Top = 89;

                        GB_VIEW_DESCRIPTION.Width = 600;
                        GB_VIEW_DESCRIPTION.Height = 200;

                        GB_VIEW_DESCRIPTION.Border3DStyle = Border3DStyle.Bump;
                        GB_VIEW_DESCRIPTION.BorderStyle = BorderStyle.Fixed3D;

                        GB_VIEW_DESCRIPTION.Visible = true;
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
                        GB_APPROVAL.Visible = false;
                        GB_VIEW_DESCRIPTION.Visible = false;
                    }
                    else if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Visible = false;
                    }
                    else if (pSub_Panel == "APPROVAL")
                    {
                        GB_APPROVAL.Visible = false;
                    }
                    else if (pSub_Panel == "VIEW_DESCRIPTION")
                    {
                        GB_VIEW_DESCRIPTION.Visible = false;
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
                vBUDGET_PERIOD = IGR_BUDGET_MOVE_LIST.GetCellValue("BUDGET_PERIOD");
            }
            else
            {
                vBUDGET_PERIOD = BUDGET_PERIOD.EditValue;
            }
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(vBUDGET_PERIOD));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0633");
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

            object vBUDGET_MOVE_HEADER_ID = BUDGET_MOVE_HEADER_ID.EditValue;
            string vBUDGET_TYPE = iString.ISNull(BUDGET_TYPE.EditValue);
            if (TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                vBUDGET_MOVE_HEADER_ID = IGR_BUDGET_MOVE_LIST.GetCellValue("BUDGET_MOVE_HEADER_ID");
                vBUDGET_TYPE = iString.ISNull(IGR_BUDGET_MOVE_LIST.GetCellValue("BUDGET_TYPE"));
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1); 

            try
            {
                //편성예산//
                if(vBUDGET_TYPE == "11")
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0633_001.xlsx"; 
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

                        IDA_PRINT_BUDGET_MOVE_HEADER.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", vBUDGET_MOVE_HEADER_ID);
                        IDA_PRINT_BUDGET_MOVE_HEADER.Fill();

                        IDA_PRINT_BUDGET_MOVE_LINE.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", IDA_PRINT_BUDGET_MOVE_HEADER.CurrentRow["BUDGET_MOVE_HEADER_ID"]);
                        IDA_PRINT_BUDGET_MOVE_LINE.Fill();

                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_TYPE", IDA_PRINT_BUDGET_MOVE_HEADER.CurrentRow["BUDGET_TYPE"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_HEADER_ID", IDA_PRINT_BUDGET_MOVE_HEADER.CurrentRow["BUDGET_MOVE_HEADER_ID"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.Fill();

                        vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_BUDGET_MOVE_HEADER, IDA_PRINT_BUDGET_MOVE_LINE, IDA_PRINT_APPROVAL_STEP_PERSON, vSOB_DESC, vLOCAL_DATE);

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
                    xlPrinting.OpenFileNameExcel = "FCMF0633_002.xlsx";
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

                        IDA_PRINT_BUDGET_MOVE_HEADER.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", vBUDGET_MOVE_HEADER_ID);
                        IDA_PRINT_BUDGET_MOVE_HEADER.Fill();

                        IDA_PRINT_BUDGET_MOVE_LINE.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", IDA_PRINT_BUDGET_MOVE_HEADER.CurrentRow["BUDGET_MOVE_HEADER_ID"]);
                        IDA_PRINT_BUDGET_MOVE_LINE.Fill();

                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_TYPE", IDA_PRINT_BUDGET_MOVE_HEADER.CurrentRow["BUDGET_TYPE"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_HEADER_ID", IDA_PRINT_BUDGET_MOVE_HEADER.CurrentRow["BUDGET_MOVE_HEADER_ID"]);
                        IDA_PRINT_APPROVAL_STEP_PERSON.Fill();

                        vPageNumber = xlPrinting.ExcelWrite_Etc(IDA_PRINT_BUDGET_MOVE_HEADER, IDA_PRINT_BUDGET_MOVE_LINE, IDA_PRINT_APPROVAL_STEP_PERSON, vSOB_DESC, vLOCAL_DATE);

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
                    //if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    //{
                    //    IDA_BUDGET_MOVE_LINE.AddOver();
                    //    Budget_MOVE_Line_Insert();
                    //} 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    //{
                    //    IDA_BUDGET_MOVE_LINE.AddUnder();
                    //    Budget_MOVE_Line_Insert();
                    //} 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_MOVE_HEADER.Update();
                    }
                    catch
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    {
                        IDA_BUDGET_MOVE_LINE.Cancel();
                    }
                    else
                    {
                        IDA_BUDGET_MOVE_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_MOVE_LINE.CurrentRow.RowState == DataRowState.Added)
                    {
                        IDA_BUDGET_MOVE_LINE.Delete();
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

        private void FCMF0633_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_MOVE_LIST.FillSchema();
            IDA_BUDGET_MOVE_HEADER.FillSchema();
            IDA_BUDGET_MOVE_LINE.FillSchema();

            //서브판넬 
            Init_Sub_Panel(false, "ALL");
        }

        private void FCMF0633_Shown(object sender, EventArgs e)
        {
            W_BUDGET_PERIOD_FR.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_BUDGET_PERIOD_TO.EditValue = iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 1)); 
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            C1_ALL_RECORD_FLAG.BringToFront();
            C2_ALL_RECORD_FLAG.BringToFront();

            BTN_CHG_APPROVAL_STEP.BringToFront();
            BTN_VIEW_DESCRIPTION.BringToFront();
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
          
        private void IGR_BUDGET_MOVE_LIST_CellDoubleClick_1(object pSender)
        {
            if (IGR_BUDGET_MOVE_LIST.RowCount > 0)
            {
                C2_ALL_RECORD_FLAG.CheckedState = C1_ALL_RECORD_FLAG.CheckedState;

                SearchDB_DTL(IGR_BUDGET_MOVE_LIST.GetCellValue("BUDGET_MOVE_HEADER_ID"));
            }
        }

        private void BTN_REQ_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_MOVE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_MOVE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_MOVE_NUM.Focus();
                return;
            }

            //서브판넬 
            Init_Sub_Panel(true, "APPROVAL");
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
            if (iString.ISNull(BUDGET_MOVE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_MOVE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_MOVE_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_BUDGET_MOVE_APPR.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_BUDGET_MOVE_APPR.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_BUDGET_MOVE_APPR.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_BUDGET_MOVE_APPR.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_MOVE_HEADER_ID.EditValue);
        }

        private void C_BTN_EXEC_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
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
            if (iString.ISNull(BUDGET_MOVE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_MOVE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_MOVE_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_BUDGET_MOVE_APPR.SetCommandParamValue("P_APPROVAL_DESCRIPTION", V_APPROVAL_DESCRIPTION.EditValue);
            IDC_EXEC_BUDGET_MOVE_APPR.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_MOVE_APPR.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_MOVE_APPR.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_EXEC_BUDGET_MOVE_APPR.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_MOVE_HEADER_ID.EditValue);
        }

        private void BTN_VIEW_DESCRIPTION_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_APPROVAL_PERSON.RowIndex < 0)
            {
                return;
            }
            //서브판넬 
            Init_Sub_Panel(true, "VIEW_DESCRIPTION");
        }

        private void BTN_VIEW_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "VIEW_DESCRIPTION");
        }

        private void C_BTN_EXEC_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "APPROVAL");
        }

        private void BTN_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_MOVE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_MOVE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_MOVE_NUM.Focus();
                return;
            }
            
            //서브판넬 
            Init_Sub_Panel(true, "RETURN");
        }

        private void C_BTN_EXEC_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_RETURN_REMARK.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_RETURN_REMARK))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_MOVE_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_BUDGET_MOVE_RETURN.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_MOVE_RETURN.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_MOVE_RETURN.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_EXEC_BUDGET_MOVE_RETURN.ExcuteError || vSTATUS == "F")
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
            SearchDB_DTL(BUDGET_MOVE_HEADER_ID.EditValue);
        }

        private void C_BTN_RETURN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "RETURN");
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
            if (iString.ISNull(BUDGET_MOVE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_MOVE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_MOVE_NUM.Focus();
                return;
            }

            DialogResult dlgResult = DialogResult.None;
            FCMF0633_APPR_STEP vFCMF0633_APPR_STEP = new FCMF0633_APPR_STEP(isAppInterfaceAdv1.AppInterface,
                                                                            BUDGET_MOVE_NUM.EditValue, APPROVAL_STEP_SEQ.EditValue,
                                                                            BUDGET_PERIOD.EditValue, BUDGET_MOVE_HEADER_ID.EditValue,
                                                                            BUDGET_TYPE_NAME.EditValue, BUDGET_TYPE.EditValue, 
                                                                            DEPT_NAME.EditValue, DEPT_CODE.EditValue, DEPT_ID.EditValue);
            dlgResult = vFCMF0633_APPR_STEP.ShowDialog();
            Application.DoEvents();

            vFCMF0633_APPR_STEP.Dispose();
            SearchDB_DTL(BUDGET_MOVE_HEADER_ID.EditValue);
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

        private void ILA_DEPT_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_H.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT_H.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT_H.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT_H.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(W_BUDGET_PERIOD_FR.EditValue));
            ILD_DEPT_H.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(W_BUDGET_PERIOD_TO.EditValue));
        }

        private void ILA_DEPT_H_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_H.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT_H.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT_H.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT_H.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT_H.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ILA_FROM_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_FR.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT_FR.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT_FR.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT_FR.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ILA_TO_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ILA_ACCOUNT_CONTROL_V_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_V_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_FROM_ACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FROM_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_TO_ACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        //private void ilaBUDGET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        //{
        //    SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'MOVE'", "Y");
        //}

        private void ilaFROM_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_FR.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT_FR.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT_FR.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT_FR.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT_FR.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ilaTO_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_FR.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT_FR.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT_FR.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT_FR.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT_FR.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ilaFROM_ACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaTO_ACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CAUSE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_CAUSE", "Value1 = 'MOVE'", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_BUDGET_MOVE_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager == null)
            {
                V_TOTAL_AMOUNT.EditValue = 0;
            }
            Set_Total_Amount();
        }

        private void IDA_BUDGET_MOVE_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
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
            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(DEPT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void IDA_BUDGET_MOVE_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["FROM_ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["TO_ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (iString.ISDecimaltoZero(e.Row["AMOUNT"],0) == 0)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10537"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        } 

        #endregion



    }
}