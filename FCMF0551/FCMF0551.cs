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

namespace FCMF0551
{
    public partial class FCMF0551 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCurrency_Code;

        #endregion;

        #region ----- Constructor -----

        public FCMF0551()
        {
            InitializeComponent();
        }

        public FCMF0551(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void Search_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_DAILY_PLAN_DUE.TabIndex)
            {
                if (iString.ISNull(W_PLAN_DATE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PLAN_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_PLAN_DATE.Focus();
                }

                IGR_TR_DAILY_PLAN_DUE.LastConfirmChanges();
                IDA_TR_DAILY_PLAN_DUE.OraSelectData.AcceptChanges();
                IDA_TR_DAILY_PLAN_DUE.Refillable = true;

                IDA_TR_DAILY_PLAN_DUE.Fill();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_DAILY_PLAN.TabIndex)
            {
                if (iString.ISNull(W_PLAN_DATE_1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PLAN_DATE_1))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_PLAN_DATE_1.Focus();
                }

                IDA_TR_DAILY_PLAN_INCOME.Fill();
                IDA_TR_DAILY_PLAN_EXPENSE.Fill();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_RESULT.TabIndex)
            {
                if (iString.ISNull(W_PLAN_DATE_2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PLAN_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_PLAN_DATE_2.Focus();
                }
                IDA_TR_DAILY_RESULT.Fill();
                IDA_TR_DAILY_RESULT_SUM.Fill();
                IGR_TR_DAILY_RESULT.Focus();
            }
        }

        private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Insert_TR_Plan_Income()
        {
            IGR_TR_DAILY_PLAN_INCOME.SetCellValue("CURRENCY_CODE", mCurrency_Code);
            Init_GRID_INCOME_STATUS(mCurrency_Code);

            IGR_TR_DAILY_PLAN_INCOME.Focus();
        }

        private void Insert_TR_Plan_Expense()
        {
            IGR_TR_DAILY_PLAN_EXPENSE.SetCellValue("CURRENCY_CODE", mCurrency_Code);
            Init_GRID_EXPENSE_STATUS(mCurrency_Code);

            IGR_TR_DAILY_PLAN_EXPENSE.Focus();
        }

        private bool SET_TR_DAILY_PLAN()
        {
            if (iString.ISNull(W_PLAN_DATE_1.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PLAN_DATE_1))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PLAN_DATE_1.Focus();
            }
            
            IDC_SET_TR_DAILY_PLAN.ExecuteNonQuery();

            string vSTATUS = iString.ISNull(IDC_SET_TR_DAILY_PLAN.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_TR_DAILY_PLAN.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_TR_DAILY_PLAN.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_TR_DAILY_PLAN.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private bool SET_TR_DAILY_PLAN_CONFIRM(string pCONFIRM_STATUS)
        {
            IDC_SET_TR_DAILY_PLAN_CONFIRM.SetCommandParamValue("W_CONFIRM_STATUS", pCONFIRM_STATUS);
            IDC_SET_TR_DAILY_PLAN_CONFIRM.ExecuteNonQuery();

            string vSTATUS = iString.ISNull(IDC_SET_TR_DAILY_PLAN_CONFIRM.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_TR_DAILY_PLAN_CONFIRM.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_TR_DAILY_PLAN_CONFIRM.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_TR_DAILY_PLAN_CONFIRM.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private bool SET_TR_DAILY_PLAN_CLOSED(string pCLOSED_STATUS)
        {
            IDC_SET_TR_DAILY_PLAN_CLOSED.SetCommandParamValue("W_CLOSED_STATUS", pCLOSED_STATUS);
            IDC_SET_TR_DAILY_PLAN_CLOSED.ExecuteNonQuery();

            string vSTATUS = iString.ISNull(IDC_SET_TR_DAILY_PLAN_CLOSED.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_TR_DAILY_PLAN_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_TR_DAILY_PLAN_CLOSED.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_TR_DAILY_PLAN_CLOSED.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private void Init_GRID_INCOME_STATUS(object pCURRENCY)
        {
            int vIDX_EXCHANGE_RATE = IGR_TR_DAILY_PLAN_INCOME.GetColumnToIndex("EXCHANGE_RATE");
            int vIDX_CURR_GL_AMOUNT = IGR_TR_DAILY_PLAN_INCOME.GetColumnToIndex("CURR_GL_AMOUNT");
            if (iString.ISNull(pCURRENCY) == iString.ISNull(mCurrency_Code))
            {
                decimal vAMOUNT = iString.ISDecimaltoZero(IGR_TR_DAILY_PLAN_INCOME.GetCellValue("CURR_GL_AMOUNT"), 0);
                if (vAMOUNT != 0)
                {
                    IGR_TR_DAILY_PLAN_INCOME.SetCellValue("CURR_GL_AMOUNT", null);
                }

                IGR_TR_DAILY_PLAN_INCOME.GridAdvExColElement[vIDX_EXCHANGE_RATE].Insertable = 0;
                IGR_TR_DAILY_PLAN_INCOME.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Insertable = 0;
                IGR_TR_DAILY_PLAN_INCOME.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Updatable = 0; 
            }
            else
            {
                IGR_TR_DAILY_PLAN_INCOME.GridAdvExColElement[vIDX_EXCHANGE_RATE].Insertable = 1;
                IGR_TR_DAILY_PLAN_INCOME.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Insertable = 1;
                IGR_TR_DAILY_PLAN_INCOME.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Updatable = 1; 
            }
            IGR_TR_DAILY_PLAN_INCOME.ResetDraw = true;
        }

        private void Init_GRID_EXPENSE_STATUS(object pCURRENCY)
        {
            int vIDX_EXCHANGE_RATE = IGR_TR_DAILY_PLAN_EXPENSE.GetColumnToIndex("EXCHANGE_RATE");
            int vIDX_CURR_GL_AMOUNT = IGR_TR_DAILY_PLAN_EXPENSE.GetColumnToIndex("CURR_GL_AMOUNT");
            if (iString.ISNull(pCURRENCY) == iString.ISNull(mCurrency_Code))
            {
                decimal vAMOUNT = iString.ISDecimaltoZero(IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("CURR_GL_AMOUNT"), 0);
                if (vAMOUNT != 0)
                {
                    IGR_TR_DAILY_PLAN_EXPENSE.SetCellValue("CURR_GL_AMOUNT", null);
                }

                IGR_TR_DAILY_PLAN_EXPENSE.GridAdvExColElement[vIDX_EXCHANGE_RATE].Insertable = 0;
                IGR_TR_DAILY_PLAN_EXPENSE.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Insertable = 0;
                IGR_TR_DAILY_PLAN_EXPENSE.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Updatable = 0;
            }
            else
            {
                IGR_TR_DAILY_PLAN_EXPENSE.GridAdvExColElement[vIDX_EXCHANGE_RATE].Insertable = 1;
                IGR_TR_DAILY_PLAN_EXPENSE.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Insertable = 1;
                IGR_TR_DAILY_PLAN_EXPENSE.GridAdvExColElement[vIDX_CURR_GL_AMOUNT].Updatable = 1;                
            }
            IGR_TR_DAILY_PLAN_EXPENSE.ResetDraw = true;
        }

        private decimal Init_GL_Amount(object pEXCHANGE_RATE, object pCURR_GL_AMOUNT)
        {
            decimal vGL_AMOUNT = iString.ISDecimaltoZero(pEXCHANGE_RATE, 0) *
                                 iString.ISDecimaltoZero(pCURR_GL_AMOUNT, 0);
            return vGL_AMOUNT;
        }

        private decimal Get_Exchange_Rate(object pAPPLY_DATE, object pCURRENCY_CODE)
        {
            decimal vExchange_Rate = 0;

            IDC_EXCHANGE_RATE.SetCommandParamValue("P_APPLY_DATE", pAPPLY_DATE);
            IDC_EXCHANGE_RATE.SetCommandParamValue("P_CURRENCY_CODE", pCURRENCY_CODE);
            IDC_EXCHANGE_RATE.ExecuteNonQuery();
            vExchange_Rate = iString.ISDecimaltoZero(IDC_EXCHANGE_RATE.GetCommandParamValue("X_EXCHANGE_RATE"));
            return vExchange_Rate;
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

        //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_NAME_L))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        #endregion;

        #region ----- XL Export Methods ----

        //private void ExportXL(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        //{
        //    System.Windows.Forms.Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    System.Windows.Forms.Application.DoEvents();

        //    string vMessageText = string.Empty;
        //    int vPageTotal = 0;
        //    int vPageNumber = 0;

        //    int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

        //    string vWeekName = WeekName(W_PLAN_DATE.DateTimeValue);
        //    string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일[{3}]", W_PLAN_DATE.DateTimeValue.Year, W_PLAN_DATE.DateTimeValue.Month, W_PLAN_DATE.DateTimeValue.Day, vWeekName);

        //    int vCountRowGrid = pGrid.RowCount;
        //    if (vCountRowGrid > 0)
        //    {
        //        vMessageText = string.Format("Excel Export Starting");
        //        isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();

        //        XLPrinting_ xlPrinting = new XLPrinting_(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //        try
        //        {
        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.OpenFileNameExcel = "FCMF0551_001.xls";
        //            //-------------------------------------------------------------------------------------

        //            //-------------------------------------------------------------------------------------
        //            bool isOpen = xlPrinting.XLFileOpen();
        //            //-------------------------------------------------------------------------------------

        //            //-------------------------------------------------------------------------------------
        //            if (isOpen == true)
        //            {
        //                vPageNumber = xlPrinting.LineWrite(pGrid,, vDate);

        //                ////[SAVE]
        //                xlPrinting.Save("PLAN_"); //저장 파일명


        //                vPageTotal = vPageTotal + vPageNumber;
        //            }
        //            //-------------------------------------------------------------------------------------

        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.Dispose();
        //            //-------------------------------------------------------------------------------------
        //        }
        //        catch (System.Exception ex)
        //        {
        //            string vMessage = ex.Message;
        //            xlPrinting.Dispose();
        //        }
        //    }

        //    //-------------------------------------------------------------------------
        //    vMessageText = string.Format("Excel Export End [Total Page : {0}]", vPageTotal);
        //    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //    System.Windows.Forms.Application.DoEvents();

        //    System.Windows.Forms.Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Application.DoEvents();
        //}

        #endregion;

        #region ----- Week Name Method ----

        private string WeekName(System.DateTime pDate)
        {
            string vWeekName = string.Empty;

            switch (pDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    vWeekName = "월";
                    break;
                case DayOfWeek.Tuesday:
                    vWeekName = "화";
                    break;
                case DayOfWeek.Wednesday:
                    vWeekName = "수";
                    break;
                case DayOfWeek.Thursday:
                    vWeekName = "목";
                    break;
                case DayOfWeek.Friday:
                    vWeekName = "금";
                    break;
                case DayOfWeek.Saturday:
                    vWeekName = "토";
                    break;
                case DayOfWeek.Sunday:
                    vWeekName = "일";
                    break;
            }

            return vWeekName;
        }

        #endregion;

        #region ----- XL Print Methods ----

        private void XLPrinting(string pOutChoice, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_Sum)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "TR_Daily_Plan";

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vWeekName = WeekName(W_PLAN_DATE_2.DateTimeValue);
            IDC_DAY_WEEK_DESC_P.ExecuteNonQuery();
            if (IDC_DAY_WEEK_DESC_P.ExcuteError)
            {
                
            }
            else
            {
                vWeekName = iString.ISNull(IDC_DAY_WEEK_DESC_P.GetCommandParamValue("O_WEEK_DESC"));
            }

            string vDate = string.Format("Date : {0:yyyy.MM.dd} ({1})", iDate.ISGetDate(W_PLAN_DATE_2.EditValue), vWeekName);

            int vCountRowGrid = pGrid.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0551_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(pGrid, pGrid_Sum, vDate);

                        ////[PRINT]
                        if (pOutChoice == "FILE")
                        {
                            ////[SAVE]
                            xlPrinting.SAVE(vSaveFileName); //저장 파일명
                        }
                        else
                        {
                            ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                            xlPrinting.Printing(1, vPageNumber);
                        }  
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
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
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
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_TR_DAILY_PLAN_INCOME.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_INCOME.AddOver();
                        Insert_TR_Plan_Income();
                    }
                    else if(IDA_TR_DAILY_PLAN_EXPENSE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_EXPENSE.AddOver();
                        Insert_TR_Plan_Expense();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_TR_DAILY_PLAN_INCOME.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_INCOME.AddUnder();
                        Insert_TR_Plan_Income();
                    }
                    else if (IDA_TR_DAILY_PLAN_EXPENSE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_EXPENSE.AddUnder();
                        Insert_TR_Plan_Expense();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_TR_DAILY_PLAN_DUE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_DUE.Update();
                    }
                    else if (IDA_TR_DAILY_PLAN_INCOME.IsFocused || IDA_TR_DAILY_PLAN_EXPENSE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_INCOME.Update();
                        IDA_TR_DAILY_PLAN_EXPENSE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_TR_DAILY_PLAN_DUE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_DUE.Cancel();
                    }
                    else if(IDA_TR_DAILY_PLAN_INCOME.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_INCOME.Cancel(); 
                    }
                    else if(IDA_TR_DAILY_PLAN_EXPENSE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_EXPENSE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_TR_DAILY_PLAN_DUE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_DUE.Delete();
                    }
                    else if (IDA_TR_DAILY_PLAN_INCOME.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_INCOME.Delete();
                    }
                    else if (IDA_TR_DAILY_PLAN_EXPENSE.IsFocused)
                    {
                        IDA_TR_DAILY_PLAN_EXPENSE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    //idaDailyPlanPRINT.Fill();
                    XLPrinting("PRINT", IGR_TR_DAILY_RESULT, IGR_TR_DAILY_RESULT_SUM);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //idaDailyPlanPRINT.Fill();
                    //ExportXL(igrTR_DAILY_SUM, idaDailyPlanPRINT);
                    XLPrinting("FILE", IGR_TR_DAILY_RESULT, IGR_TR_DAILY_RESULT_SUM);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        
        private void FCMF0551_Load(object sender, EventArgs e)
        {
            V_BATCH_TYPE_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_PLAN_STATUS_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_RB_CLOSED_ALL_2.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        private void FCMF0551_Shown(object sender, EventArgs e)
        {
            W_PLAN_DATE.EditValue = DateTime.Today;
            W_DUE_DATE_FR.EditValue = W_PLAN_DATE.EditValue;
            W_DUE_DATE_TO.EditValue = W_PLAN_DATE.EditValue;
            
            W_PLAN_DATE_1.EditValue = DateTime.Today;
            W_PLAN_DATE_2.EditValue = DateTime.Today;

            idcBASE_CURRENCY.ExecuteNonQuery();
            mCurrency_Code = idcBASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE");

            IDA_TR_DAILY_PLAN_DUE.FillSchema();
            IDA_TR_DAILY_PLAN_INCOME.FillSchema();
            IDA_TR_DAILY_PLAN_EXPENSE.FillSchema(); 
        }

        private void BTN_CHANGE_DUE_DATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_CHANGE_DUE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_CHANGE_DUE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_CHANGE_DUE_DATE.Focus();
                return;
            }

            int vIDX_SELECT_YN = IGR_TR_DAILY_PLAN_DUE.GetColumnToIndex("SELECT_YN");
            int vIDX_DUE_DATE = IGR_TR_DAILY_PLAN_DUE.GetColumnToIndex("DUE_DATE");

            for (int r = 0; r < IGR_TR_DAILY_PLAN_DUE.RowCount; r++)
            {
                if (iString.ISNull(IGR_TR_DAILY_PLAN_DUE.GetCellValue(r, vIDX_SELECT_YN)) == "Y")
                {
                    IGR_TR_DAILY_PLAN_DUE.SetCellValue(r, vIDX_DUE_DATE, V_CHANGE_DUE_DATE.EditValue);
                }
            }
        }

        private void IGR_TR_DAILY_PLAN_INCOME_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            decimal vGL_AMOUNT = 0;
            if (e.ColIndex == IGR_TR_DAILY_PLAN_INCOME.GetColumnToIndex("EXCHANGE_RATE"))
            {
                vGL_AMOUNT = Init_GL_Amount(e.NewValue, IGR_TR_DAILY_PLAN_INCOME.GetCellValue("CURR_GL_AMOUNT"));
                IGR_TR_DAILY_PLAN_INCOME.SetCellValue("GL_AMOUNT", vGL_AMOUNT);
            }
            else if (e.ColIndex == IGR_TR_DAILY_PLAN_INCOME.GetColumnToIndex("CURR_GL_AMOUNT"))
            {
                vGL_AMOUNT = Init_GL_Amount(IGR_TR_DAILY_PLAN_INCOME.GetCellValue("EXCHANGE_RATE"), e.NewValue);
                IGR_TR_DAILY_PLAN_INCOME.SetCellValue("GL_AMOUNT", vGL_AMOUNT);
            } 
        } 

        private void IGR_TR_DAILY_PLAN_EXPENSE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            decimal vGL_AMOUNT = 0;
            if (e.ColIndex == IGR_TR_DAILY_PLAN_EXPENSE.GetColumnToIndex("EXCHANGE_RATE"))
            {
                vGL_AMOUNT = Init_GL_Amount(e.NewValue, IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("CURR_GL_AMOUNT"));
                IGR_TR_DAILY_PLAN_EXPENSE.SetCellValue("GL_AMOUNT", vGL_AMOUNT); 
            }
            else if (e.ColIndex == IGR_TR_DAILY_PLAN_EXPENSE.GetColumnToIndex("CURR_GL_AMOUNT"))
            {
                vGL_AMOUNT = Init_GL_Amount(IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("EXCHANGE_RATE"), e.NewValue);
                IGR_TR_DAILY_PLAN_EXPENSE.SetCellValue("GL_AMOUNT", vGL_AMOUNT); 
            }
        }

        private void BTN_SET_DAILY_PLAN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (SET_TR_DAILY_PLAN() == true)
            {
                Search_DB();
            }
        }
        
        private void BTN_SET_CONFIRM_YES_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (SET_TR_DAILY_PLAN_CONFIRM("YES") == true)
            {
                Search_DB();
            }
        }

        private void BTN_CONFIRM_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (SET_TR_DAILY_PLAN_CONFIRM("CANCEL") == true)
            {
                Search_DB();
            }
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (SET_TR_DAILY_PLAN_CLOSED("YES") == true)
            {
                Search_DB();
            }
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (SET_TR_DAILY_PLAN_CLOSED("CANCEL") == true)
            {
                Search_DB();
            }
        }

        private void V_BATCH_TYPE_ALL_CheckChanged(object sender, EventArgs e)
        {
            if (V_BATCH_TYPE_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_BATCH_TYPE.EditValue = V_BATCH_TYPE_ALL.RadioCheckedString;
            }
        }

        private void V_BATCH_TYPE_P_CheckChanged(object sender, EventArgs e)
        {
            if (V_BATCH_TYPE_P.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_BATCH_TYPE.EditValue = V_BATCH_TYPE_P.RadioCheckedString;
            }
        }

        private void V_BATCH_TYPE_R_CheckChanged(object sender, EventArgs e)
        {
            if (V_BATCH_TYPE_R.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_BATCH_TYPE.EditValue = V_BATCH_TYPE_R.RadioCheckedString;
            }
        }

        private void V_PLAN_STATUS_A_CheckChanged(object sender, EventArgs e)
        {
            if (V_PLAN_STATUS_A.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CONFIRM_STATUS_1.EditValue = V_PLAN_STATUS_A.RadioCheckedString;
            }
        }

        private void V_CONFIRM_N_CheckChanged(object sender, EventArgs e)
        {
            if (V_CONFIRM_N.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CONFIRM_STATUS_1.EditValue = V_CONFIRM_N.RadioCheckedString;
            }
        }

        private void V_CONFIRM_Y_CheckChanged(object sender, EventArgs e)
        {
            if (V_CONFIRM_Y.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CONFIRM_STATUS_1.EditValue = V_CONFIRM_Y.RadioCheckedString;
            }
        }

        private void V_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_YN = IGR_TR_DAILY_PLAN_DUE.GetColumnToIndex("SELECT_YN");

            for (int r = 0; r < IGR_TR_DAILY_PLAN_DUE.RowCount; r++)
            {
                if (V_SELECT_YN.CheckBoxString != iString.ISNull(IGR_TR_DAILY_PLAN_DUE.GetCellValue(r, vIDX_SELECT_YN)))
                { 
                    IGR_TR_DAILY_PLAN_DUE.SetCellValue(r, vIDX_SELECT_YN, V_SELECT_YN.CheckBoxString);
                }
            }
        }

        private void V_RB_CLOSED_ALL_2_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_CLOSED_ALL_2.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CLOSED_STATUS_2.EditValue = V_RB_CLOSED_ALL_2.RadioCheckedString;
            }
        }

        private void V_RB_CLOSED_NO_2_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_CLOSED_NO_2.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CLOSED_STATUS_2.EditValue = V_RB_CLOSED_NO_2.RadioCheckedString;
            }
        }

        private void V_RB_CLOSED_YES_2_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_CLOSED_YES_2.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CLOSED_STATUS_2.EditValue = V_RB_CLOSED_YES_2.RadioCheckedString;
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_BATCH_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_BATCH.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_CURRENCY_CODE_INCOME_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CURRENCY_CODE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaLOAN_USE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("LOAN_USE", "Y");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaFUND_MOVE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("FUND_MOVE", "Y");
        }

        private void ILA_TR_PLAN_ACC_CONTROL_EXPENSE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_TR_PLAN_ACC_CONTROL.SetLookupParamValue("W_ITEM_TYPE", IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("ITEM_TYPE"));
        }

        private void ILA_TR_PLAN_ACC_CONTROL_INCOME_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_TR_PLAN_ACC_CONTROL.SetLookupParamValue("W_ITEM_TYPE", IGR_TR_DAILY_PLAN_INCOME.GetCellValue("ITEM_TYPE"));
        }

        private void ILA_CURRENCY_CODE_INCOME_SelectedRowData(object pSender)
        {
            string vCURRENCY = iString.ISNull(IGR_TR_DAILY_PLAN_INCOME.GetCellValue("CURRENCY_CODE"));
            Init_GRID_INCOME_STATUS(vCURRENCY);

            decimal vEXCHANGE_RATE = Get_Exchange_Rate(W_PLAN_DATE_1.EditValue, IGR_TR_DAILY_PLAN_INCOME.GetCellValue("CURRENCY_CODE"));
            IGR_TR_DAILY_PLAN_INCOME.SetCellValue("EXCHANGE_RATE", vEXCHANGE_RATE);
            IGR_TR_DAILY_PLAN_INCOME.SetCellValue("GL_AMOUNT", Init_GL_Amount(vEXCHANGE_RATE, IGR_TR_DAILY_PLAN_INCOME.GetCellValue("CURR_GL_AMOUNT")));
        }

        private void ILA_CURRENCY_CODE_EXPENSE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CURRENCY_CODE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_CODE_EXPENSE_SelectedRowData(object pSender)
        {
            string vCURRENCY = iString.ISNull(IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("CURRENCY_CODE"));
            Init_GRID_EXPENSE_STATUS(vCURRENCY);

            decimal vEXCHANGE_RATE = Get_Exchange_Rate(W_PLAN_DATE_1.EditValue, IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("CURRENCY_CODE"));
            IGR_TR_DAILY_PLAN_EXPENSE.SetCellValue("EXCHANGE_RATE", vEXCHANGE_RATE);
            IGR_TR_DAILY_PLAN_EXPENSE.SetCellValue("GL_AMOUNT", Init_GL_Amount(vEXCHANGE_RATE, IGR_TR_DAILY_PLAN_EXPENSE.GetCellValue("CURR_GL_AMOUNT")));
        }

        #endregion
        
        #region ----- Adapeter Event -----

        private void IDA_TR_DAILY_PLAN_DUE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10145"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_TR_DAILY_PLAN_INCOME_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Init_GRID_INCOME_STATUS(mCurrency_Code);
                return;
            }
            Init_GRID_INCOME_STATUS(pBindingManager.DataRow["CURRENCY_CODE"]);
        }

        private void IDA_TR_DAILY_PLAN_INCOME_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ITEM_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", "Item Type(항목구분)")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) != iString.ISNull(mCurrency_Code) && iString.ISNull(e.Row["CURR_GL_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["GL_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10126"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10530"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_TR_DAILY_PLAN_EXPENSE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Init_GRID_EXPENSE_STATUS(mCurrency_Code);
                return;
            }
            Init_GRID_EXPENSE_STATUS(pBindingManager.DataRow["CURRENCY_CODE"]);
        }

        private void IDA_TR_DAILY_PLAN_EXPENSE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ITEM_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", "Item Type(항목구분)")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) != iString.ISNull(mCurrency_Code) && iString.ISNull(e.Row["CURR_GL_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["GL_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10126"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10530"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaTR_DAILY_PLAN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["GL_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_DR_CR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) != iString.ISNull(mCurrency_Code) && iString.ISNull(e.Row["GL_CURR_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["GL_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10126"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
        }

        #endregion

    }
}