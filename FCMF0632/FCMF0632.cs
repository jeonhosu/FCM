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

namespace FCMF0632
{
    public partial class FCMF0632 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        bool mSUB_SHOW_FLAG = false;
        #endregion;

        #region ----- Constructor -----

        public FCMF0632()
        {
            InitializeComponent();
        }

        public FCMF0632(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        //private void Set_Default_CAPACITY()
        //{
        //    // Budget Select Type.
        //    idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "BUDGET_CAPACITY");
        //    idcDEFAULT_VALUE.ExecuteNonQuery();

        //    APPROVE_STATUS_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
        //    APPROVE_STATUS_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
        //}

        private void Set_Default_Value()
        {
            // Budget Select Type.
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "BUDGET_SELECT_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();

            W_FROM_ACCOUNT_CONTROL_ID.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
            W_FROM_ACCOUNT_CODE.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");            
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
                
                try
                {
                    IDA_BUDGET_MOVE_HEADER.SetSelectParamValue("P_BUDGET_MOVE_HEADER_ID", pBUDGET_MOVE_HEADER_ID);
                    IDA_BUDGET_MOVE_HEADER.Fill();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }
                BUDGET_PERIOD.Focus();
            }
        }


        private void Set_Item_Status()
        { 
            if (V2_ALL_RECORD_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                REMARK.Insertable = false;
                REMARK.Updatable = false;
                REMARK.Refresh();

                //금액.
                AMOUNT.Insertable = false;
                AMOUNT.Updatable = false;
                AMOUNT.Refresh();

                CAUSE_NAME.Insertable = false;
                CAUSE_NAME.Updatable = false;
                CAUSE_NAME.Refresh();

                DESCRIPTION.Insertable = false;
                DESCRIPTION.Updatable = false;
                DESCRIPTION.Refresh();
            }
            else
            {
                REMARK.Insertable = true;
                REMARK.Updatable = true;
                REMARK.Refresh();

                //금액.
                AMOUNT.Insertable = true;
                AMOUNT.Updatable = true;
                AMOUNT.Refresh();

                CAUSE_NAME.Insertable = true;
                CAUSE_NAME.Updatable = true;
                CAUSE_NAME.Refresh();

                DESCRIPTION.Insertable = true;
                DESCRIPTION.Updatable = true;
                DESCRIPTION.Refresh();
            } 
        }

        private void Budget_Move_Header_Insert()
        {
            BUDGET_PERIOD.EditValue = iDate.ISYearMonth(DateTime.Today);
            
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            BUDGET_PERIOD.Focus();
        }

        private void Budget_Move_Line_Insert()
        {
            FROM_DEPT_CODE.EditValue = DEPT_CODE.EditValue;
            FROM_DEPT_ID.EditValue = DEPT_ID.EditValue;
            FROM_DEPT_NAME.EditValue = DEPT_NAME.EditValue;
            AMOUNT.EditValue = 0;
            FROM_DEPT_NAME.Focus();
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

        private void Set_Total_Amount()
        {
            decimal vTotal_Amount = 0;
            object vAmount;
            IGR_BUDGET_MOVE_LINE.ResetDraw = true;

            int vIDXCol = IGR_BUDGET_MOVE_LINE.GetColumnToIndex("AMOUNT");
            for (int r = 0; r < IGR_BUDGET_MOVE_LINE.RowCount; r++)
            {
                vAmount = 0;
                vAmount = IGR_BUDGET_MOVE_LINE.GetCellValue(r, vIDXCol);
                vTotal_Amount = vTotal_Amount + iString.ISDecimaltoZero(vAmount);
            }
            TOTAL_AMOUNT.EditValue = vTotal_Amount;
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
                        //GB_COPY_DOCUMENT.Left = 180;
                        //GB_COPY_DOCUMENT.Top = 95;

                        //GB_COPY_DOCUMENT.Width = 550;
                        //GB_COPY_DOCUMENT.Height = 195;

                        //GB_COPY_DOCUMENT.Border3DStyle = Border3DStyle.Bump;
                        //GB_COPY_DOCUMENT.BorderStyle = BorderStyle.Fixed3D;

                        //GB_COPY_DOCUMENT.Visible = true;
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
                        //GB_COPY_DOCUMENT.Visible = false;
                    }
                    else if (pSub_Panel == "COPY_BUDGET")
                    {
                        //GB_COPY_DOCUMENT.Visible = false;
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
                    if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    {
                        IDA_BUDGET_MOVE_LINE.AddOver();
                        Budget_Move_Line_Insert();
                    }
                    else
                    {
                        if (Check_Added() == true)
                        {
                            return;
                        }

                        IDA_BUDGET_MOVE_HEADER.AddOver();
                        Budget_Move_Header_Insert();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    {
                        IDA_BUDGET_MOVE_LINE.AddUnder();
                        Budget_Move_Line_Insert();
                    }
                    else
                    {
                        if (Check_Added() == true)
                        {
                            return;
                        }

                        IDA_BUDGET_MOVE_HEADER.AddUnder();
                        Budget_Move_Header_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_MOVE_HEADER.Update();
                    }
                    catch(Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_MOVE_HEADER.IsFocused)
                    {
                        IDA_BUDGET_MOVE_LINE.Cancel();
                        IDA_BUDGET_MOVE_HEADER.Cancel();
                    }
                    else if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    {
                        IDA_BUDGET_MOVE_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_MOVE_HEADER.IsFocused)
                    {
                        IDA_BUDGET_MOVE_HEADER.Delete();
                    }
                    else if (IDA_BUDGET_MOVE_LINE.IsFocused)
                    {
                        IDA_BUDGET_MOVE_LINE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0632_Load(object sender, EventArgs e)
        {
            V_ALL_RECORD_FLAG.CheckBoxValue = "N";
            W_BUDGET_PERIOD_FR.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_BUDGET_PERIOD_TO.EditValue = iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 1));
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_ALL_RECORD_FLAG.BringToFront();
            V2_ALL_RECORD_FLAG.BringToFront();
            BTN_CHG_APPROVAL_STEP.BringToFront();
            System.Windows.Forms.Cursor.Current = Cursors.Default;

            Set_Default_Value();
            
            //서브판넬 
            Init_Sub_Panel(false, "ALL");
        }

        private void FCMF0632_Shown(object sender, EventArgs e)
        {
            IDA_BUDGET_MOVE_HEADER.FillSchema();
            IDA_BUDGET_MOVE_LINE.FillSchema();
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

        private void IGR_BUDGET_MOVE_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_BUDGET_MOVE_LIST.RowCount > 0)
            {
                V2_ALL_RECORD_FLAG.CheckedState = V_ALL_RECORD_FLAG.CheckedState;

                SearchDB_DTL(IGR_BUDGET_MOVE_LIST.GetCellValue("BUDGET_MOVE_HEADER_ID"));
            }
        }

        //private void ibtREQ_APPROVE_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    idaBUDGET_MOVE.Update();
        //    Set_Approve_Request();     // 승인요청.            
        //}

        //private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    EXE_BUDGET_MOVE_STATUS(PERIOD_NAME_0.EditValue, APPROVE_STATUS_0.EditValue, "OK");
        //}

        //private void ibtnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    EXE_BUDGET_MOVE_STATUS(PERIOD_NAME_0.EditValue, APPROVE_STATUS_0.EditValue, "CANCEL");
        //}

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
            FCMF0632_APPR_STEP vFCMF0632_APPR_STEP = new FCMF0632_APPR_STEP(isAppInterfaceAdv1.AppInterface,
                                                                            BUDGET_MOVE_NUM.EditValue, APPROVAL_STEP_SEQ.EditValue,
                                                                            BUDGET_PERIOD.EditValue, BUDGET_MOVE_HEADER_ID.EditValue,
                                                                            BUDGET_TYPE_NAME.EditValue, BUDGET_TYPE.EditValue,
                                                                            DEPT_NAME.EditValue, DEPT_CODE.EditValue, DEPT_ID.EditValue);
            dlgResult = vFCMF0632_APPR_STEP.ShowDialog();
            Application.DoEvents();

            vFCMF0632_APPR_STEP.Dispose();
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

        private void ilaSELECT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUDGET_SELECT_TYPE", DBNull.Value, "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBUDGET_MOVE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_PERIOD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Budget Period(예산년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FROM_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=From Department(전용 (전) 부서)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FROM_ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=From Account Code(전용 (전) 계정)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["TO_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=To Department(전용 (후) 부서)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["TO_ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=TO Account Code(전용 (후) 계정)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Amount(예산금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CAUSE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Cause(신청사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaBUDGET_MOVE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["LAST_YN"]) == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10262"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

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

            IDC_EXEC_BUDGET_MOVE_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_MOVE_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_MOVE_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_EXEC_BUDGET_MOVE_REQ.ExcuteError || vSTATUS == "F")
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

            IDC_CANCEL_BUDGET_MOVE_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_BUDGET_MOVE_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_BUDGET_MOVE_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_BUDGET_MOVE_REQ.ExcuteError || vSTATUS == "F")
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

    }
}