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

namespace FCMF0315
{
    public partial class FCMF0315 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        string mBase_Currency_Code;

        #endregion;

        #region ----- Constructor -----

        public FCMF0315()
        {
            InitializeComponent();
        }

        public FCMF0315(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void GetAccountBook()
        {
            IDC_BASE_CURRENCY.ExecuteNonQuery();
            mBase_Currency_Code = iConv.ISNull(IDC_BASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE"));
        }

        private void Search_DB()
        {
            if (iConv.ISNull(W_SALE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_SALE_DATE_FR.Focus();
                return;
            }

            if (iConv.ISNull(W_SALE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_SALE_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(W_SALE_DATE_FR.EditValue) > Convert.ToDateTime(W_SALE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_SALE_DATE_FR.Focus();
                return;
            }

            string vSALE_NUM = iConv.ISNull(IGR_ASSET_SALE_DTL_LIST.GetCellValue("SALE_NUM"));
            int vCOL_IDX = IGR_ASSET_SALE_DTL_LIST.GetColumnToIndex("SALE_NUM");
            IDA_ASSET_SALE_DTL_LIST.Fill();
            if (iConv.ISNull(vSALE_NUM) != string.Empty)
            {
                for (int i = 0; i < IGR_ASSET_SALE_DTL_LIST.RowCount; i++)
                {
                    if (vSALE_NUM == iConv.ISNull(IGR_ASSET_SALE_DTL_LIST.GetCellValue(i, vCOL_IDX)))
                    {
                        IGR_ASSET_SALE_DTL_LIST.CurrentCellMoveTo(i, vCOL_IDX);
                        IGR_ASSET_SALE_DTL_LIST.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
        }

        private void Search_DB_DTL()
        {
            IGR_ASSET_SALE_LINE.LastConfirmChanges();
            IDA_ASSET_SALE_LINE.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_LINE.Refillable = true;

            IDA_ASSET_SALE_HEADER.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_HEADER.Refillable = true;

            IDA_ASSET_SALE_HEADER.Fill(); 
        }

        private void Search_DB_DTL_LINE()
        {
            IDC_GET_PROFIT_AMOUNT.ExecuteNonQuery();
            IDA_ASSET_SALE_LINE.Fill();
        }

        private void INIT_INSERT()
        {
            CURRENCY_CODE.EditValue = mBase_Currency_Code;

            IDC_USER_INFO_P.ExecuteNonQuery();
            DEPT_ID.EditValue = IDC_USER_INFO_P.GetCommandParamValue("O_DEPT_ID");
            DEPT_NAME.EditValue = IDC_USER_INFO_P.GetCommandParamValue("O_DEPT_NAME");

            SALE_DATE.EditValue = iDate.ISGetDate();

            Init_Currency_Amount();

            SALE_AMOUNT.EditValue = 0;
            SALE_VAT_AMOUNT.EditValue = 0;
            ETC_AMOUNT.EditValue = 0;
            
            REMARK.Focus();
        }

        private void Init_Currency_Amount()
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) == string.Empty || iConv.ISNull(CURRENCY_CODE.EditValue) == mBase_Currency_Code)
            {
                if (iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue) != Convert.ToDecimal(0))
                {
                    EXCHANGE_RATE.EditValue = null;
                }
                if (iConv.ISDecimaltoZero(CURR_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    CURR_AMOUNT.EditValue = null;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                CURR_AMOUNT.ReadOnly = true;
                CURR_AMOUNT.Insertable = false;
                CURR_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                CURR_AMOUNT.TabStop = false;
            }
            else
            {
                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                CURR_AMOUNT.ReadOnly = false;
                CURR_AMOUNT.Insertable = true;
                CURR_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                CURR_AMOUNT.TabStop = true;
            }
            EXCHANGE_RATE.Invalidate();
            CURR_AMOUNT.Invalidate();
        }

        private void Init_SALE_Amount()
        {
            if (iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue) == 0)
            {
                return;
            }
            else if (iConv.ISDecimaltoZero(CURR_AMOUNT.EditValue) == 0)
            {
                return;
            }

            decimal mAMOUNT = iConv.ISDecimaltoZero(CURR_AMOUNT.EditValue) * iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue);
            SALE_AMOUNT.EditValue = mAMOUNT;             
        }

        // 부가세 관련 설정 제어 - 세액/공급가액(세액 * 10)
        private void Init_VAT_Amount()
        {
            IDC_VAT_AMT_P.SetCommandParamValue("W_VAT_TAX_TYPE", VAT_TAX_TYPE.EditValue);
            IDC_VAT_AMT_P.SetCommandParamValue("W_SUPPLY_AMT", SALE_AMOUNT.EditValue);
            IDC_VAT_AMT_P.ExecuteNonQuery();
            SALE_VAT_AMOUNT.EditValue = IDC_VAT_AMT_P.GetCommandParamValue("O_VAT_AMT"); 
        }

        private void Set_GRID_STATUS(object pSUM_FLAG, object pMODIFY_YN)
        {
            int vSTATUS = 0;                // INSERTABLE, UPDATABLE; 
            int vIDX_SALE_AMOUNT = IGR_ASSET_SALE_LINE.GetColumnToIndex("SALE_AMOUNT");
            int vIDX_DESCRIPTION = IGR_ASSET_SALE_LINE.GetColumnToIndex("DESCRIPTION"); 

            if (iConv.ISNull(pSUM_FLAG) == "N")
            {
                if (iConv.ISNull(pMODIFY_YN) == "Y")
                {
                    vSTATUS = 1;
                }
                else
                {
                    vSTATUS = 0;
                }
            }
            else
            {
                vSTATUS = 0;
            }

            IGR_ASSET_SALE_LINE.GridAdvExColElement[vIDX_SALE_AMOUNT].Insertable = vSTATUS;
            IGR_ASSET_SALE_LINE.GridAdvExColElement[vIDX_SALE_AMOUNT].Updatable = vSTATUS;

            IGR_ASSET_SALE_LINE.GridAdvExColElement[vIDX_DESCRIPTION].Insertable = vSTATUS;
            IGR_ASSET_SALE_LINE.GridAdvExColElement[vIDX_DESCRIPTION].Updatable = vSTATUS;
            
            // 범위를 지정해서 LOOP 이용//
            //int mGRID_START_COL = 17;   // 그리드 시작 COLUMN INDEX.
            //int mMax_Column = 24;       // 종료 COLUMN INDEX.

            //if (iConvert.ISNull(pSELECT_FLAG) == "Y")
            //{
            //    vSTATUS = 1;
            //}
            //else
            //{
            //    vSTATUS = 0;
            //}

            //for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            //{
            //    IGR_SCM_WYFC.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Insertable = vSTATUS;
            //    IGR_SCM_WYFC.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Updatable = vSTATUS;
            //}
        }

        private void Cal_Sale_Profit()
        {
            IDC_CAL_PROFIT_AMOUNT.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CAL_PROFIT_AMOUNT.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CAL_PROFIT_AMOUNT.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CAL_PROFIT_AMOUNT.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CAL_PROFIT_AMOUNT.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Search_DB_DTL_LINE();
        }

        private bool Check_Row_Added_Status()
        {
            Boolean Row_Added_Status = false;
            //헤더 체크  
            for (int r = 0; r < IDA_ASSET_SALE_HEADER.SelectRows.Count; r++)
            {
                if (IDA_ASSET_SALE_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_ASSET_SALE_HEADER.SelectRows[r].RowState == DataRowState.Modified)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            //헤더 변경없으면 라인 체크 
            if (Row_Added_Status == false)
            {
                for (int r = 0; r < IDA_ASSET_SALE_LINE.SelectRows.Count; r++)
                {
                    if (IDA_ASSET_SALE_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_ASSET_SALE_LINE.SelectRows[r].RowState == DataRowState.Modified)
                    {
                        Row_Added_Status = true;
                    }
                }
                if (Row_Added_Status == true)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            return (Row_Added_Status);
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
                    if (TB_MAIN.SelectedTab.TabIndex == TP_DETAIL.TabIndex)
                    {
                        Search_DB_DTL();
                    }
                    else
                    {
                        Search_DB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_ASSET_SALE_HEADER.IsFocused)
                    {
                        if (Check_Row_Added_Status() == true)
                        {
                            return;
                        }
                        else
                        {
                            IDA_ASSET_SALE_HEADER.AddOver();
                            INIT_INSERT();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ASSET_SALE_HEADER.IsFocused)
                    {
                        if (Check_Row_Added_Status() == true)
                        {
                            return;
                        }
                        else
                        {
                            IDA_ASSET_SALE_HEADER.AddUnder();
                            INIT_INSERT();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_ASSET_SALE_HEADER.IsFocused)
                    {
                        SALE_NUM.Focus();
                    }
                    if (IDA_ASSET_SALE_LINE.IsFocused)
                    {
                        IGR_ASSET_SALE_LINE.CurrentCellMoveTo(IGR_ASSET_SALE_LINE.GetColumnToIndex("AST_CATEGORY_NAME")); 
                    }
                    IDA_ASSET_SALE_HEADER.Update();
                    IDA_ASSET_SALE_LINE.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_SALE_HEADER.IsFocused)
                    {
                        IDA_ASSET_SALE_HEADER.Cancel();
                    }
                    else if (IDA_ASSET_SALE_LINE.IsFocused)
                    {
                        IDA_ASSET_SALE_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ASSET_SALE_HEADER.IsFocused)
                    {
                        IDA_ASSET_SALE_HEADER.Delete();
                    }
                    else if (IDA_ASSET_SALE_LINE.IsFocused)
                    {
                        if (iConv.ISNull(IDA_ASSET_SALE_LINE.CurrentRow["SUM_FLAG"]) == "N")
                        {
                            IDA_ASSET_SALE_LINE.Delete();
                        }                        
                    }
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event ----

        private void FCMF0315_Load(object sender, EventArgs e)
        {
            W_SALE_DATE_FR.EditValue = iDate.ISYearMonth(iDate.ISGetDate());
            W_SALE_DATE_TO.EditValue = iDate.ISGetDate();

            GetAccountBook();

            BTN_CONFIRM_OK.BringToFront();
            BTN_CONFIRM_CANCEL.BringToFront();
            BTN_GET_SALE_ASSET.BringToFront();
            BTN_RE_CALCULATION.BringToFront();
            BTN_RE_CAL_DPR_HISTORY.BringToFront();
            BTN_GET_SALE_ASSET.Enabled = false;
            BTN_RE_CALCULATION.Enabled = false;
            BTN_DPR_VIEW.BringToFront();

            IDA_ASSET_SALE_HEADER.FillSchema();
        }

        private void IGR_ASSET_SALE_DTL_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_ASSET_SALE_DTL_LIST.RowIndex < 0)
            {
                return; 
            }

            W_SALE_HEADER_ID.EditValue = IGR_ASSET_SALE_DTL_LIST.GetCellValue("SALE_HEADER_ID");

            TB_MAIN.SelectedIndex = (TP_DETAIL.TabIndex -1);
            TB_MAIN.SelectedTab.Focus();
            Application.DoEvents();

            Search_DB_DTL();
        }

        private void CURR_AMOUNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_SALE_Amount();
        }

        private void EXCHANGE_RATE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_SALE_Amount();
        }

        private void SALE_AMOUNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_VAT_Amount();
        }

        private void BTN_CONFIRM_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_SET_ASSET_SALE_CONFIRM.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_ASSET_SALE_CONFIRM.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_ASSET_SALE_CONFIRM.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_ASSET_SALE_CONFIRM.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_ASSET_SALE_CONFIRM.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Search_DB_DTL();
        }

        private void BTN_CONFIRM_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            IDC_CANCEL_ASSET_SALE_CONFIRM.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CANCEL_ASSET_SALE_CONFIRM.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CANCEL_ASSET_SALE_CONFIRM.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CANCEL_ASSET_SALE_CONFIRM.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CANCEL_ASSET_SALE_CONFIRM.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Search_DB_DTL();
        }

        private void BTN_GET_SALE_ASSET_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal vSALE_AMOUNT = iConv.ISDecimaltoZero(SALE_AMOUNT.EditValue, 0) - iConv.ISDecimaltoZero(ETC_AMOUNT.EditValue, 0);
            if(vSALE_AMOUNT < 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult vRESULT;
            FCMF0315_SET vFCMF0315_SET = new FCMF0315_SET(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                        , SALE_HEADER_ID.EditValue, SALE_NUM.EditValue, SALE_DATE.EditValue, vSALE_AMOUNT);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0315_SET, isAppInterfaceAdv1.AppInterface);
            vRESULT = vFCMF0315_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                Search_DB_DTL_LINE();
            }
            vFCMF0315_SET.Dispose();
        }

        private void BTN_DPR_VIEW_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISDecimaltoZero(IGR_ASSET_SALE_LINE.GetCellValue("ASSET_ID"), 0) != 0)
            {
                DialogResult vResult = DialogResult.None;
                FCMF0315_DPR vFCMF0315_DPR = new FCMF0315_DPR(MdiParent, isAppInterfaceAdv1.AppInterface, SALE_HEADER_ID.EditValue
                                                            , IGR_ASSET_SALE_LINE.GetCellValue("ASSET_ID")
                                                            , IGR_ASSET_SALE_LINE.GetCellValue("ASSET_CODE")
                                                            , IGR_ASSET_SALE_LINE.GetCellValue("ASSET_DESC"));
                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0315_DPR, isAppInterfaceAdv1.AppInterface);
                vResult = vFCMF0315_DPR.ShowDialog();
                if (vResult == DialogResult.OK)
                {

                }
                vFCMF0315_DPR.Dispose();
            }
        }

        private void BTN_RE_CALCULATION_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Cal_Sale_Profit();
        }

        private void BTN_RE_CAL_DPR_HISTORY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_CAL_DPR_HISTORY.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CAL_DPR_HISTORY.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CAL_DPR_HISTORY.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CAL_DPR_HISTORY.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CAL_DPR_HISTORY.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Search_DB_DTL();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ASSET_SALE_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_SALE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_ASSET_SALE_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_SALE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_VAT_TAX_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VAT_TAX_TYPE.SetLookupParamValue("W_AP_AR_TYPE", "AR");
            ILD_VAT_TAX_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y"); 
        }

        private void ILA_VAT_TAX_TYPE_SelectedRowData(object pSender)
        {
            Init_VAT_Amount();
        }

        private void ILA_VENDOR_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }	 
        
        private void ILA_VENDOR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_STATUS.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_DEPT_CODE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_CODE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_CURRENCY_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CURRENCY.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ILD_CURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_SelectedRowData(object pSender)
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) != string.Empty)
            {
                Init_Currency_Amount();
                if (iConv.ISNull(CURRENCY_CODE.EditValue) != mBase_Currency_Code)
                {
                    IDC_EXCHANGE_RATE.ExecuteNonQuery();
                    EXCHANGE_RATE.EditValue = IDC_EXCHANGE_RATE.GetCommandParamValue("O_EXCHANGE_RATE");

                    Init_SALE_Amount();
                }
            }
        }

        #endregion


        #region ----- Adapter Event -----

        private void IDA_ASSET_SALE_HEADER_ExcuteKeySearch(object pSender)
        {
            Search_DB_DTL();
        }

        private void IDA_ASSET_SALE_HEADER_UpdateCompleted(object pSender)
        {
            if (IDA_ASSET_SALE_HEADER.UpdateChangedRowCount != 0)
            {
                W_SALE_HEADER_ID.EditValue = SALE_HEADER_ID.EditValue;
                Search_DB_DTL();
            }
        }

        private void IDA_ASSET_SALE_HEADER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Currency_Amount();
            Search_DB_DTL_LINE();
            if (iConv.ISNull(STATUS.EditValue) == "ENTER")
            {
                BTN_GET_SALE_ASSET.Enabled = true;
                BTN_RE_CALCULATION.Enabled = true;
            }
            else
            {
                BTN_GET_SALE_ASSET.Enabled = false;
                BTN_RE_CALCULATION.Enabled = false;
            }
        }

        private void IDA_ASSET_SALE_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISDecimaltoZero(e.Row["SALE_AMOUNT"], 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10592"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_ASSET_SALE_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            
        }

        private void IDA_ASSET_SALE_LINE_UpdateCompleted(object pSender)
        {
            if (IDA_ASSET_SALE_HEADER.UpdateChangedRowCount != 0 || IDA_ASSET_SALE_LINE.UpdateChangedRowCount != 0)
            {
                Cal_Sale_Profit();
            }
        }

        private void IDA_ASSET_SALE_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Set_GRID_STATUS("T", "N");
                return;
            }
            Set_GRID_STATUS(pBindingManager.DataRow["SUM_FLAG"], pBindingManager.DataRow["MODIFY_YN"]);
        }


        #endregion

    }
}