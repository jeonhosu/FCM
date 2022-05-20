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

namespace SOMF0992
{
    public partial class SOMF0992_DETAIL : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mSTATUS;
        object mTAX_BILL_ISSUE_NO;
        object mTAX_BILL_NO;
        object mHOMETAX_ISSUE_NO;

        #endregion;

        #region ----- Constructor -----

        public SOMF0992_DETAIL()
        {
            InitializeComponent();
        }

        public SOMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public SOMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS; 
        }
        
        public SOMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS, object pTAX_BILL_ISSUE_NO)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS;
            mTAX_BILL_ISSUE_NO  = pTAX_BILL_ISSUE_NO; 
        }

        public SOMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS,
                                object pTAX_BILL_ISSUE_NO, object pTAX_BILL_NO, object pHOMETAX_ISSUE_NO)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS;
            mTAX_BILL_ISSUE_NO = pTAX_BILL_ISSUE_NO;
            mTAX_BILL_NO = pTAX_BILL_NO;
            mHOMETAX_ISSUE_NO = pHOMETAX_ISSUE_NO;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_TAX_BILL_ISSUE.SetSelectParamValue("W_TAX_BILL_ISSUE_NO", mTAX_BILL_ISSUE_NO);
            IDA_TAX_BILL_ISSUE.Fill(); 

            //수정세금계산 발행일 경우 수정사유 활성화//
            if (mSTATUS == "FIX")
            {
                SRC_TAX_BILL_ISSUE_NO.EditValue = mTAX_BILL_ISSUE_NO;
                SRC_HOMETAX_ISSUE_NO.EditValue = mHOMETAX_ISSUE_NO;

                TAX_BILL_FIX_TYPE.Updatable = true;
            }
        }

        private void Init_Header_Insert()
        {
            IDA_TAX_BILL_ISSUE.AddUnder();
            
            //과세구분.
            IDC_DEFAULT_TB_VAT_TYPE.ExecuteNonQuery();
            TAX_BILL_VAT_TYPE.EditValue = IDC_DEFAULT_TB_VAT_TYPE.GetCommandParamValue("O_TB_VAT_TYPE");
            TB_VAT_TYPE_NAME.EditValue = IDC_DEFAULT_TB_VAT_TYPE.GetCommandParamValue("O_TB_VAT_TYPE_NAME");
            VAT_TAX_TYPE.EditValue = IDC_DEFAULT_TB_VAT_TYPE.GetCommandParamValue("O_VAT_TAX_TYPE");
            VAT_RATE.EditValue = IDC_DEFAULT_TB_VAT_TYPE.GetCommandParamValue("O_VAT_RATE"); 

            //청구구분
            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TB_REQ_TYPE");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            TAX_BILL_REQ_TYPE.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
            TB_REQ_TYPE_NAME.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");

            ISSUE_DATE.EditValue = iDate.ISGetDate();

            //초기화//
            IDC_GET_OWNER_INFO.ExecuteNonQuery();
        }

        private void Init_Line_Insert()
        {
            if (iConv.ISNull(ISSUE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10298"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            IDA_TAX_BILL_ISSUE_LINE.AddUnder();

            IDC_GET_DATE_DAY_P.SetCommandParamValue("P_DATE", ISSUE_DATE.EditValue);
            IDC_GET_DATE_DAY_P.ExecuteNonQuery();

            IGR_TAX_BILL_ISSUE_LINE.SetCellValue("WRITE_MM", IDC_GET_DATE_DAY_P.GetCommandParamValue("O_MONTH"));
            IGR_TAX_BILL_ISSUE_LINE.SetCellValue("WRITE_DD", IDC_GET_DATE_DAY_P.GetCommandParamValue("O_DAY"));

            IGR_TAX_BILL_ISSUE_LINE.CurrentCellMoveTo(IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("WRITE_DD"));
            IGR_TAX_BILL_ISSUE_LINE.Focus();

        }

        private bool Sync_VAT_Amount(object pVAT_TAX_TYPE)
        {
            IGR_TAX_BILL_ISSUE_LINE.ResetDraw = true;

            try
            {
                int vIDX_SUPPLY_AMOUNT = IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT");
                int vIDX_VAT_AMOUNT = IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT");
                decimal vSUPPLY_AMOUNT = 0;
                decimal vVAT_AMOUNT = 0;
                for (int r = 0; r < IGR_TAX_BILL_ISSUE_LINE.RowCount; r++)
                {
                    vSUPPLY_AMOUNT = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue(r, vIDX_SUPPLY_AMOUNT));

                    IDC_VAT_AMT_P.SetCommandParamValue("W_VAT_TAX_TYPE", pVAT_TAX_TYPE);
                    IDC_VAT_AMT_P.SetCommandParamValue("W_SUPPLY_AMT", vSUPPLY_AMOUNT);
                    IDC_VAT_AMT_P.ExecuteNonQuery();
                    vVAT_AMOUNT = iConv.ISDecimaltoZero(IDC_VAT_AMT_P.GetCommandParamValue("O_VAT_AMT"));

                    IGR_TAX_BILL_ISSUE_LINE.SetCellValue(r, vIDX_VAT_AMOUNT, vVAT_AMOUNT);
                }
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                return false;
            }
            return true;
        }

        private void Sync_Total_Amount(int pRow_Index, decimal pSupply_Amount, decimal pVat_Amount)
        {            
            decimal vSUPPLY_AMOUNT = 0;
            decimal vVAT_AMOUNT = 0;

            int vIDX_SUPPLY_AMOUNT = IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT");
            int vIDX_VAT_AMOUNT = IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT");


            for (int r = 0; r < IGR_TAX_BILL_ISSUE_LINE.RowCount; r++)
            {
                if (r == pRow_Index)
                {
                    vSUPPLY_AMOUNT = vSUPPLY_AMOUNT + pSupply_Amount;
                    vVAT_AMOUNT = vVAT_AMOUNT + pVat_Amount;
                }
                else
                {
                    vSUPPLY_AMOUNT = vSUPPLY_AMOUNT + iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue(r, vIDX_SUPPLY_AMOUNT));
                    vVAT_AMOUNT = vVAT_AMOUNT + iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue(r, vIDX_VAT_AMOUNT));
                }
            }

            TOTAL_AMOUNT.EditValue = vSUPPLY_AMOUNT + vVAT_AMOUNT;
            SUPPLY_AMOUNT.EditValue = vSUPPLY_AMOUNT;
            VAT_AMOUNT.EditValue = vVAT_AMOUNT;
        }

        private void Show_Address_Seller()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, DBNull.Value, SELL_ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                SELL_ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_Buyer()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, DBNull.Value, BUY_ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                BUY_ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
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
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE.AddOver();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE.AddUnder();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_TAX_BILL_ISSUE.Update();
                    }
                    catch (Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.Cancel();
                        IDA_TAX_BILL_ISSUE.Cancel();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE.Delete();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Evevnt -----

        private void SOMF0992_Load(object sender, EventArgs e)
        {             
            IDA_TAX_BILL_ISSUE.FillSchema();

            if (mSTATUS == "INSERT")
            {
                Init_Header_Insert();
            }
            else
            {
                Search_DB();
            }
        }

        private void SELL_ADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Seller();
            }
        }

        private void BUY_ADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Buyer();
            }
        }

        private void BUY_ADDR2_KeyDown(object pSender, KeyEventArgs e)
        {

        }

        private void BTN_L_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Init_Line_Insert();
        }

        private void BTN_L_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_TAX_BILL_ISSUE_LINE.Update();
        }

        private void BTN_L_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            IDA_TAX_BILL_ISSUE_LINE.Delete();
            IDA_TAX_BILL_ISSUE.Update();
            Sync_Total_Amount(-1, 0, 0);
        }

        private void BTN_L_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_TAX_BILL_ISSUE_LINE.Cancel();
            Sync_Total_Amount(-1, 0, 0);
        }

        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Sync_Total_Amount(-1, 0, 0); 
            IDA_TAX_BILL_ISSUE.Update();
        }

        private void BTN_ISSUE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Questioin", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
             
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;  
             
            IDC_SET_TRANSFER_BILL365.SetCommandParamValue("W_ISSUE_DATE", ISSUE_DATE.EditValue);
            IDC_SET_TRANSFER_BILL365.SetCommandParamValue("W_TAX_BILL_ISSUE_NO", TAX_BILL_ISSUE_NO.EditValue);
            IDC_SET_TRANSFER_BILL365.SetCommandParamValue("W_DELAY_ISSUE_YN", DELAY_ISSUE_YN.CheckBoxString);
            IDC_SET_TRANSFER_BILL365.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_TRANSFER_BILL365.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_TRANSFER_BILL365.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            } 

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        private void ISSUE_DATE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            int vIDX_WRITE_MM = IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("WRITE_MM");

            IDC_GET_DATE_DAY_P.SetCommandParamValue("P_DATE", ISSUE_DATE.EditValue);
            IDC_GET_DATE_DAY_P.ExecuteNonQuery();
            object vNEW_WRITE_MM = IDC_GET_DATE_DAY_P.GetCommandParamValue("O_MONTH");
            for (int r = 0; r < IGR_TAX_BILL_ISSUE_LINE.RowCount; r++)
            {
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue(r, vIDX_WRITE_MM, vNEW_WRITE_MM);
            }
        }

        private void IGR_TAX_BILL_ISSUE_LINE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            decimal vVAT_RATE = iConv.ISDecimaltoZero(VAT_RATE.EditValue);
            decimal vQTY = 0;
            decimal vUNIT_PRICE = 0;
            decimal vSUPPLY_AMOUNT = 0;
            decimal vVAT_AMOUNT = 0;
            if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY"))
            {
                vQTY = iConv.ISDecimaltoZero(e.NewValue, 1);
                vUNIT_PRICE = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue("UNIT_PRICE"), 0);
                vSUPPLY_AMOUNT = vQTY * vUNIT_PRICE;
                vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("SUPPLY_AMOUNT", vSUPPLY_AMOUNT);
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);
            }
            else if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE"))
            {
                vQTY = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue("QTY"), 1);
                vUNIT_PRICE = iConv.ISDecimaltoZero(e.NewValue, 1);
                vSUPPLY_AMOUNT = vQTY * vUNIT_PRICE;
                vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("SUPPLY_AMOUNT", vSUPPLY_AMOUNT);
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);
            }
            else if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT"))
            {
                vSUPPLY_AMOUNT = iConv.ISDecimaltoZero(e.NewValue, 0);
                vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);
            }
            Sync_Total_Amount(e.RowIndex, vSUPPLY_AMOUNT, vVAT_AMOUNT);
        }
         
        #endregion

        #region ----- Lookup Event ------

        private void SetCommon(object pGROUP_CODE, object pENABLED_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pENABLED_YN);
        }

        private void ILA_TB_ISSUE_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_ISSUE_STATUS", "Y");   
        }

        private void ILA_TB_HT_ISSUE_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_HT_ISSUE_STATUS", "Y");   
        }

        private void ILA_TB_VAT_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_VAT_TYPE", "Y");   
        }

        private void ILA_TB_VAT_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VAT_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_TB_REQ_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_REQ_TYPE", "Y");   
        }
         
        private void ILA_CUSTOMER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_SUPP_CUST_TYPE", "C");
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        
        private void ILA_TB_VAT_TYPE_SelectedRowData(object pSender)
        {
            Sync_VAT_Amount(VAT_TAX_TYPE.EditValue); 
        }

        private void ILA_TB_FIX_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_FIX_TYPE", "Y");   
        }

        #endregion

        #region ----- Adapter Event ------

        private void IDA_TAX_BILL_ISSUE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ISSUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10144"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vCURR_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE")).Date;
            DateTime vISSUE_DATE = iDate.ISGetDate(e.Row["ISSUE_DATE"]).Date;

            if (vCURR_DATE < vISSUE_DATE)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90009", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ISSUE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iDate.ISYearMonth(vCURR_DATE) == iDate.ISYearMonth(vISSUE_DATE))
            {
                //
            }
            else
            {
                if (vCURR_DATE <= iDate.ISGetDate(string.Format("{0}-10", iDate.ISYearMonth(vCURR_DATE))) && iDate.ISGetDate(iDate.ISMonth_1st(iDate.ISDate_Month_Add(vCURR_DATE, -1))) <= vISSUE_DATE)
                {
                    //
                }
                else
                {
                    if (DELAY_ISSUE_YN.CheckBoxString == "N" && MessageBoxAdv.Show("지연교부 발행대상입니다. 지연교부발행여부를 체크하시겠습니까?", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.Yes) 
                    {   
                        e.Cancel = true;
                        return;
                    }
                }
            }

            if (iConv.ISNull(e.Row["BUY_VENDOR_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SUPPLY_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["VAT_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10281"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["TAX_BILL_VAT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10614"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUY_TAX_REG_NO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(BUY_TAX_REG_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUY_VENDOR_CEO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(BUY_VENDOR_CEO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUY_BIZ_STATUS"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(BUY_BIZ_STATUS))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUY_BIZ_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(BUY_BIZ_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUY_USER_EMAIL"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10615"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["TAX_BILL_REQ_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(TB_REQ_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (mSTATUS == "FIX")
            {
                if (iConv.ISNull(e.Row["SRC_TAX_BILL_ISSUE_NO"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(SRC_TAX_BILL_ISSUE_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (iConv.ISNull(e.Row["SRC_HOMETAX_ISSUE_NO"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(SRC_HOMETAX_ISSUE_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (iConv.ISNull(e.Row["TAX_BILL_FIX_TYPE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(TAX_BILL_FIX_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            else
            {
                if (iConv.ISNull(e.Row["TAX_BILL_FIX_TYPE"]) == string.Empty)
                {
                    MessageBoxAdv.Show("신규등록시 수정세금계산서 사유는 등록 할 수 없습니다.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }

        }

        private void IDA_TAX_BILL_ISSUE_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        { 
            if (iConv.ISNull(e.Row["WRITE_MM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10616"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WRITE_DD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10617"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ITEM_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SUPPLY_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["VAT_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10281"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["UNIT_PRICE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10618"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        #endregion



    }
}