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

namespace FCMF0992
{
    public partial class FCMF0992_DETAIL : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mSTATUS; 

        #endregion;

        #region ----- Constructor -----

        public FCMF0992_DETAIL()
        {
            InitializeComponent();
        }

        public FCMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public FCMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS; 
        }
        
        public FCMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS, object pTAX_BILL_ISSUE_NO)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS;
            TAX_BILL_ISSUE_NO.EditValue = pTAX_BILL_ISSUE_NO;  
        }

        public FCMF0992_DETAIL(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS,
                                object pTAX_BILL_ISSUE_NO, object pTAX_BILL_NO, object pHOMETAX_ISSUE_NO,
                                object pSRC_TAX_BILL_ISSUE_NO, object pSRC_TAX_BILL_NO, object pSRC_HOMETAX_ISSUE_NO, object pREL_TAX_BILL_ISSUE_NO)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS;

            TAX_BILL_ISSUE_NO.EditValue = pTAX_BILL_ISSUE_NO;
            TAX_BILL_NO.EditValue = pTAX_BILL_NO;
            HOMETAX_ISSUE_NO.EditValue = pHOMETAX_ISSUE_NO;

            SRC_TAX_BILL_ISSUE_NO.EditValue = pSRC_TAX_BILL_ISSUE_NO;
            SRC_HOMETAX_ISSUE_NO.EditValue = pSRC_HOMETAX_ISSUE_NO;
            REL_TAX_BILL_ISSUE_NO.EditValue = pREL_TAX_BILL_ISSUE_NO;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_TAX_BILL_ISSUE.Fill(); 
            Item_Status(iConv.ISNull(MINUS_TAX_BILL_FLAG.EditValue)); 
        }

        private void Search_DB_Minus()
        {
            IDA_TAX_BILL_ISSUE_MINUS.Fill(); 
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

            CASH_AMOUNT.EditValue = 0;
            CHECK_AMOUNT.EditValue = 0;
            NOTE_AMOUNT.EditValue = 0;
            CREDIT_AMOUNT.EditValue = 0;

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

        private void Item_Status(string pModifyStatus)
        {
            if (pModifyStatus == "Y")
            {
                SELL_REG_ID.Insertable = true;
                SELL_REG_ID.Updatable = true;

                SELL_ADDR1.Insertable = true;
                SELL_ADDR1.Updatable = true;

                SELL_ADDR2.Insertable = true;
                SELL_ADDR2.Updatable = true;
                
                SELL_BIZ_STATUS.Insertable = true;
                SELL_BIZ_STATUS.Updatable = true;

                SELL_BIZ_TYPE.Insertable = true;
                SELL_BIZ_TYPE.Updatable = true;

                SELL_USER_EMAIL.Insertable = true;
                SELL_USER_EMAIL.Updatable = true;

                SELL_USER_EMAIL.Insertable = true;
                SELL_USER_EMAIL.Updatable = true;

                BUY_TAX_REG_NO.Insertable = true;
                BUY_TAX_REG_NO.Updatable = true;

                BUY_REG_ID.Insertable = true;
                BUY_REG_ID.Updatable = true;

                BUY_VENDOR_NAME.Insertable = true;
                BUY_VENDOR_NAME.Updatable = true;

                BUY_VENDOR_CEO.Insertable = true;
                BUY_VENDOR_CEO.Updatable = true;

                BUY_ADDR1.Insertable = true;
                BUY_ADDR1.Updatable = true;

                BUY_ADDR2.Insertable = true;
                BUY_ADDR2.Updatable = true;

                BUY_BIZ_STATUS.Insertable = true;
                BUY_BIZ_STATUS.Updatable = true;

                BUY_BIZ_TYPE.Insertable = true;
                BUY_BIZ_TYPE.Updatable = true;

                BUY_USER_EMAIL.Insertable = true;
                BUY_USER_EMAIL.Updatable = true;

                BUY_USER2_EMAIL.Insertable = true;
                BUY_USER2_EMAIL.Updatable = true;

                DELAY_ISSUE_YN.Enabled = true;
                HOMETAX_SEND_TYPE.Enabled = true;

                BTN_L_INSERT.Enabled = true;
                BTN_L_DELETE.Enabled = true;

                TB_VAT_TYPE_NAME.Insertable = true;
                TB_VAT_TYPE_NAME.Updatable = true;

                ALCOHOL_FLAG.Enabled = true;

                ISSUE_DATE.Insertable = true;
                ISSUE_DATE.Updatable = true;

                REMARK.Insertable = true;
                REMARK.Updatable = true;

                CASH_AMOUNT.Insertable = true;
                CASH_AMOUNT.Updatable = true;

                CHECK_AMOUNT.Insertable = true;
                CHECK_AMOUNT.Updatable = true;

                NOTE_AMOUNT.Insertable = true;
                NOTE_AMOUNT.Updatable = true;

                CREDIT_AMOUNT.Insertable = true;
                CREDIT_AMOUNT.Updatable = true;

                TB_REQ_TYPE_NAME.Insertable = true;
                TB_REQ_TYPE_NAME.Updatable = true;
                
                //item//
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("WRITE_DD")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("WRITE_DD")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_NAME")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_NAME")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_SPEC")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_SPEC")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Updatable = 1;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("REMARK")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("REMARK")].Updatable = 1; 
            }            
            else
            {
                if (pModifyStatus == "N_ALL")
                {
                    SELL_REG_ID.Insertable = false;
                    SELL_REG_ID.Updatable = false;

                    SELL_BIZ_STATUS.Insertable = false;
                    SELL_BIZ_STATUS.Updatable = false;

                    SELL_BIZ_TYPE.Insertable = false;
                    SELL_BIZ_TYPE.Updatable = false;

                    SELL_USER_EMAIL.Insertable = false;
                    SELL_USER_EMAIL.Updatable = false;

                    SELL_USER_EMAIL.Insertable = false;
                    SELL_USER_EMAIL.Updatable = false;

                    BUY_TAX_REG_NO.Insertable = false;
                    BUY_TAX_REG_NO.Updatable = false;

                    BUY_REG_ID.Insertable = false;
                    BUY_REG_ID.Updatable = false;

                    BUY_VENDOR_NAME.Insertable = false;
                    BUY_VENDOR_NAME.Updatable = false;

                    BUY_VENDOR_CEO.Insertable = false;
                    BUY_VENDOR_CEO.Updatable = false;

                    BUY_ADDR1.Insertable = false;
                    BUY_ADDR1.Updatable = false;

                    BUY_ADDR2.Insertable = false;
                    BUY_ADDR2.Updatable = false;

                    BUY_BIZ_STATUS.Insertable = false;
                    BUY_BIZ_STATUS.Updatable = false;

                    BUY_BIZ_TYPE.Insertable = false;
                    BUY_BIZ_TYPE.Updatable = false;

                    BUY_USER_EMAIL.Insertable = false;
                    BUY_USER_EMAIL.Updatable = false;

                    BUY_USER2_EMAIL.Insertable = false;
                    BUY_USER2_EMAIL.Updatable = false;

                    DELAY_ISSUE_YN.Enabled = false;
                    HOMETAX_SEND_TYPE.Enabled = false;
                }

                BTN_L_INSERT.Enabled = false;
                BTN_L_DELETE.Enabled = false;

                SELL_ADDR1.Insertable = false;
                SELL_ADDR1.Updatable = false;

                SELL_ADDR2.Insertable = false;
                SELL_ADDR2.Updatable = false;

                TB_VAT_TYPE_NAME.Insertable = false;
                TB_VAT_TYPE_NAME.Updatable = false;

                ALCOHOL_FLAG.Enabled = false;

                ISSUE_DATE.Insertable = false;
                ISSUE_DATE.Updatable = false;

                REMARK.Insertable = false;
                REMARK.Updatable = false;

                CASH_AMOUNT.Insertable = false;
                CASH_AMOUNT.Updatable = false;

                CHECK_AMOUNT.Insertable = false;
                CHECK_AMOUNT.Updatable = false;

                NOTE_AMOUNT.Insertable = false;
                NOTE_AMOUNT.Updatable = false;

                CREDIT_AMOUNT.Insertable = false;
                CREDIT_AMOUNT.Updatable = false;

                TB_REQ_TYPE_NAME.Insertable = false;
                TB_REQ_TYPE_NAME.Updatable = false;

                //item//
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("WRITE_DD")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("WRITE_DD")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_NAME")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_NAME")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_SPEC")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("ITEM_SPEC")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Updatable = 0;

                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("REMARK")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("REMARK")].Updatable = 0; 
            }
            IGR_TAX_BILL_ISSUE_LINE.LastConfirmChanges(); 
            IGR_TAX_BILL_ISSUE_LINE.ResetDraw = true;
            Application.DoEvents();

            ISSUE_DATE.Focus();
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

        private void FCMF0992_DETAIL_Load(object sender, EventArgs e)
        {
            IDA_TAX_BILL_ISSUE.FillSchema();
            if (mSTATUS == "FIX")
            {
                TP_MINUS.TabVisible = true;
            }
            else
            {
                TP_MINUS.TabVisible = false;
            }

            if (mSTATUS == "INSERT")
            {
                Init_Header_Insert();
            }
            else
            {
                Search_DB();
            }

            if (mSTATUS == "FIX")
            {
                Search_DB_Minus();
            }
            BTN_GET_SUPPLIER.BringToFront();
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

        private void BTN_GET_SUPPLIER_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_GET_SUPPLIER_INFO_P.SetCommandParamValue("W_ISSUE_DATE", ISSUE_DATE.EditValue);
            IDC_GET_SUPPLIER_INFO_P.ExecuteNonQuery();
            SELL_USER_ID.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_USER_ID");
            SELL_TAX_REG_NO.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_TAX_REG_NO");
            SELL_REG_ID.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_REG_ID");
            SELL_VENDOR_NAME.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_VENDOR_NAME");
            SELL_VENDOR_CEO.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_VENDOR_CEO");
            SELL_ADDR1.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_ADDR1");
            SELL_ADDR2.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_ADDR2");
            SELL_BIZ_STATUS.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_BIZ_STATUS");
            SELL_BIZ_TYPE.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_BIZ_TYPE");
            SELL_USER_EMAIL.EditValue = IDC_GET_SUPPLIER_INFO_P.GetCommandParamValue("O_SELL_USER_EMAIL"); 
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
            //변경분 존재 체크//
            int vChg_Rec_Cnt = IDA_TAX_BILL_ISSUE.ModifiedRowCount;
            if (vChg_Rec_Cnt != 0)
            {
                //수정된 데이터 존재 :: 저장부터 처리//
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return; 
            }

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
            //1. 공급사 정보 확인//

            //
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
            string vTAX_BILL_VAT_TYPE = iConv.ISNull(TAX_BILL_VAT_TYPE.EditValue);

            if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY"))
            {
                vQTY = iConv.ISDecimaltoZero(e.NewValue, 1);
                vUNIT_PRICE = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue("UNIT_PRICE"), 0);
                vSUPPLY_AMOUNT = vQTY * vUNIT_PRICE;
                if (vTAX_BILL_VAT_TYPE == "3")
                {
                    //면세.
                }
                else
                {
                    vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);
                }
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("SUPPLY_AMOUNT", vSUPPLY_AMOUNT);
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);

                Sync_Total_Amount(e.RowIndex, vSUPPLY_AMOUNT, vVAT_AMOUNT);
            }
            else if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE"))
            {
                vQTY = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue("QTY"), 1);
                vUNIT_PRICE = iConv.ISDecimaltoZero(e.NewValue, 1);
                vSUPPLY_AMOUNT = vQTY * vUNIT_PRICE;
                if (vTAX_BILL_VAT_TYPE == "3")
                {
                    //면세.
                }
                else
                {
                    vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);
                }

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("SUPPLY_AMOUNT", vSUPPLY_AMOUNT);
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);

                Sync_Total_Amount(e.RowIndex, vSUPPLY_AMOUNT, vVAT_AMOUNT);
            }
            else if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT"))
            {
                vSUPPLY_AMOUNT = iConv.ISDecimaltoZero(e.NewValue, 0);
                if (vTAX_BILL_VAT_TYPE == "3")
                {
                    //면세.
                }
                else
                {
                    vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);
                }

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);

                Sync_Total_Amount(e.RowIndex, vSUPPLY_AMOUNT, vVAT_AMOUNT);
            } 
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
        
        private void ILA_TB_VAT_TYPE_SelectedRowData_1(object pSender)
        {
            //과세 구분에 따라 항목의 부가세 제어.
            if(iConv.ISNull(TAX_BILL_VAT_TYPE.EditValue) == "3")
            {
                //면세
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Insertable = 0;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Updatable = 0;
            }
            else 
            {
                //과세,영세
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Insertable = 1;
                IGR_TAX_BILL_ISSUE_LINE.GridAdvExColElement[IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("VAT_AMOUNT")].Updatable = 1;
            }

            Sync_VAT_Amount(VAT_TAX_TYPE.EditValue);
            Sync_Total_Amount(-1, 0, 0);
            IGR_TAX_BILL_ISSUE_LINE.ResetDraw = true;
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
                        DELAY_ISSUE_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;
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

            if (iConv.ISNull(e.Row["TAX_BILL_FIX_TYPE"]) == string.Empty)
            {
                 
            }
            else
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
        }



        #endregion

    }
}