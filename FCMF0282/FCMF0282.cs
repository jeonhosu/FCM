using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0282
{
    public partial class FCMF0282 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ----- Constructor -----

        public FCMF0282()
        {
            InitializeComponent();
        }

        public FCMF0282(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iConv.ISNull(W_TRX_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TRX_DATE_FR.Focus();
                return;
            }

            if (iConv.ISNull(W_TRX_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TRX_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(W_TRX_DATE_FR.EditValue) > Convert.ToDateTime(W_TRX_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TRX_DATE_FR.Focus();
                return;
            }

            IGR_ACCOUNTS_HISTORY.LastConfirmChanges();
            IDA_ACCOUNTS_HISTORY.OraSelectData.AcceptChanges();
            IDA_ACCOUNTS_HISTORY.Refillable = true;
            IGR_ACCOUNTS_HISTORY.ResetDraw = true;

            decimal vCMS_ACCOUNTS_HISTORY_ID = -1;
            int vCOL_IDX = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("CMS_ACCOUNTS_HISTORY_ID");
            if (IGR_ACCOUNTS_HISTORY.RowCount < 1)
            {
                vCMS_ACCOUNTS_HISTORY_ID = -1;
            }
            else
            {
                vCMS_ACCOUNTS_HISTORY_ID = iConv.ISDecimaltoZero(IGR_ACCOUNTS_HISTORY.GetCellValue("CMS_ACCOUNTS_HISTORY_ID"));                
            }
            IDA_ACCOUNTS_HISTORY.Fill();

            if (iConv.ISNull(vCMS_ACCOUNTS_HISTORY_ID) != string.Empty)
            {
                for (int i = 0; i < IGR_ACCOUNTS_HISTORY.RowCount; i++)
                {
                    if (vCMS_ACCOUNTS_HISTORY_ID == iConv.ISDecimaltoZero(IGR_ACCOUNTS_HISTORY.GetCellValue(i, vCOL_IDX)))
                    {
                        IGR_ACCOUNTS_HISTORY.CurrentCellMoveTo(i, vCOL_IDX);
                        IGR_ACCOUNTS_HISTORY.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
        }

        private void INIT_EXCHANGE_RATE(object pCURRENCY_CODE, object pEXCHANGE_DATE)
        {
            IDC_EXCHANGE_RATE.SetCommandParamValue("P_CURRENCY_CODE_FR", pCURRENCY_CODE);
            IDC_EXCHANGE_RATE.SetCommandParamValue("P_EXCHANGE_DATE", pEXCHANGE_DATE);
            IDC_EXCHANGE_RATE.ExecuteNonQuery();
            IGR_ACCOUNTS_HISTORY.SetCellValue("EXCHANGE_RATE", IDC_EXCHANGE_RATE.GetCommandParamValue("O_EXCHANGE_RATE"));
        }

        
        private void Init_Select_YN(string pStatus)
        {
            int vIDX_SELECT_YN = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SELECT_YN");
            int vIDX_SLIP_DATE = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SLIP_DATE");
            int vIDX_ACCOUNT_CODE = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("ACCOUNT_CODE");
            int vIDX_SUS_REC_ACCOUNT_CODE = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SUS_REC_ACCOUNT_CODE");
            int vIDX_SLIP_REMARK = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SLIP_REMARK");

            if (pStatus == "N")
            {
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 1;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 1;
                 
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_DATE].Insertable = 1;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_DATE].Updatable = 1;

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_ACCOUNT_CODE].Insertable = 1;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_ACCOUNT_CODE].Updatable = 1;

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SUS_REC_ACCOUNT_CODE].Insertable = 1;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SUS_REC_ACCOUNT_CODE].Updatable = 1;

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_REMARK].Insertable = 1;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_REMARK].Updatable = 1;
            }
            else
            {
                IDA_ACCOUNTS_HISTORY.Cancel();

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 0;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 0;
                 
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_DATE].Insertable = 0;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_DATE].Updatable = 0;

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_ACCOUNT_CODE].Insertable = 0;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_ACCOUNT_CODE].Updatable = 0;

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SUS_REC_ACCOUNT_CODE].Insertable = 0;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SUS_REC_ACCOUNT_CODE].Updatable = 0;

                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_REMARK].Insertable = 0;
                IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_SLIP_REMARK].Updatable = 0;
            }
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
                    if (pSub_Panel == "AP_VAT")
                    {
                        GB_AP_VAT.Left = 190;
                        GB_AP_VAT.Top = 140;

                        GB_AP_VAT.Width = 690;
                        GB_AP_VAT.Height = 305;

                        GB_AP_VAT.Border3DStyle = Border3DStyle.Bump;
                        GB_AP_VAT.BorderStyle = BorderStyle.Fixed3D;

                        GB_AP_VAT.Visible = true;
                    }

                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                igbCONDITION.Enabled = false;
                IGR_ACCOUNTS_HISTORY.Enabled = false;
                IGR_CMS_SLIP_DETAIL.Enabled = false;
                IGB_SLIP_DETAIL.Enabled = false;
            }
            else
            {
                try
                {
                    if (pSub_Panel == "ALL")
                    {
                        GB_AP_VAT.Visible = false;
                    }
                    else if (pSub_Panel == "AP_VAT")
                    {
                        GB_AP_VAT.Visible = false;
                    }

                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }

                igbCONDITION.Enabled = true;
                IGR_ACCOUNTS_HISTORY.Enabled = true;
                IGR_CMS_SLIP_DETAIL.Enabled = true;
                IGB_SLIP_DETAIL.Enabled = true;
            }
        }

        private bool Check_Sub_Panel()
        {
            if (mSUB_SHOW_FLAG == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
        
        private void GetSubForm()
        {
            ibtSUB_FORM.Visible = false;
            ACCOUNT_CLASS_YN.EditValue = null;
            ACCOUNT_CLASS_TYPE.EditValue = null;
            string vBTN_CAPTION = null;
            
            idcGET_SUB_FORM.ExecuteNonQuery();
            ACCOUNT_CLASS_YN.EditValue = idcGET_SUB_FORM.GetCommandParamValue("O_ACCOUNT_CLASS_YN");
            ACCOUNT_CLASS_TYPE.EditValue = idcGET_SUB_FORM.GetCommandParamValue("O_ACCOUNT_CLASS_TYPE");
            vBTN_CAPTION = iConv.ISNull(idcGET_SUB_FORM.GetCommandParamValue("O_BTN_CAPTION"));
            if (iConv.ISNull(ACCOUNT_CLASS_YN.EditValue, "N") == "N".ToString())
            {
                return;
            }
            
            ibtSUB_FORM.Left = 801;
            ibtSUB_FORM.Top = 120;
            ibtSUB_FORM.ButtonTextElement[0].Default = vBTN_CAPTION;
            ibtSUB_FORM.BringToFront();
            ibtSUB_FORM.Visible = true;
            ibtSUB_FORM.TabStop = true;
        }

        private void SetManagementParameter(string pManagement_Field, string pEnabled_YN, object pLookup_Type)
        {
            string mLookup_Type = iConv.ISNull(pLookup_Type);
            
            if (mLookup_Type == "VAT_TAX_TYPE")
            {//세무구분
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE",IGR_CMS_SLIP_DETAIL.GetCellValue("ACCOUNT_CODE"));
            }
            else if (mLookup_Type == "VAT_REASON")
            {//부가세사유
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("VAT_TAX_TYPE"));
            }
            else if (mLookup_Type == "DEPT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", IGR_CMS_SLIP_DETAIL.GetCellValue("BUDGET_DEPT_CODE"));
            }
            else if (mLookup_Type == "COSTCENTER".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("DEPT"));
            }
            else if (mLookup_Type == "BANK_ACCOUNT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("BANK_SITE"));
            }
            else if (mLookup_Type == "RECEIVABLE_BILL".ToString())
            {//받을어음
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", "2");
            }
            else if (mLookup_Type == "PAYABLE_BILL".ToString())
            {//지급어음
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", "1");
            }
            else if (mLookup_Type == "LC_NO".ToString())
            {
                string vSLIP_DATE = null;
                if (iConv.ISNull(IGR_CMS_SLIP_DETAIL.GetCellValue("SLIP_DATE")) != string.Empty)
                {
                    vSLIP_DATE = iDate.ISGetDate(IGR_CMS_SLIP_DETAIL.GetCellValue("SLIP_DATE")).ToShortDateString();
                }
                else if (iConv.ISNull(IGR_CMS_SLIP_DETAIL.GetCellValue("SLIP_DATE")) != string.Empty)
                {
                    vSLIP_DATE = iDate.ISGetDate(IGR_CMS_SLIP_DETAIL.GetCellValue("SLIP_DATE")).ToShortDateString();
                }
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", vSLIP_DATE);
            }
            else
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", null);
            }
            ildMANAGEMENT.SetLookupParamValue("W_ACCOUNT_CONTROL_ID", IGR_CMS_SLIP_DETAIL.GetCellValue("ACCOUNT_CONTROL_ID"));
            ildMANAGEMENT.SetLookupParamValue("W_GL_DATE", IGR_CMS_SLIP_DETAIL.GetCellValue("SLIP_DATE")); 
            ildMANAGEMENT.SetLookupParamValue("W_MANAGEMENT_FIELD", pManagement_Field);
            ildMANAGEMENT.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private object GetLookup_Type(object pLookup_Type)
        {
            if (iConv.ISNull(pLookup_Type) == string.Empty)
            {
                return null;
            }
            
            object mLookup_Value;
            if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT1.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT2.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER1_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER1_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER1.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER2_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER2_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER2.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER3_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER3_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER3.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER4_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER4_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER4.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER5_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER5_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER5.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER6_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER6_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER6.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER7_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER7_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER7.EditValue;
            }
            else if (iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER8_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_CMS_SLIP_DETAIL.CurrentRow["REFER8_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER8.EditValue;
            }
            else
            {
                mLookup_Value = null;
            }
            return mLookup_Value;
        }

        #endregion;

        #region ----- Init Component -----

        private void Set_Control_Item_Prompt()
        {
            idaCONTROL_ITEM_PROMPT.Fill();
            if (idaCONTROL_ITEM_PROMPT.CurrentRows.Count > 0)
            {
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_NAME"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_NAME"]);

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_YN"]);

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_YN"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_YN"]);

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_TYPE"]);

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_DATA_TYPE"]);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_DATA_TYPE"]);
            }
            else
            {
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_NAME", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_NAME", null);

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_YN", "F");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_YN", "F");

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_LOOKUP_YN", "N");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_LOOKUP_YN", "N");

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_LOOKUP_TYPE", null);
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_LOOKUP_TYPE", null);

                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_DATA_TYPE", "VARCHAR2");
                IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_DATA_TYPE", "VARCHAR2");
            }
        }

        private void Init_Total_GL_Amount()
        {
            decimal vSUM_APPR_AMOUNT = iConv.ISDecimaltoZero(IGR_ACCOUNTS_HISTORY.GetCellValue("BASE_TRX_AMT"), 0);
            decimal vGL_Amount = Convert.ToDecimal(0);

            foreach (DataRow vRow in IDA_CMS_SLIP_DETAIL.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Deleted)
                {
                    vGL_Amount = vGL_Amount + iConv.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                }
            }

            APPR_AMOUNT.EditValue = vSUM_APPR_AMOUNT;
            SUM_GL_AMOUNT.EditValue = vGL_Amount;
            GAP_AMOUNT.EditValue = -(System.Math.Abs(vSUM_APPR_AMOUNT - vGL_Amount)); ;
        }

        private void Init_Control_Management_Value()
        {
            IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT1_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("MANAGEMENT2_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER1_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER2_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER3_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER4_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER5_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER6_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER7_DESC", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8", null);
            IGR_CMS_SLIP_DETAIL.SetCellValue("REFER8_DESC", null);
        }

        private void Init_Control_Item_Default()
        {
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            MANAGEMENT1.NumberDecimalDigits = 0;
            MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT1.Nullable = true;
            MANAGEMENT1.Refresh();

            MANAGEMENT2.NumberDecimalDigits = 0;
            MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT2.Nullable = true;
            MANAGEMENT2.Refresh();

            REFER1.NumberDecimalDigits = 0;
            REFER1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER1.Nullable = true;
            REFER1.Refresh();

            REFER2.NumberDecimalDigits = 0;
            REFER2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER2.Nullable = true;
            REFER2.Refresh();

            REFER3.NumberDecimalDigits = 0;
            REFER3.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER3.Nullable = true;
            REFER3.Refresh();

            REFER4.NumberDecimalDigits = 0;
            REFER4.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER4.Nullable = true;
            REFER4.Refresh();

            REFER5.NumberDecimalDigits = 0;
            REFER5.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER5.Nullable = true;
            REFER5.Refresh();

            REFER6.NumberDecimalDigits = 0;
            REFER6.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER6.Nullable = true;
            REFER6.Refresh();

            REFER7.NumberDecimalDigits = 0;
            REFER7.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER7.Nullable = true;
            REFER7.Refresh();

            REFER8.NumberDecimalDigits = 0;
            REFER8.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER8.Nullable = true;
            REFER8.Refresh();
        }

        private void Init_Set_Item_Prompt(DataRow pDataRow)
        {// edit 데이터 형식, 사용여부 변경.
            if (pDataRow == null)
            {
                return;
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            string mDATA_TYPE = "VARCHAR2";

            mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
            MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT1.NumberDecimalDigits = 0;
            MANAGEMENT1.Nullable = true;
            MANAGEMENT1.ReadOnly = true;
            MANAGEMENT1.Insertable = false;
            MANAGEMENT1.Updatable = false;
            MANAGEMENT1.TabStop = false;
            if (iConv.ISNull(pDataRow["MANAGEMENT1_YN"], "F") != "F".ToString())
            {
                MANAGEMENT1.ReadOnly = false;
                MANAGEMENT1.Insertable = true;
                MANAGEMENT1.Updatable = true;
                MANAGEMENT1.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT1.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                MANAGEMENT1.Nullable = true;
            }
            MANAGEMENT1.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
            MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT2.NumberDecimalDigits = 0;
            MANAGEMENT2.Nullable = true;
            MANAGEMENT2.ReadOnly = true;
            MANAGEMENT2.Insertable = false;
            MANAGEMENT2.Updatable = false;
            MANAGEMENT2.TabStop = false;
            if (iConv.ISNull(pDataRow["MANAGEMENT2_YN"], "F") != "F".ToString())
            {
                MANAGEMENT2.ReadOnly = false;
                MANAGEMENT2.Insertable = true;
                MANAGEMENT2.Updatable = true;
                MANAGEMENT2.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT2.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                MANAGEMENT2.Nullable = true;
            }
            MANAGEMENT2.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER1_DATA_TYPE"]);
            REFER1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER1.NumberDecimalDigits = 0;
            REFER1.Nullable = true;
            REFER1.ReadOnly = true;
            REFER1.Insertable = false;
            REFER1.Updatable = false;
            REFER1.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER1_YN"], "F") != "F".ToString())
            {
                REFER1.ReadOnly = false;
                REFER1.Insertable = true;
                REFER1.Updatable = true;
                REFER1.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER1.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER1.Nullable = true;
            }
            REFER1.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER2_DATA_TYPE"]);
            REFER2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER2.NumberDecimalDigits = 0;
            REFER2.Nullable = true;
            REFER2.ReadOnly = true;
            REFER2.Insertable = false;
            REFER2.Updatable = false;
            REFER2.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER2_YN"], "F") != "F".ToString())
            {
                REFER2.ReadOnly = false;
                REFER2.Insertable = true;
                REFER2.Updatable = true;
                REFER2.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER2.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER2.Nullable = true;
            }
            REFER2.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER3_DATA_TYPE"]);
            REFER3.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER3.NumberDecimalDigits = 0;
            REFER3.Nullable = true;
            REFER3.ReadOnly = true;
            REFER3.Insertable = false;
            REFER3.Updatable = false;
            REFER3.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER3_YN"], "F") != "F".ToString())
            {
                REFER3.ReadOnly = false;
                REFER3.Insertable = true;
                REFER3.Updatable = true;
                REFER3.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER3.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER3.Nullable = true;
            }
            REFER3.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER4_DATA_TYPE"]);
            REFER4.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER4.NumberDecimalDigits = 0;
            REFER4.Nullable = true;
            REFER4.ReadOnly = true;
            REFER4.Insertable = false;
            REFER4.Updatable = false;
            REFER4.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER4_YN"], "F") != "F".ToString())
            {
                REFER4.ReadOnly = false;
                REFER4.Insertable = true;
                REFER4.Updatable = true;
                REFER4.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER4.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER4.Nullable = true;
            }
            REFER4.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER5_DATA_TYPE"]);
            REFER5.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER5.NumberDecimalDigits = 0;
            REFER5.Nullable = true;
            REFER5.ReadOnly = true;
            REFER5.Insertable = false;
            REFER5.Updatable = false;
            REFER5.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER5_YN"], "F") != "F".ToString())
            {
                REFER5.ReadOnly = false;
                REFER5.Insertable = true;
                REFER5.Updatable = true;
                REFER5.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER5.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER5.Nullable = true;
            }
            REFER5.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER6_DATA_TYPE"]);
            REFER6.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER6.NumberDecimalDigits = 0;
            REFER6.Nullable = true;
            REFER6.ReadOnly = true;
            REFER6.Insertable = false;
            REFER6.Updatable = false;
            REFER6.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER6_YN"], "F") != "F".ToString())
            {
                REFER6.ReadOnly = false;
                REFER6.Insertable = true;
                REFER6.Updatable = true;
                REFER6.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER6.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER6.Nullable = true;
            }
            REFER6.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER7_DATA_TYPE"]);
            REFER7.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER7.NumberDecimalDigits = 0;
            REFER7.Nullable = true;
            REFER7.ReadOnly = true;
            REFER7.Insertable = false;
            REFER7.Updatable = false;
            REFER7.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER7_YN"], "F") != "F".ToString())
            {
                REFER7.ReadOnly = false;
                REFER7.Insertable = true;
                REFER7.Updatable = true;
                REFER7.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER7.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER7.Nullable = true;
            }
            REFER7.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER8_DATA_TYPE"]);
            REFER8.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER8.NumberDecimalDigits = 0;
            REFER8.Nullable = true;
            REFER8.ReadOnly = true;
            REFER8.Insertable = false;
            REFER8.Updatable = false;
            REFER8.TabStop = false;
            if (iConv.ISNull(pDataRow["REFER8_YN"], "F") != "F".ToString())
            {
                REFER8.ReadOnly = false;
                REFER8.Insertable = true;
                REFER8.Updatable = true;
                REFER8.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER8.NumberDecimalDigits = 4;
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                REFER8.Nullable = true;
            }
            REFER8.Refresh();

            ///////////////////////////////////////////////////////////////////////////////////////////////////            
            if (iConv.ISNull(pDataRow["MANAGEMENT1_LOOKUP_YN"], "N") == "Y".ToString())
            {
                MANAGEMENT1.LookupAdapter = ilaMANAGEMENT1;
            }
            else
            {
                MANAGEMENT1.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["MANAGEMENT2_LOOKUP_YN"], "N") == "Y".ToString())
            {
                MANAGEMENT2.LookupAdapter = ilaMANAGEMENT2;
            }
            else
            {
                MANAGEMENT2.LookupAdapter = null;
            }
            if (iConv.ISNull(pDataRow["REFER1_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER1.LookupAdapter = ilaREFER1;
            }
            else
            {
                REFER1.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER2_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER2.LookupAdapter = ilaREFER2;
            }
            else
            {
                REFER2.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER3_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER3.LookupAdapter = ilaREFER3;
            }
            else
            {
                REFER3.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER4_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER4.LookupAdapter = ilaREFER4;
            }
            else
            {
                REFER4.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER5_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER5.LookupAdapter = ilaREFER5;
            }
            else
            {
                REFER5.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER6_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER6.LookupAdapter = ilaREFER6;
            }
            else
            {
                REFER6.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER7_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER7.LookupAdapter = ilaREFER7;
            }
            else
            {
                REFER7.LookupAdapter = null;
            }

            if (iConv.ISNull(pDataRow["REFER8_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER8.LookupAdapter = ilaREFER8;
            }
            else
            {
                REFER8.LookupAdapter = null;
            }
        }

        private void Init_Set_Item_Need(DataRow pDataRow)
        {// 관리항목 필수여부 세팅.
            if (pDataRow == null)
            {
                return;
            }

            object mDATA_VALUE;
            string mDATA_TYPE;
            string mDR_CR_YN = "N";
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            //--1
            mDATA_VALUE = MANAGEMENT1.EditValue;
            MANAGEMENT1.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["MANAGEMENT1_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_CR_YN"];
            //}
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                MANAGEMENT1.Nullable = false;
            }
            MANAGEMENT1.EditValue = mDATA_VALUE;
            MANAGEMENT1.Refresh();
            //--2
            mDATA_VALUE = MANAGEMENT2.EditValue;
            MANAGEMENT2.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["MANAGEMENT2_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                MANAGEMENT2.Nullable = false;
            }
            MANAGEMENT2.EditValue = mDATA_VALUE;
            MANAGEMENT2.Refresh();
            //--3
            mDATA_VALUE = REFER1.EditValue;
            REFER1.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER1_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER1_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER1.Nullable = false;
            }
            REFER1.EditValue = mDATA_VALUE;
            REFER1.Refresh();
            //--4
            REFER2.Nullable = true;
            mDATA_VALUE = REFER2.EditValue;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER2_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER2_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER2.Nullable = false;
            }
            REFER2.EditValue = mDATA_VALUE;
            REFER2.Refresh();
            //--5
            mDATA_VALUE = REFER3.EditValue;
            REFER3.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER3_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER3_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER3.Nullable = false;
            }
            REFER3.EditValue = mDATA_VALUE;
            REFER3.Refresh();
            //--6
            mDATA_VALUE = REFER4.EditValue;
            REFER4.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER4_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER4_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER4.Nullable = false;
            }
            REFER4.EditValue = mDATA_VALUE;
            REFER4.Refresh();
            //--7
            mDATA_VALUE = REFER5.EditValue;
            REFER5.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER5_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER5_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER5.Nullable = false;
            }
            REFER5.EditValue = mDATA_VALUE;
            REFER5.Refresh();
            //--8
            mDATA_VALUE = REFER6.EditValue;
            REFER6.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER6_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER6_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER6.Nullable = false;
            }
            REFER6.EditValue = mDATA_VALUE;
            REFER6.Refresh();
            //--9
            mDATA_VALUE = REFER7.EditValue;
            REFER7.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER7_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER7_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER7.Nullable = false;
            }
            REFER7.EditValue = mDATA_VALUE;
            REFER7.Refresh();
            //--10
            mDATA_VALUE = REFER8.EditValue;
            REFER8.Nullable = true;
            mDATA_TYPE = iConv.ISNull(pDataRow["REFER8_DATA_TYPE"]);
            mDR_CR_YN = iConv.ISNull(pDataRow["REFER8_YN"]);
            //if (iConv.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = igrSLIP_LINE.GetCellValue("REFER8_DR_YN"];
            //}
            //else if (iConv.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = igrSLIP_LINE.GetCellValue("REFER8_CR_YN"];
            //} 
            if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
            {
                REFER8.Nullable = false;
            }
            REFER8.EditValue = mDATA_VALUE;
            REFER8.Refresh();
        }

        //관리항목 LOOKUP 선택시 처리.
        private void Init_SELECT_LOOKUP(object pManagement_Type)
        {
            string mMANAGEMENT = iConv.ISNull(pManagement_Type);
        }
         
        #endregion

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

        private void XLPrinting1(string pOutput_Type)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            object vSLIP_IF_HEADER_ID = IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_IF_HEADER_ID");
            idaSLIP_HEADER.SetSelectParamValue("W_HEADER_ID", vSLIP_IF_HEADER_ID); 
            idaSLIP_HEADER.Fill();

            int vCount = idaSLIP_HEADER.CurrentRows.Count;
            if (vCount < 1)
            {
                vMessageText = "Printing is not data";
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            vMessageText = string.Format("Printing Starting", vPageTotal);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0282_001.xls";
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


                    xlPrinting.HeaderWrite(idaSLIP_HEADER, vLOCAL_DATE);

                    idaPRINT_SLIP_LINE.SetSelectParamValue("W_HEADER_ID", vSLIP_IF_HEADER_ID);
                    idaPRINT_SLIP_LINE.Fill();
                    vCountRow = idaPRINT_SLIP_LINE.CurrentRows.Count;
                    if (vCountRow > 0)
                    {
                        vPageNumber = xlPrinting.LineWrite(idaPRINT_SLIP_LINE);
                    }

                    if (pOutput_Type == "PRINT")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PreView(1, vPageNumber);

                    }
                    else if (pOutput_Type == "EXCEL")
                    {
                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명
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
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ACCOUNTS_HISTORY.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ACCOUNTS_HISTORY.IsFocused)
                    {
                        IDA_ACCOUNTS_HISTORY.Cancel();
                        IDA_CMS_SLIP_DETAIL.Cancel();
                    }
                    else
                    {
                        IDA_CMS_SLIP_DETAIL.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting1("EXCEL");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0282_Load(object sender, EventArgs e)
        {
            V_RB_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SLIP_IF_FLAG.EditValue = V_RB_NO.RadioCheckedString;

            W_TRX_DATE_FR.EditValue = iDate.ISMonth_1st(iDate.ISGetDate());
            W_TRX_DATE_TO.EditValue = iDate.ISGetDate();

            V_AUTO_APPR_REQ.BringToFront();
            BTN_REQ_OK.BringToFront();
            BTN_REQ_CANCEL.BringToFront();

            //서브판넬 
            Init_Sub_Panel(false, "ALL");

            // 콤퍼넌트 동기화.
            //Init_Currency_Code();
            ibtSUB_FORM.Visible = false;

        }

        private void FCMF0282_Shown(object sender, EventArgs e)
        {
            IDA_ACCOUNTS_HISTORY.FillSchema();
        }

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            if (V_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_ALL.RadioCheckedString; 
                Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
            }
        }

        private void V_RB_YES_Click(object sender, EventArgs e)
        {
            if (V_RB_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_YES.RadioCheckedString;
                Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
            }
        }

        private void V_RB_NO_Click(object sender, EventArgs e)
        {
            if (V_RB_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_NO.RadioCheckedString;
                Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
            }
        }

        private void IGR_ACCOUNTS_HISTORY_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SLIP_DATE"))
            {
                INIT_EXCHANGE_RATE(IGR_ACCOUNTS_HISTORY.GetCellValue("CURRENCY_CODE"), e.NewValue);
            }

            int vIDX_EXCHANGE_RATE = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("EXCHANGE_RATE");
            if (vIDX_EXCHANGE_RATE == e.ColIndex)
            {
                int vIDX_DETAIL_EXCHANGE_RATE = IGR_CMS_SLIP_DETAIL.GetColumnToIndex("EXCHANGE_RATE");
                int vIDX_GL_CURRENCY_AMOUNT = IGR_CMS_SLIP_DETAIL.GetColumnToIndex("GL_CURRENCY_AMOUNT");
                int vIDX_GL_AMOUNT = IGR_CMS_SLIP_DETAIL.GetColumnToIndex("GL_AMOUNT");
                for (int r = 0; r < IGR_CMS_SLIP_DETAIL.RowCount; r++)
                {
                    decimal vGL_AMOUNT = iConv.ISDecimaltoZero(e.NewValue, 1) *
                                             iConv.ISDecimaltoZero(IGR_CMS_SLIP_DETAIL.GetCellValue(r, vIDX_GL_CURRENCY_AMOUNT), 0);
                    IGR_CMS_SLIP_DETAIL.SetCellValue(r, vIDX_GL_AMOUNT, vGL_AMOUNT);
                }
            }
        }

        private void IGR_ACCOUNTS_HISTORY_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_EXCHANGE_RATE = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("EXCHANGE_RATE");
            if (vIDX_EXCHANGE_RATE == e.ColIndex)
            {
                decimal vBASE_TRX_AMOUNT = iConv.ISDecimaltoZero(e.NewValue, 1) *
                                             iConv.ISDecimaltoZero(IGR_ACCOUNTS_HISTORY.GetCellValue("CURR_TRX_AMT"), 0);
                IGR_ACCOUNTS_HISTORY.SetCellValue("BASE_TRX_AMT", vBASE_TRX_AMOUNT);
            }
        }

        private void BTN_CMS_SLIP_DETAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vCMS_ACCOUNTS_HISTORY_ID = IGR_ACCOUNTS_HISTORY.GetCellValue("CMS_ACCOUNTS_HISTORY_ID");
            if (iConv.ISNull(vCMS_ACCOUNTS_HISTORY_ID) == string.Empty)
            {
                return;
            }

            if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_DATE")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("ACCOUNT_CONTROL_ID")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10413"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SUS_REC_ACCOUNT_CONTROL_ID")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10413"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_REMARK")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10530"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_ACCOUNTS_HISTORY.Update();

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_CMS_SLIP_DETAIL.SetCommandParamValue("W_CMS_ACCOUNTS_HISTORY_ID", vCMS_ACCOUNTS_HISTORY_ID);
            IDC_SET_CMS_SLIP_DETAIL.ExecuteNonQuery(); 
            string vSTATUS = iConv.ISNull(IDC_SET_CMS_SLIP_DETAIL.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_CMS_SLIP_DETAIL.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_CMS_SLIP_DETAIL.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_CMS_SLIP_DETAIL.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_CMS_SLIP_DETAIL.Fill();
        }

        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            //변경사항 저장//
            IDA_ACCOUNTS_HISTORY.Update();        

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            object vSLIP_IF_HEADER_ID = 0;
            int vIDX_SELECT_YN = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SELECT_YN");
            int vIDX_CMS_ACCOUNTS_HISTORY_ID = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("CMS_ACCOUNTS_HISTORY_ID");
            for (int r = 0; r < IGR_ACCOUNTS_HISTORY.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue(r, vIDX_SELECT_YN)))
                {
                    IGR_ACCOUNTS_HISTORY.CurrentCellMoveTo(r, vIDX_SELECT_YN);
                    IGR_ACCOUNTS_HISTORY.CurrentCellActivate(r, vIDX_SELECT_YN);

                    if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_DATE")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SUS_REC_ACCOUNT_CONTROL_ID")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10413"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SUS_REC_ACCOUNT_CONTROL_ID")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10413"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_REMARK")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10530"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    } 

                    IDC_SET_SLIP_BUDGET_TRANSFER.SetCommandParamValue("P_SELECT_YN", IGR_ACCOUNTS_HISTORY.GetCellValue(r, vIDX_SELECT_YN));
                    IDC_SET_SLIP_BUDGET_TRANSFER.SetCommandParamValue("P_CMS_ACCOUNTS_HISTORY_ID", IGR_ACCOUNTS_HISTORY.GetCellValue(r, vIDX_CMS_ACCOUNTS_HISTORY_ID));
                    IDC_SET_SLIP_BUDGET_TRANSFER.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_SLIP_BUDGET_TRANSFER.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_SLIP_BUDGET_TRANSFER.GetCommandParamValue("O_MESSAGE"));
                    vSLIP_IF_HEADER_ID = IDC_SET_SLIP_BUDGET_TRANSFER.GetCommandParamValue("O_SLIP_IF_HEADER_ID");

                    if (IDC_SET_SLIP_BUDGET_TRANSFER.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_SET_SLIP_BUDGET_TRANSFER.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (V_AUTO_APPR_REQ.CheckedState == ISUtil.Enum.CheckedState.Checked)
                    {
                        IDC_SET_APPROVAL_REQUEST_OK.SetCommandParamValue("P_SLIP_IF_HEADER_ID", vSLIP_IF_HEADER_ID);
                        IDC_SET_APPROVAL_REQUEST_OK.ExecuteNonQuery();

                        vSTATUS = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_OK.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_OK.GetCommandParamValue("O_MESSAGE"));

                        if (IDC_SET_APPROVAL_REQUEST_OK.ExcuteError)
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            Application.DoEvents();

                            MessageBoxAdv.Show(IDC_SET_APPROVAL_REQUEST_OK.ExcuteErrorMsg, "Appr.Req-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if (vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            Application.DoEvents();

                            MessageBoxAdv.Show(vMESSAGE, "Appr.Req-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    IGR_ACCOUNTS_HISTORY.SetCellValue(r, vIDX_SELECT_YN, "N");

                    IGR_ACCOUNTS_HISTORY.LastConfirmChanges();
                    IDA_ACCOUNTS_HISTORY.OraSelectData.AcceptChanges();
                    IDA_ACCOUNTS_HISTORY.Refillable = true;
                }
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            //다시 조회//
            Search_DB();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
     
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            int vIDX_SELECT_YN = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SELECT_YN");
            int vIDX_SLIP_IF_HEADER_ID = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("SLIP_IF_HEADER_ID");
            for (int r = 0; r < IGR_ACCOUNTS_HISTORY.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue(r, vIDX_SELECT_YN)))
                {
                    IGR_ACCOUNTS_HISTORY.CurrentCellMoveTo(r, vIDX_SELECT_YN);
                    IGR_ACCOUNTS_HISTORY.CurrentCellActivate(r, vIDX_SELECT_YN);

                    if (iConv.ISNull(IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_IF_HEADER_ID")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10128"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    IDC_CANCEL_SLIP_BUDGET_TRANSFER.SetCommandParamValue("P_SELECT_YN", IGR_ACCOUNTS_HISTORY.GetCellValue(r, vIDX_SELECT_YN));
                    IDC_CANCEL_SLIP_BUDGET_TRANSFER.SetCommandParamValue("P_SLIP_IF_HEADER_ID", IGR_ACCOUNTS_HISTORY.GetCellValue(r, vIDX_SLIP_IF_HEADER_ID));
                    IDC_CANCEL_SLIP_BUDGET_TRANSFER.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_CANCEL_SLIP_BUDGET_TRANSFER.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_CANCEL_SLIP_BUDGET_TRANSFER.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_CANCEL_SLIP_BUDGET_TRANSFER.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_CANCEL_SLIP_BUDGET_TRANSFER.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    IGR_ACCOUNTS_HISTORY.SetCellValue(r, vIDX_SELECT_YN, "N");
                    
                    IGR_ACCOUNTS_HISTORY.LastConfirmChanges();
                    IDA_ACCOUNTS_HISTORY.OraSelectData.AcceptChanges();
                    IDA_ACCOUNTS_HISTORY.Refillable = true;
                }
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            //다시 조회//
            Search_DB();
        }

        private void BTN_REQ_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            object vSLIP_IF_HEADER_ID = IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_IF_HEADER_ID");
            IDC_SET_APPROVAL_REQUEST_OK.SetCommandParamValue("P_SLIP_IF_HEADER_ID", vSLIP_IF_HEADER_ID);
            IDC_SET_APPROVAL_REQUEST_OK.ExecuteNonQuery();

            vSTATUS = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_OK.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_OK.GetCommandParamValue("O_MESSAGE"));

            if (IDC_SET_APPROVAL_REQUEST_OK.ExcuteError)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(IDC_SET_APPROVAL_REQUEST_OK.ExcuteErrorMsg, "Appr.Req-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(vMESSAGE, "Appr.Req-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Search_DB();
        }

        private void BTN_REQ_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            object vSLIP_IF_HEADER_ID = IGR_ACCOUNTS_HISTORY.GetCellValue("SLIP_IF_HEADER_ID");
            IDC_SET_APPROVAL_REQUEST_CANCEL.SetCommandParamValue("P_SLIP_IF_HEADER_ID", vSLIP_IF_HEADER_ID);
            IDC_SET_APPROVAL_REQUEST_CANCEL.ExecuteNonQuery();

            vSTATUS = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_CANCEL.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_CANCEL.GetCommandParamValue("O_MESSAGE"));

            if (IDC_SET_APPROVAL_REQUEST_CANCEL.ExcuteError)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(IDC_SET_APPROVAL_REQUEST_CANCEL.ExcuteErrorMsg, "Appr.Req-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(vMESSAGE, "Appr.Req-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Search_DB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_BANK_ACCOUNT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_BANK_SITE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BANK_SITE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_ACCOUNT_CONTROL_FROM_TO_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_FROM_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {            
            SetManagementParameter("MANAGEMENT1_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT2_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
        }

        private void ilaREFER1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER1_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER1_LOOKUP_TYPE"));
        }

        private void ilaREFER2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER2_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER2_LOOKUP_TYPE"));
        }

        private void ilaREFER3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER3_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER3_LOOKUP_TYPE"));
        }

        private void ilaREFER4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER4_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER4_LOOKUP_TYPE"));
        }

        private void ilaREFER5_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER5_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER5_LOOKUP_TYPE"));
        }

        private void ilaREFER6_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER6_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER6_LOOKUP_TYPE"));
        }

        private void ilaREFER7_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER7_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER7_LOOKUP_TYPE"));
        }

        private void ilaREFER8_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER8_ID", "Y", IGR_CMS_SLIP_DETAIL.GetCellValue("REFER8_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT1_SelectedRowData(object pSender)
        {// 관리항목1 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT1");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("MANAGEMENT1", "VAT_TAX_TYPE", "VAT_REASON", null, null); 
        }

        private void ilaMANAGEMENT2_SelectedRowData(object pSender)
        {// 관리항목2 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT2");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("MANAGEMENT2", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER1_SelectedRowData(object pSender)
        {// 관리항목3 선택시 적용.
            Init_SELECT_LOOKUP("REFER1");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER1", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER2_SelectedRowData(object pSender)
        {// 관리항목4 선택시 적용.
            Init_SELECT_LOOKUP("REFER2");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER2", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER3_SelectedRowData(object pSender)
        {// 관리항목5 선택시 적용.
            Init_SELECT_LOOKUP("REFER3");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER3", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER4_SelectedRowData(object pSender)
        {// 관리항목6 선택시 적용.
            Init_SELECT_LOOKUP("REFER4");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER4", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER5_SelectedRowData(object pSender)
        {// 관리항목7 선택시 적용.
            Init_SELECT_LOOKUP("REFER5");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER5", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER6_SelectedRowData(object pSender)
        {// 관리항목8 선택시 적용.
            Init_SELECT_LOOKUP("REFER6");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER6", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER7_SelectedRowData(object pSender)
        {// 관리항목9 선택시 적용.
            Init_SELECT_LOOKUP("REFER7");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER7", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void ilaREFER8_SelectedRowData(object pSender)
        {// 관리항목10 선택시 적용.
            Init_SELECT_LOOKUP("REFER8");

            ////부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            //Set_Ref_Management_Value("REFER8", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_ACCOUNTS_HISTORY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["SLIP_DATE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10413"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["SUS_REC_ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10413"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["SLIP_REMARK"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10530"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void IDA_ACCOUNTS_HISTORY_UpdateCompleted(object pSender)
        {
            IDA_CMS_SLIP_DETAIL.Fill();
        }

        private void IDA_ACCOUNTS_HISTORY_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            int vIDX_EXCHANGE_RATE = IGR_ACCOUNTS_HISTORY.GetColumnToIndex("EXCHANGE_RATE");
            object vUpdatable = 0;
            if (pBindingManager.DataRow == null)
            {
                vUpdatable = 0;
            }
            else
            {
                if (iConv.ISNull(pBindingManager.DataRow["CURRENCY_CODE"]) == iConv.ISNull(pBindingManager.DataRow["BASE_CURRENCY_CODE"]))
                {
                    vUpdatable = 0;
                }
                else
                {
                    vUpdatable = 1;
                }
            }
            IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_EXCHANGE_RATE].Insertable = vUpdatable;
            IGR_ACCOUNTS_HISTORY.GridAdvExColElement[vIDX_EXCHANGE_RATE].Updatable = vUpdatable;
             
            Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
            Init_Total_GL_Amount();
        }

        private void IDA_CMS_SLIP_DETAIL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {// 계정과목.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["MANAGEMENT1"]) == string.Empty && iConv.ISNull(e.Row["MANAGEMENT1_YN"], "N") == "Y".ToString())
            {// 관리항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["MANAGEMENT1_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["MANAGEMENT2"]) == string.Empty && iConv.ISNull(e.Row["MANAGEMENT2_YN"], "N") == "Y".ToString())
            {// 관리항목2 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["MANAGEMENT2_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER1"]) == string.Empty && iConv.ISNull(e.Row["REFER1_YN"], "N") == "Y".ToString())
            {// 참고항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER1_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER2"]) == string.Empty && iConv.ISNull(e.Row["REFER2_YN"], "N") == "Y".ToString())
            {// 참고항목2 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER2_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER3"]) == string.Empty && iConv.ISNull(e.Row["REFER3_YN"], "N") == "Y".ToString())
            {// 참고항목3 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER3_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER4"]) == string.Empty && iConv.ISNull(e.Row["REFER4_YN"], "N") == "Y".ToString())
            {// 참고항목4 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER4_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER5"]) == string.Empty && iConv.ISNull(e.Row["REFER5_YN"], "N") == "Y".ToString())
            {// 참고항목5 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER5_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER6"]) == string.Empty && iConv.ISNull(e.Row["REFER6_YN"], "N") == "Y".ToString())
            {// 참고항목6 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER6_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER7"]) == string.Empty && iConv.ISNull(e.Row["REFER7_YN"], "N") == "Y".ToString())
            {// 참고항목7 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER7_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REFER8"]) == string.Empty && iConv.ISNull(e.Row["REFER8_YN"], "N") == "Y".ToString())
            {// 참고항목8 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER8_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_SLIP_LINE_BUDGET_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            GetSubForm();
        }

        private void IDA_CMS_SLIP_DETAIL_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Set_Item_Prompt(pBindingManager.DataRow);
            Init_Set_Item_Need(pBindingManager.DataRow);
        }

        private void IDA_CMS_SLIP_DETAIL_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            GetSubForm();
            Init_Total_GL_Amount();
        }
        #endregion



    }
}