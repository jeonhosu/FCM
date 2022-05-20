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

namespace FCMF0525
{
    public partial class FCMF0525 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mCurrency_Code;
        object mBudget_Control_YN;

        object mOffset_Account_Control_ID;
        object mOffset_Account_Code;
        object mOffset_Account_Desc;
        string mOffset_Account_DR_CR;
        object mOffset_Account_DR_CR_Name; 
        #endregion;

        #region ----- Constructor -----

        public FCMF0525()
        {
            InitializeComponent();
        }

        public FCMF0525(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void GetAccountBook()
        {
            idcACCOUNT_BOOK.ExecuteNonQuery();
            mSession_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_SESSION_ID");
            mAccount_Book_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = idcACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = idcACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN"); 
            if(iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_SLIP_REMARK_FLAG")) == "Y")
            {
                REMARK.LookupAdapter = ILA_SLIP_REMARK;
            }
            else
            {
                REMARK.LookupAdapter = null;
            }

            //상쇄계정 존재여부.
            IDC_OFFSET_ACCOUNT_P.ExecuteNonQuery();
            mOffset_Account_Control_ID = IDC_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_CONTROL_ID");
            mOffset_Account_Code = IDC_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_CODE");
            mOffset_Account_Desc = IDC_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_DESC");
            mOffset_Account_DR_CR = iString.ISNull(IDC_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_DR_CR"));
            mOffset_Account_DR_CR_Name = IDC_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_DR_CR_DESC");

            if (iString.ISNull(mOffset_Account_Control_ID) == string.Empty)
            {
                BTN_OFFSET_ACCOUNT.Visible = false;
            }
            else
            {
                BTN_OFFSET_ACCOUNT.Visible = true;
            }
        }

        private void Search_DB()
        {
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            if (TB_BATCH.SelectedTab.TabIndex == 2)
            {
                Search_DB_DETAIL(BATCH_HEADER_ID.EditValue);
            }
            else
            {
                if (iString.ISNull(BATCH_DATE_FR_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    BATCH_DATE_FR_0.Focus();
                    return;
                }

                if (iString.ISNull(BATCH_DATE_TO_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    BATCH_DATE_TO_0.Focus();
                    return;
                }

                if (Convert.ToDateTime(BATCH_DATE_FR_0.EditValue) > Convert.ToDateTime(BATCH_DATE_TO_0.EditValue))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    BATCH_DATE_FR_0.Focus();
                    return;
                }

                string vGL_NUM = iString.ISNull(IGR_BATCH_LIST.GetCellValue("GL_NUM"));
                int vCOL_IDX = IGR_BATCH_LIST.GetColumnToIndex("GL_NUM");
                IDA_BATCH_LIST.Fill();
                if (iString.ISNull(vGL_NUM) != string.Empty)
                {
                    for (int i = 0; i < IGR_BATCH_LIST.RowCount; i++)
                    {
                        if (vGL_NUM == iString.ISNull(IGR_BATCH_LIST.GetCellValue(i, vCOL_IDX)))
                        {
                            IGR_BATCH_LIST.CurrentCellMoveTo(i, vCOL_IDX);
                            IGR_BATCH_LIST.CurrentCellActivate(i, vCOL_IDX);
                            return;
                        }
                    }
                }
            }
        }

        private void Search_DB_DETAIL(object pSLIP_HEADER_ID)
        {
            if (iString.ISNull(pSLIP_HEADER_ID) != string.Empty)
            {
                SLIP_QUERY_STATUS.EditValue = "QUERY";
                TB_BATCH.SelectedIndex = 1;
                TB_BATCH.SelectedTab.Focus();
                IDA_BATCH_HEADER.SetSelectParamValue("W_BATCH_HEADER_ID", pSLIP_HEADER_ID);
                try
                {
                    IDA_BATCH_HEADER.Fill();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }
                IDA_BATCH_LINE.OraSelectData.AcceptChanges();
                IDA_BATCH_LINE.Refillable = true;
                IDA_BATCH_HEADER.OraSelectData.AcceptChanges();
                IDA_BATCH_HEADER.Refillable = true;
            }
        }

        private void Set_Control_Item_Prompt()
        {
            idaCONTROL_ITEM_PROMPT.Fill();
            if (idaCONTROL_ITEM_PROMPT.CurrentRows.Count > 0)
            {
                igrSLIP_LINE.SetCellValue("MANAGEMENT1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_NAME"]);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER3_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER4_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER5_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER6_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER7_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_NAME"]);
                igrSLIP_LINE.SetCellValue("REFER8_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_NAME"]);
                
                igrSLIP_LINE.SetCellValue("MANAGEMENT1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_YN"]);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_YN"]);
                igrSLIP_LINE.SetCellValue("REFER1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_YN"]);
                igrSLIP_LINE.SetCellValue("REFER2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_YN"]);
                igrSLIP_LINE.SetCellValue("REFER3_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_YN"]);
                igrSLIP_LINE.SetCellValue("REFER4_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_YN"]);
                igrSLIP_LINE.SetCellValue("REFER5_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_YN"]);
                igrSLIP_LINE.SetCellValue("REFER6_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_YN"]);
                igrSLIP_LINE.SetCellValue("REFER7_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_YN"]);
                igrSLIP_LINE.SetCellValue("REFER8_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_YN"]);

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER3_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER4_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER5_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER6_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER7_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_YN"]);
                igrSLIP_LINE.SetCellValue("REFER8_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_YN"]);
                
                igrSLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER3_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER4_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER5_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER6_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER7_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER8_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_TYPE"]);
                
                igrSLIP_LINE.SetCellValue("MANAGEMENT1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER3_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER4_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER5_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER6_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER7_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_DATA_TYPE"]);
                igrSLIP_LINE.SetCellValue("REFER8_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_DATA_TYPE"]);
            }
            else
            {
                igrSLIP_LINE.SetCellValue("MANAGEMENT1_NAME", null);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER1_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER2_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER3_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER4_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER5_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER6_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER7_NAME", null);
                igrSLIP_LINE.SetCellValue("REFER8_NAME", null);

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_YN", "F");
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER1_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER2_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER3_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER4_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER5_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER6_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER7_YN", "F");
                igrSLIP_LINE.SetCellValue("REFER8_YN", "F");

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER1_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER2_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER3_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER4_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER5_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER6_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER7_LOOKUP_YN", "N");
                igrSLIP_LINE.SetCellValue("REFER8_LOOKUP_YN", "N");

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER1_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER2_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER3_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER4_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER5_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER6_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER7_LOOKUP_TYPE", null);
                igrSLIP_LINE.SetCellValue("REFER8_LOOKUP_TYPE", null);

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER1_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER2_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER3_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER4_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER5_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER6_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER7_DATA_TYPE", "VARCHAR2");
                igrSLIP_LINE.SetCellValue("REFER8_DATA_TYPE", "VARCHAR2");
            }
        }

        //private void Set_SlipLIne_ControlItem()
        //{
        //    igrSLIP_LINE.SetCellValue("MANAGEMENT1_NAME", iString.ISNull(MANAGEMENT1_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("MANAGEMENT2_NAME", iString.ISNull(MANAGEMENT2_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER1_NAME", iString.ISNull(REFER1_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER2_NAME", iString.ISNull(REFER2_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER3_NAME", iString.ISNull(REFER3_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER4_NAME", iString.ISNull(REFER4_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER5_NAME", iString.ISNull(REFER5_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER6_NAME", iString.ISNull(REFER6_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER7_NAME", iString.ISNull(REFER7_PROMPT.EditValue));
        //    igrSLIP_LINE.SetCellValue("REFER8_NAME", iString.ISNull(REFER8_PROMPT.EditValue));
        //    //igrSLIP_LINE.SetCellValue("REFER_RATE_NAME", iString.ISNull(REFER_RATE_PROMPT.EditValue));
        //    //igrSLIP_LINE.SetCellValue("REFER_AMOUNT_NAME", iString.ISNull(REFER_AMOUNT_PROMPT.EditValue));
        //    //igrSLIP_LINE.SetCellValue("REFER_DATE1_NAME", iString.ISNull(REFER_DATE1_PROMPT.EditValue));
        //    //igrSLIP_LINE.SetCellValue("REFER_DATE2_NAME", iString.ISNull(REFER_DATE2_PROMPT.EditValue));

        //    igrSLIP_LINE.SetCellValue("MANAGEMENT1_YN", iString.ISNull(MANAGEMENT1_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("MANAGEMENT2_YN", iString.ISNull(MANAGEMENT2_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER1_YN", iString.ISNull(REFER1_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER2_YN", iString.ISNull(REFER2_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER3_YN", iString.ISNull(REFER3_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER4_YN", iString.ISNull(REFER4_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER5_YN", iString.ISNull(REFER5_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER6_YN", iString.ISNull(REFER6_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER7_YN", iString.ISNull(REFER7_YN.EditValue, "F"));
        //    igrSLIP_LINE.SetCellValue("REFER8_YN", iString.ISNull(REFER8_YN.EditValue, "F"));
        //    //igrSLIP_LINE.SetCellValue("REFER_RATE_YN", iString.ISNull(REFER_RATE_YN.EditValue, "F"));
        //    //igrSLIP_LINE.SetCellValue("REFER_AMOUNT_YN", iString.ISNull(REFER_AMOUNT_YN.EditValue, "F"));
        //    //igrSLIP_LINE.SetCellValue("REFER_DATE1_YN", iString.ISNull(REFER_DATE1_YN.EditValue, "F"));
        //    //igrSLIP_LINE.SetCellValue("REFER_DATE2_YN", iString.ISNull(REFER_DATE2_YN.EditValue, "F"));
        //    //igrSLIP_LINE.SetCellValue("VOUCH_YN", iString.ISNull(VOUCH_YN.EditValue, "F"));

        //    igrSLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_YN", iString.ISNull(MANAGEMENT1_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_YN", iString.ISNull(MANAGEMENT2_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER1_LOOKUP_YN", iString.ISNull(REFER1_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER2_LOOKUP_YN", iString.ISNull(REFER2_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER3_LOOKUP_YN", iString.ISNull(REFER3_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER4_LOOKUP_YN", iString.ISNull(REFER4_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER5_LOOKUP_YN", iString.ISNull(REFER5_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER6_LOOKUP_YN", iString.ISNull(REFER6_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER7_LOOKUP_YN", iString.ISNull(REFER7_LOOKUP_YN.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER8_LOOKUP_YN", iString.ISNull(REFER8_LOOKUP_YN.EditValue, "N"));

        //    igrSLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", MANAGEMENT1_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", MANAGEMENT2_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER1_LOOKUP_TYPE", REFER1_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER2_LOOKUP_TYPE", REFER2_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER3_LOOKUP_TYPE", REFER3_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER4_LOOKUP_TYPE", REFER4_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER5_LOOKUP_TYPE", REFER5_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER6_LOOKUP_TYPE", REFER6_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER7_LOOKUP_TYPE", REFER7_LOOKUP_TYPE.EditValue);
        //    igrSLIP_LINE.SetCellValue("REFER8_LOOKUP_TYPE", REFER8_LOOKUP_TYPE.EditValue);

        //    igrSLIP_LINE.SetCellValue("MANAGEMENT1_DATA_TYPE", iString.ISNull(MANAGEMENT1_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("MANAGEMENT2_DATA_TYPE", iString.ISNull(MANAGEMENT2_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER1_DATA_TYPE", iString.ISNull(REFER1_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER2_DATA_TYPE", iString.ISNull(REFER2_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER3_DATA_TYPE", iString.ISNull(REFER3_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER4_DATA_TYPE", iString.ISNull(REFER4_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER5_DATA_TYPE", iString.ISNull(REFER5_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER6_DATA_TYPE", iString.ISNull(REFER6_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER7_DATA_TYPE", iString.ISNull(REFER7_DATA_TYPE.EditValue, "N"));
        //    igrSLIP_LINE.SetCellValue("REFER8_DATA_TYPE", iString.ISNull(REFER8_DATA_TYPE.EditValue, "N"));
        //}

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter_W(string pGroup_Code, string pWhere, string pEnabled_YN)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private string Get_Lookup_Type(string pManagement)
        {
            string vLookup_Type = "";
            if (pManagement == "MANAGEMENT1")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
            }
            else if (pManagement == "MANAGEMENT2")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER1")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER2")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER3")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER4")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER5")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER6")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER7")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
            }
            else if (pManagement == "REFER8")
            {
                vLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
            }
            return vLookup_Type;
        }

        private void SetManagementParameter(string pManagement_Field, string pEnabled_YN, object pLookup_Type)
        {
            if (iString.ISNull(pLookup_Type) == "DEPT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", DEPT_CODE.EditValue);
            }
            else if (iString.ISNull(pLookup_Type) == "COSTCENTER".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("DEPT"));
            }
            else if (iString.ISNull(pLookup_Type) == "BANK_ACCOUNT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("BANK_SITE"));
            }
            else if (iString.ISNull(pLookup_Type) == "RECEIVABLE_BILL".ToString())
            {//받을어음
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", "2");
            }
            else if (iString.ISNull(pLookup_Type) == "PAYABLE_BILL".ToString())
            {//지급어음
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", "1");
            }
            else if (iString.ISNull(pLookup_Type) == "LC_NO".ToString())
            {
                string vGL_DATE = null;
                if (iString.ISNull(GL_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = GL_DATE.DateTimeValue.ToShortDateString();
                }
                else if (iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = BATCH_DATE.DateTimeValue.ToShortDateString();
                }
                else
                {
                    vGL_DATE = null;
                }
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", vGL_DATE);
            }
            
            ildMANAGEMENT.SetLookupParamValue("W_MANAGEMENT_FIELD", pManagement_Field);
            ildMANAGEMENT.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private object GetLookup_Type(object pLookup_Type)
        {
            if (iString.ISNull(pLookup_Type) == string.Empty)
            {
                return null;
            }
            object mLookup_Value;
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT1.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT2.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER1.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER2.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER3.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER4.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER5.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER6.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER7.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER8.EditValue;
            }
            else
            {
                mLookup_Value = null;
            }
            return mLookup_Value;
        }

        private void GetBatchNum()
        {
            if (iString.ISNull(DOCUMENT_TYPE.EditValue) == string.Empty)
            {
                return;
            }
            idcSLIP_NUM.SetCommandParamValue("W_DOCUMENT_TYPE", DOCUMENT_TYPE.EditValue);
            idcSLIP_NUM.ExecuteNonQuery();
            BATCH_NUM.EditValue = idcSLIP_NUM.GetCommandParamValue("O_DOCUMENT_NUM"); 
        }

        private void GetSubForm()
        {
            ibtSUB_FORM.Visible = false;
            ACCOUNT_CLASS_YN.EditValue = null;
            ACCOUNT_CLASS_TYPE.EditValue = null;
            string vBTN_CAPTION = null;

            if (iString.ISNull(ACCOUNT_CONTROL_ID.EditValue) == string.Empty || iString.ISNull(ACCOUNT_DR_CR.EditValue) == string.Empty)
            {
                return;
            }
            idcGET_SUB_FORM.ExecuteNonQuery();
            ACCOUNT_CLASS_YN.EditValue = idcGET_SUB_FORM.GetCommandParamValue("O_ACCOUNT_CLASS_YN");
            ACCOUNT_CLASS_TYPE.EditValue = idcGET_SUB_FORM.GetCommandParamValue("O_ACCOUNT_CLASS_TYPE");
            vBTN_CAPTION = iString.ISNull(idcGET_SUB_FORM.GetCommandParamValue("O_BTN_CAPTION"));
            if (iString.ISNull(ACCOUNT_CLASS_YN.EditValue, "N") == "N".ToString())
            {
                return;
            }
            ibtSUB_FORM.Left = 780;
            ibtSUB_FORM.Top = 75;
            ibtSUB_FORM.ButtonTextElement[0].Default = vBTN_CAPTION;
            ibtSUB_FORM.BringToFront();
            ibtSUB_FORM.Visible = true;
            ibtSUB_FORM.TabStop = true;
        }

        private void Set_Management_Value(string pLookup_Type, object pManagement_Value, object pManagement_Desc)
        {
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                MANAGEMENT1.EditValue = pManagement_Value;
                MANAGEMENT1_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                MANAGEMENT2.EditValue = pManagement_Value;
                MANAGEMENT2_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                REFER1.EditValue = pManagement_Value;
                REFER1_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                REFER2.EditValue = pManagement_Value;
                REFER2_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                REFER3.EditValue = pManagement_Value;
                REFER3_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                REFER4.EditValue = pManagement_Value;
                REFER4_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                REFER5.EditValue = pManagement_Value;
                REFER5_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                REFER6.EditValue = pManagement_Value;
                REFER6_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                REFER7.EditValue = pManagement_Value;
                REFER7_DESC.EditValue = pManagement_Desc;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                REFER8.EditValue = pManagement_Value;
                REFER8_DESC.EditValue = pManagement_Desc;
            }
        }

        private object Get_Management_Value(string pLookup_Type)
        {
            object vManagement_Value = null;
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                vManagement_Value = MANAGEMENT1.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                vManagement_Value = MANAGEMENT2.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                vManagement_Value = REFER1.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                vManagement_Value = REFER2.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                vManagement_Value = REFER3.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                vManagement_Value = REFER4.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                vManagement_Value = REFER5.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                vManagement_Value = REFER6.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                vManagement_Value = REFER7.EditValue;
            }
            else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                vManagement_Value = REFER8.EditValue;
            }
            return vManagement_Value;
        }

        private void Set_Validate_Management_Value(string pManagement, string pLookup_Type, string pRef_Lookup_Type, object pManagement_Value, object pManagement_Desc)
        {
            if (pManagement == "MANAGEMENT1" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "MANAGEMENT2" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER1" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER2" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER3" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER4" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER5" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER6" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER7" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER8" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
        }

        private void Set_Ref_Management_Value(string pManagement, string pLookup_Type, string pRef_Lookup_Type, object pManagement_Value)
        {
            Set_Ref_Management_Value(pManagement, pLookup_Type, pRef_Lookup_Type, pManagement_Value, null, null, null);
        }

        private void Set_Ref_Management_Value(string pManagement, string pLookup_Type, string pRef_Lookup_Type, object pManagement_Value, object pVarchar2, object pDate, object pNumber)
        {
            if (pManagement == string.Empty)
            {
                //기본값 처리 위해 추가//
            }
            else
            {
                string vLookup_Type = Get_Lookup_Type(pManagement);
                if (vLookup_Type != pLookup_Type)
                {
                    return;
                }
            }

            object vManagement_Value = "";
            object vManagement_Desc = "";

            try
            {
                //관련 관리항목 기본값 설정//
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.SetCommandParamValue("W_LOOKUP_TYPE", pLookup_Type);
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.SetCommandParamValue("W_REF_LOOKUP_TYPE", pRef_Lookup_Type);
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.SetCommandParamValue("W_MANAGEMENT_VALUE", pManagement_Value);
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.SetCommandParamValue("W_VARCHAR2", pVarchar2);
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.SetCommandParamValue("W_DATE", iDate.ISGetDate(pDate));
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.SetCommandParamValue("W_NUMBER", iString.ISDecimaltoZero(pNumber));
                IDC_GET_CONTROL_ITEM_MANAGEMENT_P.ExecuteNonQuery();
                vManagement_Value = IDC_GET_CONTROL_ITEM_MANAGEMENT_P.GetCommandParamValue("O_MANAGEMENT_CODE");
                vManagement_Desc = IDC_GET_CONTROL_ITEM_MANAGEMENT_P.GetCommandParamValue("O_MANAGEMENT_DESC");
            }
            catch
            {
                vManagement_Value = "";
                vManagement_Desc = "";
            }
            Set_Management_Value(pRef_Lookup_Type, vManagement_Value, vManagement_Desc);
        }

        private void Set_Ref_Management_Value(string pManagement, string pLookup_Type, string pRef_Lookup_Type, object pManagement_Value, object pManagement_Desc)
        {
            if (pManagement == "MANAGEMENT1" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "MANAGEMENT2" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER1" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER2" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER3" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER4" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER5" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER6" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER7" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER8" &&
                iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
        }

        #endregion;

        #region ----- Initialize Event -----

        private Boolean Check_SlipHeader_Added()
        {
            Boolean Row_Added_Status = false;
            //헤더 체크 
            for (int r = 0; r < IDA_BATCH_HEADER.SelectRows.Count; r++)
            {
                if (IDA_BATCH_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_BATCH_HEADER.SelectRows[r].RowState == DataRowState.Modified)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10261"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            //헤더 변경없으면 라인 체크 
            if (Row_Added_Status == false)
            {
                for (int r = 0; r < IDA_BATCH_LINE.SelectRows.Count; r++)
                {
                    if (IDA_BATCH_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_BATCH_LINE.SelectRows[r].RowState == DataRowState.Modified)
                    {
                        Row_Added_Status = true;
                    }
                }
                if (Row_Added_Status == true)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10261"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            return (Row_Added_Status);
        }

        private void InsertSlipHeader()
        {
            TB_BATCH.SelectedIndex = 1;
            TB_BATCH.SelectedTab.Focus();
            
            BATCH_DATE.EditValue = DateTime.Today;
            GL_DATE.EditValue = BATCH_DATE.EditValue;

            idcDV_SLIP_TYPE.ExecuteNonQuery();
            SLIP_TYPE.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_SLIP_TYPE");
            SLIP_TYPE_NAME.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_SLIP_TYPE_NAME");
            SLIP_TYPE_CLASS.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_SLIP_TYPE_CLASS");
            DOCUMENT_TYPE.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_DOCUMENT_TYPE");

            idcUSER_INFO.ExecuteNonQuery();
            DEPT_NAME.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_NAME");
            DEPT_CODE.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_CODE");
            DEPT_ID.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_ID");
            PERSON_NAME.EditValue = idcUSER_INFO.GetCommandParamValue("O_PERSON_NAME");
            PERSON_ID.EditValue = isAppInterfaceAdv1.PERSON_ID;

            BATCH_DATE.Focus();
        }
        
        private void Set_Slip_Line_Seq()
        {
            //LINE SEQ 채번//
            decimal mSLIP_LINE_SEQ = 0;
            decimal vPre_Line_Seq = 0;
            decimal vNext_Line_Seq = 0;

            int mPreviousRowPosition = 0;
            try
            {
                mPreviousRowPosition = IDA_BATCH_LINE.CurrentRowPosition() - 1;
            }
            catch
            {
                mPreviousRowPosition = 0;
            }

            //현재 이전 line seq 
            if (mPreviousRowPosition > -1)
            {
                vPre_Line_Seq = iString.ISDecimaltoZero(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["BATCH_LINE_SEQ"], 1);
            }
            else
            {
                vPre_Line_Seq = 0;
            }

            try
            {
                //현재 다음 line seq
                int mNextRowPosition = 0;
                try
                {
                    mNextRowPosition = IDA_BATCH_LINE.CurrentRowPosition() + 1;
                }
                catch
                {
                    mNextRowPosition = 0;
                }

                if (mNextRowPosition == IDA_BATCH_LINE.CurrentRows.Count)
                {
                    vNext_Line_Seq = 0;
                }
                else
                {
                    vNext_Line_Seq = iString.ISDecimaltoZero(IDA_BATCH_LINE.CurrentRows[mNextRowPosition]["BATCH_LINE_SEQ"], 1);
                }

                //실재 Slip Line Seq 채번//
                if (vNext_Line_Seq == 0)
                {
                    mSLIP_LINE_SEQ = Math.Truncate(vPre_Line_Seq) + 10;
                }
                else
                {
                    decimal vAvg = Math.Round(((vNext_Line_Seq - vPre_Line_Seq) / 2), 10);
                    mSLIP_LINE_SEQ = vPre_Line_Seq + vAvg;
                }
            }
            catch
            {
                mSLIP_LINE_SEQ = Math.Truncate(vPre_Line_Seq) + 10;
            }
            igrSLIP_LINE.SetCellValue("BATCH_LINE_SEQ", mSLIP_LINE_SEQ); 
        }

        private void InsertSlipLine()
        {
            //LINE SEQ 채번//
            Set_Slip_Line_Seq();    //LINE SEQ 채번//
            CURRENCY_CODE.EditValue = mCurrency_Code;
            CURRENCY_DESC.EditValue = mCurrency_Code;
            Init_Currency_Amount();
            Init_Budget_Dept();
            GL_AMOUNT.EditValue = 0;
            GL_CURR_AMOUNT.EditValue = 0;

            BUDGET_DEPT_NAME_L.Focus();
        }
        
        private void Set_Delete_Batch_Line()
        {
            if (igrSLIP_LINE.RowCount < 1)
            {
                return;
            }

            IDA_BATCH_LINE.Cancel();
            IDA_BATCH_LINE.MoveFirst(igrSLIP_LINE.Name);
            for (int c = 0; c < igrSLIP_LINE.RowCount; c++)
            {
                IDA_BATCH_LINE.Delete();
                IDA_BATCH_LINE.MoveNext(igrSLIP_LINE.Name);
            }
            IDA_BATCH_LINE.OraSelectData.AcceptChanges();
            IDA_BATCH_LINE.Refillable = true;
        }

        private void Set_Insert_Batch_Line(object pPAYMENT_DATE)
        {
            IDA_BATCH_PAYMENT_SLIP.SetSelectParamValue("P_PAYMENT_DATE", pPAYMENT_DATE);
            IDA_BATCH_PAYMENT_SLIP.Fill();
            if (IDA_BATCH_PAYMENT_SLIP.SelectRows.Count < 0)
            {
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Not found data, Check data");
                return;
            }
            try
            {
                int Row_Count = igrSLIP_LINE.RowCount;

                string vSTATUS = string.Empty;

                igrSLIP_LINE.BeginUpdate();
                for (int i = 0; i < IDA_BATCH_PAYMENT_SLIP.SelectRows.Count; i++)
                {
                    IDA_BATCH_LINE.AddUnder();
                    for (int c = 0; c < igrSLIP_LINE.GridAdvExColElement.Count; c++)
                    {
                        if (igrSLIP_LINE.GridAdvExColElement[c].DataColumn.ToString() != "BATCH_HEADER_ID")
                        {
                            igrSLIP_LINE.SetCellValue(i + Row_Count, c, IDA_BATCH_PAYMENT_SLIP.OraDataSet().Rows[i][c]);
                        }
                    }
                    Set_Slip_Line_Seq();    //LINE SEQ 채번//
                }
                igrSLIP_LINE.EndUpdate();

                //상태변경.
                IDC_BATCH_PAYMENT_STATUS.SetCommandParamValue("P_PAYMENT_DATE", BATCH_DATE.EditValue);
                IDC_BATCH_PAYMENT_STATUS.ExecuteNonQuery();
                vSTATUS = IDC_BATCH_PAYMENT_STATUS.GetCommandParamValue("O_STATUS").ToString();

                IDA_BATCH_LINE.MoveFirst(igbSLIP_LINE.Name);
            }
            catch (Exception ex)
            {
                igrSLIP_LINE.EndUpdate();
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private string GET_BATCH_CLOSED_YN(object pBATCH_HEADER_ID)
        {
            string vBATCH_CLOSED_YN = "N";
            IDC_BATCH_CLOSED_YN.SetCommandParamValue("W_BATCH_HEADER_ID", pBATCH_HEADER_ID);
            IDC_BATCH_CLOSED_YN.ExecuteNonQuery();
            vBATCH_CLOSED_YN = iString.ISNull(IDC_BATCH_CLOSED_YN.GetCommandParamValue("P_CLOSED_YN"));
            if (IDC_BATCH_CLOSED_YN.ExcuteError)
            {
                vBATCH_CLOSED_YN = "F";
            }
            return vBATCH_CLOSED_YN;
        }

        private void Init_GL_Amount()
        {
            if (iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue) == 0)
            {
                return;
            }
            else if (iString.ISDecimaltoZero(GL_CURR_AMOUNT.EditValue) == 0)
            {
                return;
            }
            if (iString.ISNull(REF_SLIP_FLAG.EditValue) == "R" || iString.ISNull(REF_SLIP_FLAG.EditValue) == "S")
            {
                return;
            }
            
            decimal mGL_AMOUNT = iString.ISDecimaltoZero(GL_CURR_AMOUNT.EditValue) * iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue);
            try
            {
                idcCONVERSION_BASE_AMOUNT.SetCommandParamValue("W_BASE_CURRENCY_CODE", mCurrency_Code);
                idcCONVERSION_BASE_AMOUNT.SetCommandParamValue("W_CONVERSION_AMOUNT", mGL_AMOUNT);
                idcCONVERSION_BASE_AMOUNT.ExecuteNonQuery();
                GL_AMOUNT.EditValue = Convert.ToDecimal(idcCONVERSION_BASE_AMOUNT.GetCommandParamValue("O_BASE_AMOUNT"));
            }
            catch
            {
                GL_AMOUNT.EditValue = Convert.ToDecimal(Math.Round(mGL_AMOUNT, 0));
            }
            Init_DR_CR_Amount();
        }

        private bool Init_Exchange_Profit_Loss(decimal pNew_Exchange_Rate, decimal pOld_Exchange_Rate, decimal pCurrency_Amount, decimal pGL_Amount, int pCurrent_Row_Index)
        {
            bool mExchange_Profit_Loss = false;
            object vAccount_DR_CR;
            object vAccount_DR_CR_Name;
            object vAccount_ID;
            object vAccount_Code;
            object vAccount_Desc;
            decimal vExchange_Profit_Loss_Amount = Convert.ToDecimal(0);
            decimal vNew_GL_Amount = iString.ISDecimaltoZero(pNew_Exchange_Rate) * iString.ISDecimaltoZero(pCurrency_Amount) ;

            vExchange_Profit_Loss_Amount = vNew_GL_Amount - pGL_Amount;
            if (pCurrency_Amount != Convert.ToDecimal(0) && vExchange_Profit_Loss_Amount != Convert.ToDecimal(0))
            {                
                idcEXCHANGE_PROFIT_LOSS.SetCommandParamValue("W_CONVERSION_AMOUNT", vExchange_Profit_Loss_Amount);
                idcEXCHANGE_PROFIT_LOSS.ExecuteNonQuery();
                vAccount_ID = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_ID");
                vAccount_Code = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_CODE");
                vAccount_Desc = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_DESC");
                vAccount_DR_CR = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_DR_CR");
                vAccount_DR_CR_Name = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_DR_CR_NAME");

                // LINE 추가.
                IDA_BATCH_LINE.AddUnder();
                for (int c = 0; c < igrSLIP_LINE.ColCount; c++)
                {
                    igrSLIP_LINE.SetCellValue(igrSLIP_LINE.RowIndex, c, igrSLIP_LINE.GetCellValue(pCurrent_Row_Index, c));
                }
                // 반제 SLIP_LINE_ID.
                //igrSLIP_LINE.SetCellValue(igrSLIP_LINE.RowIndex, igrSLIP_LINE.GetColumnToIndex("UNLIQUIDATE_SLIP_LINE_ID"), );
                ACCOUNT_DR_CR.EditValue = vAccount_DR_CR;
                ACCOUNT_DR_CR_NAME.EditValue = vAccount_DR_CR_Name;
                ACCOUNT_CONTROL_ID.EditValue = vAccount_ID;
                ACCOUNT_CODE.EditValue = vAccount_Code;
                ACCOUNT_DESC.EditValue = vAccount_Desc;
                EXCHANGE_RATE.EditValue = iString.ISDecimaltoZero(pOld_Exchange_Rate);
                GL_CURR_AMOUNT.EditValue = iString.ISDecimaltoZero(pCurrency_Amount);
                GL_AMOUNT.EditValue = Math.Abs(iString.ISDecimaltoZero(vExchange_Profit_Loss_Amount));

                //참고항목 동기화.                
                Set_Control_Item_Prompt();
                Init_Set_Item_Prompt(IDA_BATCH_LINE.CurrentRow);

                Init_DR_CR_Amount();
                Init_Total_GL_Amount();
                mExchange_Profit_Loss = true;
            }
            return mExchange_Profit_Loss;
        }
         
        private bool Init_Offset_Account(int pCurrent_Row_Index)
        {
            int vIDX_ACCOUNT_DR_CR = igrSLIP_LINE.GetColumnToIndex("ACCOUNT_DR_CR");
            int vIDX_GL_CURRENCY_AMOUNT = igrSLIP_LINE.GetColumnToIndex("GL_CURR_AMOUNT");
            int vIDX_GL_AMOUNT = igrSLIP_LINE.GetColumnToIndex("GL_AMOUNT");
            int vIDX_CURRENCY_CODE = igrSLIP_LINE.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_EXCHANGE_RATE = igrSLIP_LINE.GetColumnToIndex("EXCHANGE_RATE");

            decimal vDR_Curr_GL_Amount = 0;
            decimal vCR_Curr_GL_Amount = 0;
            decimal vDR_GL_Amount = 0;
            decimal vCR_GL_Amount = 0;

            decimal vOffset_Curr_Amount = 0;
            decimal vOffset_Amount = 0;
            decimal vExchange_Rate = 0;

            string vCURRENCY_CODE = iString.ISNull(mCurrency_Code);
            object vOffset_Account_DR_CR = mOffset_Account_DR_CR;
            object vOffset_Account_DR_CR_NAME = mOffset_Account_DR_CR_Name;


            //외화금액이 있고 차액이 있을 경우만 처리//
            for (int r = 0; r < igrSLIP_LINE.RowCount; r++)
            {
                if (iString.ISNull(igrSLIP_LINE.GetCellValue(r, vIDX_ACCOUNT_DR_CR)) == "1")
                {
                    vDR_Curr_GL_Amount = vDR_Curr_GL_Amount + iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, vIDX_GL_CURRENCY_AMOUNT));
                    vDR_GL_Amount = vDR_GL_Amount + iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, vIDX_GL_AMOUNT));
                }
                else
                {
                    vCR_Curr_GL_Amount = vCR_Curr_GL_Amount + iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, vIDX_GL_CURRENCY_AMOUNT));
                    vCR_GL_Amount = vCR_GL_Amount + iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, vIDX_GL_AMOUNT));
                }
                //통화//
                if (iString.ISNull(igrSLIP_LINE.GetCellValue(r, vIDX_CURRENCY_CODE)) != mCurrency_Code)
                {
                    vCURRENCY_CODE = iString.ISNull(igrSLIP_LINE.GetCellValue(r, vIDX_CURRENCY_CODE));
                }
            }
            //차액 생성//
            if (mOffset_Account_DR_CR == "1")
            {
                vOffset_Curr_Amount = vCR_Curr_GL_Amount - vDR_Curr_GL_Amount;
                vOffset_Amount = vCR_GL_Amount - vDR_GL_Amount;
            }
            else
            {
                vOffset_Curr_Amount = vDR_Curr_GL_Amount - vCR_Curr_GL_Amount;
                vOffset_Amount = vDR_GL_Amount - vCR_GL_Amount;
            }

            if (vOffset_Amount == 0 && vOffset_Curr_Amount == 0)
            {
                return true;
            }

            if (vOffset_Amount < 0)
            {
                if (mOffset_Account_DR_CR == "1")
                {
                    vOffset_Account_DR_CR = "2";
                }
                else
                {
                    vOffset_Account_DR_CR = "1";
                }
                IDC_GET_ACCOUNT_DR_CR.SetCommandParamValue("W_GROUP_CODE", "ACCOUNT_DR_CR");
                IDC_GET_ACCOUNT_DR_CR.SetCommandParamValue("W_CODE", vOffset_Account_DR_CR);
                IDC_GET_ACCOUNT_DR_CR.ExecuteNonQuery();
                vOffset_Account_DR_CR_NAME = IDC_GET_ACCOUNT_DR_CR.GetCommandParamValue("O_RETURN_VALUE");
                vOffset_Amount = Math.Abs(vOffset_Amount);
                vOffset_Curr_Amount = Math.Abs(vOffset_Curr_Amount);
            }

            if (vOffset_Curr_Amount != 0)
            {
                //환율//
                vExchange_Rate = Math.Round(vOffset_Amount / vOffset_Curr_Amount, 4);
            }
            
            // LINE 추가.
            IDA_BATCH_LINE.AddUnder();
            InsertSlipLine();

            //Set_Slip_Line_Seq();    //LINE SEQ 채번//       
            ACCOUNT_CONTROL_ID.EditValue = mOffset_Account_Control_ID;
            ACCOUNT_CODE.EditValue = mOffset_Account_Code;
            ACCOUNT_DESC.EditValue = mOffset_Account_Desc;
            ACCOUNT_DR_CR.EditValue = vOffset_Account_DR_CR;
            ACCOUNT_DR_CR_NAME.EditValue = vOffset_Account_DR_CR_NAME;
            CURRENCY_CODE.EditValue = vCURRENCY_CODE;
            CURRENCY_DESC.EditValue = vCURRENCY_CODE;
            EXCHANGE_RATE.EditValue = vExchange_Rate;
            GL_CURR_AMOUNT.EditValue = vOffset_Curr_Amount;
            GL_AMOUNT.EditValue = vOffset_Amount;
            
            //참고항목 동기화.                
            Set_Control_Item_Prompt();
            Init_Set_Item_Prompt(IDA_BATCH_LINE.CurrentRow);

            Init_DR_CR_Amount();    // 차대금액 생성 //
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //

            return true;
        }

        private void Init_DR_CR_Amount()
        {
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            isAppInterfaceAdv1.OnAppMessage(string.Empty);

            if (igrSLIP_LINE.RowCount < 1)
            {
                return;
            }
            try
            {
                int vIDX_ROW_CURR = igrSLIP_LINE.RowIndex;
                if (IDA_BATCH_LINE.CurrentRowPosition() != vIDX_ROW_CURR)
                {
                    return;
                }

                int vIDX_COL_DRCR = igrSLIP_LINE.GetColumnToIndex("ACCOUNT_DR_CR");
                int vIDX_COL_GL_AMOUNT = igrSLIP_LINE.GetColumnToIndex("GL_AMOUNT");
                int vIDX_COL_DR = igrSLIP_LINE.GetColumnToIndex("DR_AMOUNT");
                int vIDX_COL_CR = igrSLIP_LINE.GetColumnToIndex("CR_AMOUNT");

                if (iString.ISNull(igrSLIP_LINE.GetCellValue(vIDX_ROW_CURR, vIDX_COL_DRCR), "1") == "1".ToString())
                {
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, igrSLIP_LINE.GetCellValue(vIDX_ROW_CURR, vIDX_COL_GL_AMOUNT));
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, 0);
                }
                else if (iString.ISNull(igrSLIP_LINE.GetCellValue(vIDX_ROW_CURR, vIDX_COL_DRCR), "1") == "2".ToString())
                {
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, 0);
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, igrSLIP_LINE.GetCellValue(vIDX_ROW_CURR, vIDX_COL_GL_AMOUNT));
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private void Init_Total_GL_Amount()
        {
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";

            decimal vDR_Amount = Convert.ToDecimal(0);
            decimal vCR_Amount = Convert.ToDecimal(0);
            decimal vCurrency_DR_Amount = Convert.ToInt32(0);

            if (IDA_BATCH_LINE.CurrentRows.Count < 1)
            {
                return;
            }

            foreach (DataRow vRow in IDA_BATCH_LINE.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Deleted)
                {
                    if (iString.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "1".ToString())
                    {
                        vDR_Amount = vDR_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                        vCurrency_DR_Amount = vCurrency_DR_Amount + iString.ISDecimaltoZero(vRow["GL_CURR_AMOUNT"]);
                    }
                    else if (iString.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "2".ToString())
                    {
                        vCR_Amount = vCR_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]); ;
                    }
                }
            }
            TOTAL_DR_AMOUNT.EditValue = iString.ISDecimaltoZero(vDR_Amount);
            TOTAL_CR_AMOUNT.EditValue = iString.ISDecimaltoZero(vCR_Amount);
            MARGIN_AMOUNT.EditValue = -(System.Math.Abs(iString.ISDecimaltoZero(vDR_Amount) - iString.ISDecimaltoZero(vCR_Amount)));
        }

        private void Init_Control_Management_Value()
        {
            igrSLIP_LINE.SetCellValue("MANAGEMENT1", null);
            igrSLIP_LINE.SetCellValue("MANAGEMENT1_DESC", null);
            igrSLIP_LINE.SetCellValue("MANAGEMENT2", null);
            igrSLIP_LINE.SetCellValue("MANAGEMENT2_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER1", null);
            igrSLIP_LINE.SetCellValue("REFER1_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER2", null);
            igrSLIP_LINE.SetCellValue("REFER2_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER3", null);
            igrSLIP_LINE.SetCellValue("REFER3_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER4", null);
            igrSLIP_LINE.SetCellValue("REFER4_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER5", null);
            igrSLIP_LINE.SetCellValue("REFER5_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER6", null);
            igrSLIP_LINE.SetCellValue("REFER6_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER7", null);
            igrSLIP_LINE.SetCellValue("REFER7_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER8", null);
            igrSLIP_LINE.SetCellValue("REFER8_DESC", null);
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

            //참조된 전표 계정과목, 차대구분, 통화, 환율 제어//
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) == "R" || iString.ISNull(pDataRow["REF_SLIP_FLAG"]) == "S")
            {
                ACCOUNT_CODE.ReadOnly = true;
                ACCOUNT_CODE.Insertable = false;
                ACCOUNT_CODE.Updatable = false;
                ACCOUNT_CODE.TabStop = false;
                ACCOUNT_CODE.Refresh();

                ACCOUNT_DR_CR_NAME.ReadOnly = true;
                ACCOUNT_DR_CR_NAME.Insertable = false;
                ACCOUNT_DR_CR_NAME.Updatable = false;
                ACCOUNT_DR_CR_NAME.TabStop = false;
                ACCOUNT_DR_CR_NAME.Refresh();

                CURRENCY_DESC.ReadOnly = true;
                CURRENCY_DESC.Insertable = false;
                CURRENCY_DESC.Updatable = false;
                CURRENCY_DESC.TabStop = false;
                CURRENCY_DESC.Refresh();

                ////원전표인 경우 금액수정 불가//
                if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) == "S")
                {
                    GL_AMOUNT.ReadOnly = true;
                    GL_AMOUNT.Insertable = false;
                    GL_AMOUNT.Updatable = false;
                    GL_AMOUNT.TabStop = false;
                    GL_AMOUNT.Refresh();
                }

                //외화//
                Init_Currency_Amount();
            }
            else
            {
                ACCOUNT_CODE.ReadOnly = false;
                ACCOUNT_CODE.Insertable = true;
                ACCOUNT_CODE.Updatable = true;
                ACCOUNT_CODE.TabStop = true;
                ACCOUNT_CODE.Refresh();

                ACCOUNT_DR_CR_NAME.ReadOnly = false;
                ACCOUNT_DR_CR_NAME.Insertable = true;
                ACCOUNT_DR_CR_NAME.Updatable = true;
                ACCOUNT_DR_CR_NAME.TabStop = true;
                ACCOUNT_DR_CR_NAME.Refresh();

                CURRENCY_DESC.ReadOnly = false;
                CURRENCY_DESC.Insertable = true;
                CURRENCY_DESC.Updatable = true;
                CURRENCY_DESC.TabStop = true;
                CURRENCY_DESC.Refresh();

                GL_AMOUNT.ReadOnly = false;
                GL_AMOUNT.Insertable = true;
                GL_AMOUNT.Updatable = true;
                GL_AMOUNT.TabStop = true;
                GL_AMOUNT.Refresh();

                //외화//
                Init_Currency_Amount();
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            CURRENCY_DESC.Nullable = true;
            if (iString.ISNull(pDataRow["CURRENCY_ENABLED_FLAG"], "N") == "Y".ToString())
            {
                CURRENCY_DESC.Nullable = false;
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            string mDATA_TYPE = "VARCHAR2";
            object mValue;
            mDATA_TYPE = iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
            MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT1.Nullable = true;
            MANAGEMENT1.ReadOnly = true;
            MANAGEMENT1.Insertable = false;
            MANAGEMENT1.Updatable = false;
            MANAGEMENT1.TabStop = false;

            if (iString.ISNull(pDataRow["MANAGEMENT1_YN"], "F") != "F".ToString())
            {
                MANAGEMENT1.ReadOnly = false;
                MANAGEMENT1.Insertable = true;
                MANAGEMENT1.Updatable = true;
                MANAGEMENT1.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("MANAGEMENT1");
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT1.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("MANAGEMENT1", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("MANAGEMENT1");
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT1.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("MANAGEMENT1", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("MANAGEMENT1");
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    MANAGEMENT1.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("MANAGEMENT1", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                MANAGEMENT1.ReadOnly = true;
                MANAGEMENT1.Insertable = false;
                MANAGEMENT1.Updatable = false;
                MANAGEMENT1.TabStop = false;
            }
            MANAGEMENT1.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
            MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT2.Nullable = true;
            MANAGEMENT2.ReadOnly = true;
            MANAGEMENT2.Insertable = false;
            MANAGEMENT2.Updatable = false;
            MANAGEMENT2.TabStop = false;
            if (iString.ISNull(pDataRow["MANAGEMENT2_YN"], "F") != "F".ToString())
            {
                MANAGEMENT2.ReadOnly = false;
                MANAGEMENT2.Insertable = true;
                MANAGEMENT2.Updatable = true;
                MANAGEMENT2.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("MANAGEMENT2");
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT2.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("MANAGEMENT2", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("MANAGEMENT2");
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT2.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("MANAGEMENT2", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("MANAGEMENT2");
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    MANAGEMENT2.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("MANAGEMENT2", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                MANAGEMENT2.ReadOnly = true;
                MANAGEMENT2.Insertable = false;
                MANAGEMENT2.Updatable = false;
                MANAGEMENT2.TabStop = false;
            }
            MANAGEMENT2.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER1_DATA_TYPE"]);
            REFER1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER1.Nullable = true;
            REFER1.ReadOnly = true;
            REFER1.Insertable = false;
            REFER1.Updatable = false;
            REFER1.TabStop = false;
            if (iString.ISNull(pDataRow["REFER1_YN"], "F") != "F".ToString())
            {
                REFER1.ReadOnly = false;
                REFER1.Insertable = true;
                REFER1.Updatable = true;
                REFER1.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER1");
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER1.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER1", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER1");
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER1.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER1", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER1");
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER1.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER1", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER1.ReadOnly = true;
                REFER1.Insertable = false;
                REFER1.Updatable = false;
                REFER1.TabStop = false;
            }
            REFER1.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER2_DATA_TYPE"]);
            REFER2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER2.Nullable = true;
            REFER2.ReadOnly = true;
            REFER2.Insertable = false;
            REFER2.Updatable = false;
            REFER2.TabStop = false;
            if (iString.ISNull(pDataRow["REFER2_YN"], "F") != "F".ToString())
            {
                REFER2.ReadOnly = false;
                REFER2.Insertable = true;
                REFER2.Updatable = true;
                REFER2.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER2");
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER2.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER2", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER2");
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER2.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER2", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER2");
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER2.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER2", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER2.ReadOnly = true;
                REFER2.Insertable = false;
                REFER2.Updatable = false;
                REFER2.TabStop = false;
            }
            REFER2.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER3_DATA_TYPE"]);
            REFER3.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER3.Nullable = true;
            REFER3.ReadOnly = true;
            REFER3.Insertable = false;
            REFER3.Updatable = false;
            REFER3.TabStop = false;
            if (iString.ISNull(pDataRow["REFER3_YN"], "F") != "F".ToString())
            {
                REFER3.ReadOnly = false;
                REFER3.Insertable = true;
                REFER3.Updatable = true;
                REFER3.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER3");
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER3.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER3", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER3");
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER3.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER3", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER3");
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER3.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER3", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER3.ReadOnly = true;
                REFER3.Insertable = false;
                REFER3.Updatable = false;
                REFER3.TabStop = false;
            }
            REFER3.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER4_DATA_TYPE"]);
            REFER4.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER4.Nullable = true;
            REFER4.ReadOnly = true;
            REFER4.Insertable = false;
            REFER4.Updatable = false;
            REFER4.TabStop = false;
            if (iString.ISNull(pDataRow["REFER4_YN"], "F") != "F".ToString())
            {
                REFER4.ReadOnly = false;
                REFER4.Insertable = true;
                REFER4.Updatable = true;
                REFER4.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER4");
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER4.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER4", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER4");
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER4.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER4", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER4");
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER4.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER4", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER4.ReadOnly = true;
                REFER4.Insertable = false;
                REFER4.Updatable = false;
                REFER4.TabStop = false;
            }
            REFER4.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER5_DATA_TYPE"]);
            REFER5.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER5.Nullable = true;
            REFER5.ReadOnly = true;
            REFER5.Insertable = false;
            REFER5.Updatable = false;
            REFER5.TabStop = false;
            if (iString.ISNull(pDataRow["REFER5_YN"], "F") != "F".ToString())
            {
                REFER5.ReadOnly = false;
                REFER5.Insertable = true;
                REFER5.Updatable = true;
                REFER5.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER5");
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER5.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER5", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER5");
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER5.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER5", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER5");
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER5.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER5", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER5.ReadOnly = true;
                REFER5.Insertable = false;
                REFER5.Updatable = false;
                REFER5.TabStop = false;
            }
            REFER5.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER6_DATA_TYPE"]);
            REFER6.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER6.Nullable = true;
            REFER6.ReadOnly = true;
            REFER6.Insertable = false;
            REFER6.Updatable = false;
            REFER6.TabStop = false;
            if (iString.ISNull(pDataRow["REFER6_YN"], "F") != "F".ToString())
            {
                REFER6.ReadOnly = false;
                REFER6.Insertable = true;
                REFER6.Updatable = true;
                REFER6.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER6");
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER6.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER6", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER6");
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER6.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER6", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER6");
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER6.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER6", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER6.ReadOnly = true;
                REFER6.Insertable = false;
                REFER6.Updatable = false;
                REFER6.TabStop = false;
            }
            REFER6.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER7_DATA_TYPE"]);
            REFER7.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER7.Nullable = true;
            REFER7.ReadOnly = true;
            REFER7.Insertable = false;
            REFER7.Updatable = false;
            REFER7.TabStop = false;
            if (iString.ISNull(pDataRow["REFER7_YN"], "F") != "F".ToString())
            {
                REFER7.ReadOnly = false;
                REFER7.Insertable = true;
                REFER7.Updatable = true;
                REFER7.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER7");
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER7.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER7", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER7");
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER7.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER7", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER7");
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER7.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER7", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER7.ReadOnly = true;
                REFER7.Insertable = false;
                REFER7.Updatable = false;
                REFER7.TabStop = false;
            }
            REFER7.Refresh();

            mDATA_TYPE = iString.ISNull(pDataRow["REFER8_DATA_TYPE"]);
            REFER8.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER8.Nullable = true;
            REFER8.ReadOnly = true;
            REFER8.Insertable = false;
            REFER8.Updatable = false;
            REFER8.TabStop = false;
            if (iString.ISNull(pDataRow["REFER8_YN"], "F") != "F".ToString())
            {
                REFER8.ReadOnly = false;
                REFER8.Insertable = true;
                REFER8.Updatable = true;
                REFER8.TabStop = true;
                if (mDATA_TYPE == "NUMBER".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER8");
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER8.NumberDecimalDigits = 0;
                    igrSLIP_LINE.SetCellValue("REFER8", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER8");
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER8.NumberDecimalDigits = 4;
                    igrSLIP_LINE.SetCellValue("REFER8", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = igrSLIP_LINE.GetCellValue("REFER8");
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER8.DateFormat = "yyyy-MM-dd";
                    igrSLIP_LINE.SetCellValue("REFER8", mValue);
                }
            }
            if (iString.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER8.ReadOnly = true;
                REFER8.Insertable = false;
                REFER8.Updatable = false;
                REFER8.TabStop = false;
            }
            REFER8.Refresh();

            ///////////////////////////////////////////////////////////////////////////////////////////////////            
            if (iString.ISNull(pDataRow["MANAGEMENT1_LOOKUP_YN"], "N") == "Y".ToString())
            {
                MANAGEMENT1.LookupAdapter = ilaMANAGEMENT1;
            }
            else
            {
                MANAGEMENT1.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["MANAGEMENT2_LOOKUP_YN"], "N") == "Y".ToString())
            {
                MANAGEMENT2.LookupAdapter = ilaMANAGEMENT2;
            }
            else
            {
                MANAGEMENT2.LookupAdapter = null;
            }
            if (iString.ISNull(pDataRow["REFER1_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER1.LookupAdapter = ilaREFER1;
            }
            else
            {
                REFER1.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER2_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER2.LookupAdapter = ilaREFER2;
            }
            else
            {
                REFER2.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER3_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER3.LookupAdapter = ilaREFER3;
            }
            else
            {
                REFER3.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER4_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER4.LookupAdapter = ilaREFER4;
            }
            else
            {
                REFER4.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER5_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER5.LookupAdapter = ilaREFER5;
            }
            else
            {
                REFER5.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER6_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER6.LookupAdapter = ilaREFER6;
            }
            else
            {
                REFER6.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER7_LOOKUP_YN"], "N") == "Y".ToString())
            {
                REFER7.LookupAdapter = ilaREFER7;
            }
            else
            {
                REFER7.LookupAdapter = null;
            }

            if (iString.ISNull(pDataRow["REFER8_LOOKUP_YN"], "N") == "Y".ToString())
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
            if (MANAGEMENT1.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("MANAGEMENT1");
                MANAGEMENT1.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["MANAGEMENT1_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    MANAGEMENT1.ReadOnly = true;
                    MANAGEMENT1.Nullable = false;
                    MANAGEMENT1.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("MANAGEMENT1", mDATA_VALUE);
                MANAGEMENT1.Refresh();
            }

            //--2
            if (MANAGEMENT2.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("MANAGEMENT2");
                MANAGEMENT2.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["MANAGEMENT2_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    MANAGEMENT2.ReadOnly = true;
                    MANAGEMENT2.Nullable = false;
                    MANAGEMENT2.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("MANAGEMENT2", mDATA_VALUE);
                MANAGEMENT2.Refresh();
            }

            //--3
            if (REFER1.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER1");
                REFER1.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER1_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER1_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER1.ReadOnly = true;
                    REFER1.Nullable = false;
                    REFER1.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER1", mDATA_VALUE);
                REFER1.Refresh();
            }

            //--4
            if (REFER2.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER2");
                REFER2.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER2_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER2_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER2.ReadOnly = true;
                    REFER2.Nullable = false;
                    REFER2.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER2", mDATA_VALUE);
                REFER2.Refresh();
            }

            //--5
            if (REFER3.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER3");
                REFER3.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER3_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER3_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER3.ReadOnly = true;
                    REFER3.Nullable = false;
                    REFER3.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER3", mDATA_VALUE);
                REFER3.Refresh();
            }

            //--6
            if (REFER4.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER4");
                REFER4.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER4_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER4_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER4.ReadOnly = true;
                    REFER4.Nullable = false;
                    REFER4.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER4", mDATA_VALUE);
                REFER4.Refresh();
            }

            //--7
            if (REFER5.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER5");
                REFER5.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER5_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER5_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER5.ReadOnly = true;
                    REFER5.Nullable = false;
                    REFER5.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER5", mDATA_VALUE);
                REFER5.Refresh();
            }

            //--8
            if (REFER6.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER6");
                REFER6.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER6_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER6_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER6.ReadOnly = true;
                    REFER6.Nullable = false;
                    REFER6.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER6", mDATA_VALUE);
                REFER6.Refresh();
            }

            //--9
            if (REFER7.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER7");
                REFER7.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER7_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER7_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER7.ReadOnly = true;
                    REFER7.Nullable = false;
                    REFER7.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER7", mDATA_VALUE);
                REFER7.Refresh();
            }

            //--10
            if (REFER8.ReadOnly == false)
            {
                mDATA_VALUE = igrSLIP_LINE.GetCellValue("REFER8");
                REFER8.Nullable = true;
                mDATA_TYPE = iString.ISNull(pDataRow["REFER8_DATA_TYPE"]);
                mDR_CR_YN = iString.ISNull(pDataRow["REFER8_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER8.ReadOnly = true;
                    REFER8.Nullable = false;
                    REFER8.ReadOnly = false;
                }
                igrSLIP_LINE.SetCellValue("REFER8", mDATA_VALUE);
                REFER8.Refresh();
            }
        }

        private void Init_Default_Value()
        {
            int mPreviousRowPosition = IDA_BATCH_LINE.CurrentRowPosition() - 1;
            object mPrevious_Code;
            object mPrevious_Name;
            string mData_Type;
            string mLookup_Type;

            if (mPreviousRowPosition > -1
                && iString.ISNull(REMARK.EditValue) == string.Empty
                && iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REMARK"]) != string.Empty)
            {//REMARK.
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REMARK"];
                REMARK.EditValue = mPrevious_Name;
            }

            //1
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT1.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT1.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_LOOKUP_TYPE"]))
            {//MANAGEMENT1_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_DESC"];

                MANAGEMENT1.EditValue = mPrevious_Code;
                MANAGEMENT1_DESC.EditValue = mPrevious_Name;
            }
            //2
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT2.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT2.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_LOOKUP_TYPE"]))
            {//MANAGEMENT2_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_DESC"];

                MANAGEMENT2.EditValue = mPrevious_Code;
                MANAGEMENT2_DESC.EditValue = mPrevious_Name;
            }
            //3
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER1.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER1.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER1_LOOKUP_TYPE"]))
            {//REFER1_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER1"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER1_DESC"];

                REFER1.EditValue = mPrevious_Code;
                REFER1_DESC.EditValue = mPrevious_Name;
            }
            //4
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER2.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER2.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER2_LOOKUP_TYPE"]))
            {//REFER2_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER2"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER2_DESC"];

                REFER2.EditValue = mPrevious_Code;
                REFER2_DESC.EditValue = mPrevious_Name;
            }
            //5
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER3.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER3.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER3_LOOKUP_TYPE"]))
            {//REFER3_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER3"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER3_DESC"];

                REFER3.EditValue = mPrevious_Code;
                REFER3_DESC.EditValue = mPrevious_Name;
            }
            //6
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER4.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER4.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER4_LOOKUP_TYPE"]))
            {//REFER4_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER4"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER4_DESC"];

                REFER4.EditValue = mPrevious_Code;
                REFER4_DESC.EditValue = mPrevious_Name;
            }
            //7
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER5.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER5.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER5_LOOKUP_TYPE"]))
            {//REFER5_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER5"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER5_DESC"];

                REFER5.EditValue = mPrevious_Code;
                REFER5_DESC.EditValue = mPrevious_Name;
            }
            //8
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER6.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER6.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER6_LOOKUP_TYPE"]))
            {//REFER6_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER6"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER6_DESC"];

                REFER6.EditValue = mPrevious_Code;
                REFER6_DESC.EditValue = mPrevious_Name;
            }
            //9
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER7.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER7.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER7_LOOKUP_TYPE"]))
            {//REFER7_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER7"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER7_DESC"];

                REFER7.EditValue = mPrevious_Code;
                REFER7_DESC.EditValue = mPrevious_Name;
            }
            //10
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER8.EditValue) == string.Empty && iString.ISNull(BATCH_DATE.EditValue) != string.Empty)
                {
                    REFER8.EditValue = Convert.ToDateTime(BATCH_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER8_LOOKUP_TYPE"]))
            {//REFER8_LOOKUP_TYPE
                mPrevious_Code = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER8"];
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["REFER8_DESC"];

                REFER8.EditValue = mPrevious_Code;
                REFER8_DESC.EditValue = mPrevious_Name;
            }
        }

        private void Init_Currency_Code(string pInit_YN)
        {
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("CURRENCY_ENABLED_FLAG"), "N") == "Y")
            {
                CURRENCY_DESC.ReadOnly = false;
                CURRENCY_DESC.Insertable = true;
                CURRENCY_DESC.Updatable = true;
                CURRENCY_DESC.TabStop = true;
            }
            else
            {
                CURRENCY_DESC.ReadOnly = true;
                CURRENCY_DESC.Insertable = false;
                CURRENCY_DESC.Updatable = false;
                CURRENCY_DESC.TabStop = false;
                if (pInit_YN == "Y")
                {
                    CURRENCY_CODE.EditValue = mCurrency_Code;
                    CURRENCY_DESC.EditValue = mCurrency_Code;
                    Init_Currency_Amount();
                }
            }
            CURRENCY_CODE.Invalidate();
            CURRENCY_DESC.Invalidate();
        }

        private void Init_Currency_Amount()
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) == string.Empty || iString.ISNull(CURRENCY_CODE.EditValue) == mCurrency_Code)
            {
                if (iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue) != Convert.ToDecimal(0))
                {
                    EXCHANGE_RATE.EditValue = null;
                }
                if (iString.ISDecimaltoZero(GL_CURR_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    GL_CURR_AMOUNT.EditValue = null;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                GL_CURR_AMOUNT.ReadOnly = true;
                GL_CURR_AMOUNT.Insertable = false;
                GL_CURR_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                GL_CURR_AMOUNT.TabStop = false;
            }
            else
            {
                if (iString.ISNull(IDA_BATCH_LINE.CurrentRow["REF_SLIP_FLAG"]) != string.Empty)
                {
                    EXCHANGE_RATE.ReadOnly = true;
                    EXCHANGE_RATE.Insertable = false;
                    EXCHANGE_RATE.Updatable = false;
                    EXCHANGE_RATE.TabStop = false;

                    //원전표인 경우 금액수정 불가//
                    if (iString.ISNull(IDA_BATCH_LINE.CurrentRow["REF_SLIP_FLAG"]) == "S")
                    {
                        GL_CURR_AMOUNT.ReadOnly = true;
                        GL_CURR_AMOUNT.Insertable = false;
                        GL_CURR_AMOUNT.Updatable = false;
                        GL_CURR_AMOUNT.TabStop = false;
                    }
                }
                else
                {
                    EXCHANGE_RATE.ReadOnly = false;
                    EXCHANGE_RATE.Insertable = true;
                    EXCHANGE_RATE.Updatable = true;
                    EXCHANGE_RATE.TabStop = true;

                    GL_CURR_AMOUNT.ReadOnly = false;
                    GL_CURR_AMOUNT.Insertable = true;
                    GL_CURR_AMOUNT.Updatable = true;
                    GL_CURR_AMOUNT.TabStop = true;
                }
            }
            EXCHANGE_RATE.Refresh();
            GL_CURR_AMOUNT.Refresh();
        }

        // 부가세 관련 설정 제어 - 세액/공급가액(세액 * 10)
        private void Init_VAT_Amount()
        {
            object mVAT_ENABLED_FLAG = IDA_BATCH_LINE.CurrentRow["VAT_ENABLED_FLAG"];
            if (iString.ISNull(mVAT_ENABLED_FLAG, "N") != "Y")
            {
                return;
            }

            IDC_GET_ACCOUNT_DEFAULT_VALUE.SetCommandParamValue("W_ACCOUNT_TYPE", "DEFAULT_VAT_RATE");
            IDC_GET_ACCOUNT_DEFAULT_VALUE.ExecuteNonQuery();
            decimal vVAT_RATE = iString.ISDecimaltoZero(IDC_GET_ACCOUNT_DEFAULT_VALUE.GetCommandParamValue("O_VAT_RATE"));

            decimal mGL_AMOUNT = iString.ISDecimaltoZero(GL_AMOUNT.EditValue);
            decimal mSUPPLY_AMOUNT = mGL_AMOUNT * vVAT_RATE; //공급가액 설정.

            Set_Management_Value("SUPPLY_AMOUNT", mSUPPLY_AMOUNT, null);
            Set_Management_Value("VAT_AMOUNT", mGL_AMOUNT, null);
        }

        //예산부서 동기화
        private void Init_Budget_Dept()
        {
            int mPreviousRowPosition = IDA_BATCH_LINE.CurrentRowPosition() - 1;
            object mPrevious_ID; 
            object mPrevious_Name;

            if (mPreviousRowPosition > -1
                && iString.ISNull(BUDGET_DEPT_ID_L.EditValue) == string.Empty
                && iString.ISNull(IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_ID"]) != string.Empty)
            {//budget dept
                mPrevious_ID = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_ID"]; 
                mPrevious_Name = IDA_BATCH_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_NAME"];

                BUDGET_DEPT_NAME_L.EditValue = mPrevious_Name; 
                BUDGET_DEPT_ID_L.EditValue = mPrevious_ID;
            }
            else
            {
                BUDGET_DEPT_NAME_L.EditValue = DEPT_NAME.EditValue; 
                BUDGET_DEPT_ID_L.EditValue = DEPT_ID.EditValue;
            }
        }
        
        //관리항목 기본값//
        private void Init_Default_Management(string pLookup_Type)
        {
            if (iString.ISNull(Get_Management_Value(pLookup_Type)) != string.Empty)
            {
                return;
            }

            if (pLookup_Type == "DEPT")
            {
                //예산부서//
                Set_Management_Value("DEPT", BUDGET_DEPT_CODE_L.EditValue, BUDGET_DEPT_NAME_L.EditValue);
            }
            else if (pLookup_Type == "TAX_CODE")
            {
                //부가세 사업장코드//
                Set_Ref_Management_Value(string.Empty, "TAX_CODE", "TAX_CODE", null);
            }
        }

        //관리항목 LOOKUP 선택시 처리.
        private void Init_SELECT_LOOKUP(object pManagement_Type)
        {
            string mMANAGEMENT = iString.ISNull(pManagement_Type);
        }

        #endregion

        #region ----- XL Export Methods ----

        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.CurrentRows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xls";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(pAdapter.OraSelectData, vsSaveExcelFileName, vsSheetName);
                if (vXLSaveOK == true)
                {
                    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                else
                {
                    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                xlExport.XLClose();
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

        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID, object pSlip_Header_ID, object pGL_Date, object pGL_Num)
        {
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vCurrAssemblyFileVersion = string.Empty;

            //[EAPP_ASSEMBLY_INFO_G.MENU_ENTRY_PROCESS_START]
            IDC_MENU_ENTRY_MANUAL_START.SetCommandParamValue("W_ASSEMBLY_ID", pAssembly_ID);
            IDC_MENU_ENTRY_MANUAL_START.ExecuteNonQuery();

            string vREAD_FLAG = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_READ_FLAG"));
            string vUSER_TYPE = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_USER_TYPE"));
            string vPRINT_FLAG = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_PRINT_FLAG"));

            decimal vASSEMBLY_INFO_ID = iString.ISDecimaltoZero(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_INFO_ID"));
            string vASSEMBLY_ID = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_ID"));
            string vASSEMBLY_NAME = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_NAME"));
            string vASSEMBLY_FILE_NAME = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_FILE_NAME"));

            string vASSEMBLY_VERSION = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_VERSION"));
            string vDIR_FULL_PATH = iString.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_DIR_FULL_PATH"));

            System.IO.FileInfo vFile = new System.IO.FileInfo(vASSEMBLY_FILE_NAME);
            if (vFile.Exists)
            {
                vCurrAssemblyFileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(vASSEMBLY_FILE_NAME).FileVersion;
            }

            vREAD_FLAG = "Y";  //무조건 인쇄

            //1. Assembly file Name(.dll) 있을겨우만 실행//
            if (vASSEMBLY_FILE_NAME != string.Empty)
            {
                //2. 읽기 권한 있을 경우만 실행 //
                if (vREAD_FLAG == "Y")
                {
                    if (vCurrAssemblyFileVersion != vASSEMBLY_VERSION)
                    {
                        ISFileTransferAdv vFileTransferAdv = new ISFileTransferAdv();

                        vFileTransferAdv.Host = isAppInterfaceAdv1.AppInterface.AppHostInfo.Host;
                        vFileTransferAdv.Port = isAppInterfaceAdv1.AppInterface.AppHostInfo.Port;
                        vFileTransferAdv.UserId = isAppInterfaceAdv1.AppInterface.AppHostInfo.UserId;
                        vFileTransferAdv.Password = isAppInterfaceAdv1.AppInterface.AppHostInfo.Password;
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive == "Y")
                        {
                            vFileTransferAdv.UsePassive = true;
                        }
                        else
                        {
                            vFileTransferAdv.UsePassive = false;
                        }

                        vFileTransferAdv.SourceDirectory = vDIR_FULL_PATH;
                        vFileTransferAdv.SourceFileName = vASSEMBLY_FILE_NAME;
                        vFileTransferAdv.TargetDirectory = Application.StartupPath;
                        vFileTransferAdv.TargetFileName = "_" + vASSEMBLY_FILE_NAME;

                        if (vFileTransferAdv.Download() == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vASSEMBLY_FILE_NAME);
                                System.IO.File.Move("_" + vASSEMBLY_FILE_NAME, vASSEMBLY_FILE_NAME);
                            }
                            catch
                            {
                                try
                                {
                                    System.IO.FileInfo vFileInfo = new System.IO.FileInfo("_" + vASSEMBLY_FILE_NAME);
                                    if (vFileInfo.Exists == true)
                                    {
                                        try
                                        {
                                            System.IO.File.Delete("_" + vASSEMBLY_FILE_NAME);
                                        }
                                        catch
                                        {
                                            // ignore
                                        }
                                    }
                                }
                                catch
                                {
                                    //ignore
                                }
                            }
                        }

                        //report update//
                        ReportUpdate(vASSEMBLY_INFO_ID);
                    }

                    try
                    {
                        System.Reflection.Assembly vAssembly = System.Reflection.Assembly.LoadFrom(vASSEMBLY_FILE_NAME);
                        Type vType = vAssembly.GetType(vAssembly.GetName().Name + "." + vAssembly.GetName().Name);

                        if (vType != null)
                        {
                            if (vFile.Exists)
                            {
                                vCurrAssemblyFileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(vASSEMBLY_FILE_NAME).FileVersion;
                            }

                            object[] vParam = new object[6];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = pSlip_Header_ID;     //전표 헤더 id
                            vParam[3] = pGL_Date;     //전표일자
                            vParam[4] = pGL_Num;                   //전표번호
                            vParam[5] = "Y";      //프린트 옵션 표시 여부

                            object vCreateInstance = Activator.CreateInstance(vType, vParam);
                            Office2007Form vForm = vCreateInstance as Office2007Form;
                            Point vPoint = new Point(30, 30);
                            vForm.Location = vPoint;
                            vForm.StartPosition = FormStartPosition.Manual;
                            vForm.Text = string.Format("{0}[{1}] - {2}", vASSEMBLY_NAME, vASSEMBLY_ID, vCurrAssemblyFileVersion);

                            vForm.Show();
                        }
                        else
                        {
                            MessageBoxAdv.Show("Form Namespace Error");
                        }
                    }
                    catch
                    {
                        //
                    }
                }
            }

            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        //report download//
        private void ReportUpdate(object pAssemblyInfoID)
        {
            string vPathReportFTP = string.Empty;
            string vReportFileName = string.Empty;
            string vReportFileNameTarget = string.Empty;

            try
            {
                IDA_REPORT_INFO_DOWNLOAD.SetSelectParamValue("W_ASSEMBLY_INFO_ID", pAssemblyInfoID);
                IDA_REPORT_INFO_DOWNLOAD.Fill();
                if (IDA_REPORT_INFO_DOWNLOAD.OraSelectData.Rows.Count > 0)
                {
                    ISFileTransferAdv vFileTransferAdv = new ISFileTransferAdv();

                    vFileTransferAdv.Host = isAppInterfaceAdv1.AppInterface.AppHostInfo.Host;
                    vFileTransferAdv.Port = isAppInterfaceAdv1.AppInterface.AppHostInfo.Port;
                    if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive == "Y")
                    {
                        vFileTransferAdv.UsePassive = true;
                    }
                    else
                    {
                        vFileTransferAdv.UsePassive = false;
                    }
                    vFileTransferAdv.UserId = isAppInterfaceAdv1.AppInterface.AppHostInfo.UserId;
                    vFileTransferAdv.Password = isAppInterfaceAdv1.AppInterface.AppHostInfo.Password;

                    foreach (System.Data.DataRow vRow in IDA_REPORT_INFO_DOWNLOAD.OraSelectData.Rows)
                    {
                        if (iString.ISNull(vRow["REPORT_FILE_NAME"]) != string.Empty)
                        {
                            vReportFileName = iString.ISNull(vRow["REPORT_FILE_NAME"]);
                            vReportFileNameTarget = string.Format("_{0}", vReportFileName);
                        }
                        if (iString.ISNull(vRow["REPORT_PATH_FTP"]) != string.Empty)
                        {
                            vPathReportFTP = iString.ISNull(vRow["REPORT_PATH_FTP"]);
                        }

                        if (vReportFileName != string.Empty && vPathReportFTP != string.Empty)
                        {
                            string vPathReportClient = string.Format("{0}\\{1}", System.Windows.Forms.Application.StartupPath, "Report");
                            System.IO.DirectoryInfo vReport = new System.IO.DirectoryInfo(vPathReportClient);
                            if (vReport.Exists == false) //있으면 True, 없으면 False
                            {
                                vReport.Create();
                            }
                            ////------------------------------------------------------------------------
                            ////[Test Path]
                            ////------------------------------------------------------------------------
                            //string vPathTest = @"K:\00_2_FXE\ERPMain\FXEMain\bin\Debug";
                            //string vPathReportClient = string.Format("{0}\\{1}", vPathTest, "Report");
                            ////------------------------------------------------------------------------

                            vFileTransferAdv.SourceDirectory = vPathReportFTP;
                            vFileTransferAdv.SourceFileName = vReportFileName;
                            vFileTransferAdv.TargetDirectory = vPathReportClient;
                            vFileTransferAdv.TargetFileName = vReportFileNameTarget;

                            string vFullPathReportClient = string.Format("{0}\\{1}", vPathReportClient, vReportFileName);
                            string vFullPathReportTarget = string.Format("{0}\\{1}", vPathReportClient, vReportFileNameTarget);

                            if (vFileTransferAdv.Download() == true)
                            {
                                try
                                {
                                    System.IO.File.Delete(vFullPathReportClient);
                                    System.IO.File.Move(vFullPathReportTarget, vFullPathReportClient);
                                }
                                catch
                                {
                                    try
                                    {
                                        System.IO.FileInfo vFileInfo = new System.IO.FileInfo(vFullPathReportTarget);
                                        if (vFileInfo.Exists == true)
                                        {
                                            System.IO.File.Delete(vFullPathReportTarget);
                                        }
                                    }
                                    catch
                                    {
                                        //
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
            }
        }

        #endregion;

        #region ----- XL Print 1 Methods ----

        private void XLPrinting_Main()
        {
            object vSlip_Header_id;
            object vGL_Date;
            object vGL_Num;

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            
            vSlip_Header_id = SLIP_HEADER_ID.EditValue;
            vGL_Date = GL_DATE.EditValue;
            vGL_Num = GL_NUM.EditValue;

            //print type 설정
            DialogResult vdlgResult;
            FCMF0525_PRINT_TYPE vFCMF00525_PRINT_TYPE = new FCMF0525_PRINT_TYPE(isAppInterfaceAdv1.AppInterface);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF00525_PRINT_TYPE, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vFCMF00525_PRINT_TYPE.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            string vPRINT_TYPE = iString.ISNull(vFCMF00525_PRINT_TYPE.Get_Printer_Type);
            if (vPRINT_TYPE == string.Empty)
            {
                return;
            }
            vFCMF00525_PRINT_TYPE.Dispose();


            XLPrinting1(vPRINT_TYPE);
           
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting1(string pOutput_Type)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            int vCountRowGrid = IGR_BATCH_LIST.RowCount;
            if ((TB_BATCH.SelectedIndex == 0 && vCountRowGrid > 0) ||
                (TB_BATCH.SelectedIndex == 1 && iString.ISNull(BATCH_HEADER_ID.EditValue) != string.Empty))
            {
                vMessageText = string.Format("Printing Starting", vPageTotal);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0525_001.xlsx";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        object vObject;
                        int vCountRow = 0;
                        int vRow = IGR_BATCH_LIST.RowIndex;
                        if (TB_BATCH.SelectedTab.TabIndex == 2)
                        {
                            xlPrinting.HeaderWrite(IDA_BATCH_HEADER);
                            vObject = BATCH_HEADER_ID.EditValue;
                        }
                        else
                        {
                            xlPrinting.HeaderWrite_1(IGR_BATCH_LIST, vRow);
                            vObject = IGR_BATCH_LIST.GetCellValue("BATCH_HEADER_ID");
                        }
                        idaPRINT_SLIP_LINE.SetSelectParamValue("P_BATCH_HEADER_ID", vObject);
                        idaPRINT_SLIP_LINE.Fill();

                        vCountRow = idaPRINT_SLIP_LINE.CurrentRows.Count;
                        if (vCountRow > 0)
                        {
                            if (pOutput_Type == "PDF") { 
                                vPageNumber = xlPrinting.LineWrite2(idaPRINT_SLIP_LINE);
                            }
                             else
                            {
                                vPageNumber = xlPrinting.LineWrite(idaPRINT_SLIP_LINE);
                             }
                        }

                        if (pOutput_Type == "PRINTER")
                        {//[PRINT]
                            ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                            xlPrinting.PreView(1, vPageNumber);
                            
                        }
                        else if (pOutput_Type == "EXCEL")
                        {
                            ////[SAVE]
                            xlPrinting.Save("EXCEL_"); //저장 파일명
                        }
                        else if (pOutput_Type == "PREVIEW")
                        {
                            xlPrinting.PreView(1, vPageNumber);
                        }
                        else if (pOutput_Type == "PDF")
                        {
                           // xlPrinting.PreView(1, vPageNumber);

                            xlPrinting. PDF("PDF_");  //PDF 파일명
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
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
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
                //전표 행번호 보정위해 주석 
                //else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                //{
                //    if (IDA_BATCH_LINE.IsFocused)
                //    {
                //        IDA_BATCH_LINE.AddOver();
                //        InsertSlipLine();
                //    }
                //    else
                //    {
                //        if (Check_SlipHeader_Added() == true)
                //        {
                //            return;
                //        }
                //        else
                //        {
                //            IDA_BATCH_HEADER.SetSelectParamValue("W_BATCH_HEADER_ID", 0);
                //            IDA_BATCH_HEADER.Fill();

                //            IDA_BATCH_HEADER.AddOver();
                //            IDA_BATCH_LINE.AddOver();
                //            InsertSlipLine();
                //            InsertSlipHeader();
                //        }
                //    }
                //}
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BATCH_LINE.IsFocused)
                    {
                        IDA_BATCH_LINE.AddUnder();
                        InsertSlipLine();
                    }
                    else
                    {
                        if (Check_SlipHeader_Added() == true)
                        {
                            return;
                        }
                        else
                        {
                            IDA_BATCH_HEADER.SetSelectParamValue("W_BATCH_HEADER_ID", 0);
                            IDA_BATCH_HEADER.Fill();

                            IDA_BATCH_HEADER.AddUnder();
                            IDA_BATCH_LINE.AddUnder();                            
                            InsertSlipHeader();
                            InsertSlipLine();
                            BATCH_DATE.Focus();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_BATCH_LIST.IsFocused)
                    {
                        IDA_BATCH_LIST.Update();
                    }
                    else
                    {
                        ACCOUNT_CODE.Focus();
                        Init_DR_CR_Amount();
                        Init_Total_GL_Amount();

                        if (iString.ISDecimaltoZero(TOTAL_DR_AMOUNT.EditValue) != iString.ISDecimaltoZero(TOTAL_CR_AMOUNT.EditValue))
                        {// 차대금액 일치 여부 체크.
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10134"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        string vBATCH_CLOSED_YN = GET_BATCH_CLOSED_YN(BATCH_HEADER_ID.EditValue);
                        if (vBATCH_CLOSED_YN == "Y" || vBATCH_CLOSED_YN == "F")
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10408"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        IDA_BATCH_HEADER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    SLIP_QUERY_STATUS.EditValue = "QUERY";
                    if (IDA_BATCH_LIST.IsFocused)
                    {
                        IDA_BATCH_LIST.Cancel();
                    }
                    else if (IDA_BATCH_HEADER.IsFocused)
                    {
                        IDA_BATCH_LINE.Cancel();
                        IDA_BATCH_HEADER.Cancel();
                    }
                    else if (IDA_BATCH_LINE.IsFocused)
                    {
                        IDA_BATCH_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BATCH_LIST.IsFocused)
                    {
                        IDA_BATCH_LIST.Delete();
                    }
                    else if (IDA_BATCH_HEADER.IsFocused)
                    {
                        for (int r = 0; r < igrSLIP_LINE.RowCount; r++)
                        {
                            IDA_BATCH_LINE.Delete();
                        }
                        IDA_BATCH_HEADER.Delete();
                    }
                    else if (IDA_BATCH_LINE.IsFocused)
                    {
                        IDA_BATCH_LINE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main();
                    //XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main();
                    //XLPrinting1("EXCEL");
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0525_Load(object sender, EventArgs e)
        {
            IDA_BATCH_HEADER.FillSchema();
        }

        private void FCMF0525_Shown(object sender, EventArgs e)
        {
            ibtSUB_FORM.Visible = false;

            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            BATCH_DATE_FR_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            BATCH_DATE_TO_0.EditValue = iDate.ISGetDate();

            REF_SLIP_FLAG.BringToFront();

            // 회계장부 정보 설정.
            GetAccountBook();   

            TB_BATCH.SelectedIndex = 1;
            TB_BATCH.SelectedTab.Focus();
        }


        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (IDA_BATCH_LINE.CurrentRow != null && IDA_BATCH_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_GL_Amount();
            }
        }

        private void GL_CURR_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (IDA_BATCH_LINE.CurrentRow != null && IDA_BATCH_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_GL_Amount();
            }
        }

        private void GL_AMOUNT_EditValueChanged(object pSender)
        {
            if (IDA_BATCH_LINE.CurrentRow != null && IDA_BATCH_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_VAT_Amount();
            }
        }

        private void GL_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
        }

        private void IGR_BATCH_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_BATCH_LIST.RowCount > 0)
            {
                Search_DB_DETAIL(IGR_BATCH_LIST.GetCellValue("BATCH_HEADER_ID"));
            }
        }

        private void BATCH_DATE_EditValueChanged(object pSender)
        {
            if (BATCH_DATE.DataAdapter.IsEditing == true)
            {
                GL_DATE.EditValue = BATCH_DATE.EditValue;
            }
        }

        private void btnSET_BATCH_LIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BATCH_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BATCH_DATE.Focus();
                return;
            }

            if (iString.ISNull(CLOSED_YN.EditValue) == "Y")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult vRESULT;
            FCMF0525_SET vFCMF0525_SET = new FCMF0525_SET(isAppInterfaceAdv1.AppInterface, BATCH_DATE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0525_SET, isAppInterfaceAdv1.AppInterface);
            vRESULT = vFCMF0525_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                object vPAYMENT_DATE = vFCMF0525_SET.Get_Payment_Date;
                //Set_Delete_Batch_Line();
                if (iString.ISNull(ACCOUNT_CONTROL_ID.EditValue) == string.Empty || iString.ISNull(ACCOUNT_DR_CR.EditValue) == string.Empty)
                {
                    IDA_BATCH_LINE.Delete();
                }
                IDA_BATCH_LINE.MoveLast(igrSLIP_LINE.Name);
                Set_Insert_Batch_Line(vPAYMENT_DATE);
                Init_Currency_Code("Y");
                Init_Currency_Amount();

            }
            vFCMF0525_SET.Dispose();

            ACCOUNT_CODE.Focus();
        }

        private void ibtSUB_FORM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("ACCOUNT_DR_CR")) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_DR_CR.Focus();
                return;
            }
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("ACCOUNT_CONTROL_ID")) == string.Empty)
            {// 계정과목.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_CODE.Focus();
                return;
            }
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("CURRENCY_CODE")) == string.Empty)
            {// 통화
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CURRENCY_DESC.Focus();
                return;
            }
            if (mCurrency_Code.ToString() != igrSLIP_LINE.GetCellValue("CURRENCY_CODE").ToString()
                  && iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue("EXCHANGE_RATE")) == Convert.ToInt32(0))
            {// 입력통화와 기본 통화가 다를경우 환율입력 체크.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10125"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                EXCHANGE_RATE.Focus();
                return;
            }
            if (mCurrency_Code.ToString() != igrSLIP_LINE.GetCellValue("CURRENCY_CODE").ToString()
                  && iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue("GL_CURRENCY_AMOUNT")) == Convert.ToInt32(0))
            {// 입력통화와 기본 통화가 다를경우 외화금액 체크.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                GL_CURR_AMOUNT.Focus();
                return;
            }

            //if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "AP_VAT".ToString())
            //{// 부가세 대급금 체크 - 전표LINE ID가 존재해야 함.
            //    if (iString.ISNull(igrSLIP_LINE.GetCellValue("SLIP_LINE_ID")) == String.Empty)
            //    {// 금액
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10271"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}

            System.Windows.Forms.DialogResult dlgResult;
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "RECEIVABLE_BILL".ToString())
            {//받을어음
                object mBILL_CLASS = "2";  // 어음구분.
                object mBILL_NUM = Get_Management_Value("RECEIVABLE_BILL");
                object mBILL_AMOUNT = GL_AMOUNT.EditValue;
                object mVENDOR_CODE = Get_Management_Value("CUSTOMER");
                object mBANK_CODE = Get_Management_Value("BANK");
                object mVAT_ISSUE_DATE = Get_Management_Value("VAT_ISSUE_DATE");
                object mISSUE_DATE = Get_Management_Value("ISSUE_DATE");
                if (iString.ISNull(mISSUE_DATE) == string.Empty)
                {
                    mISSUE_DATE = GL_DATE.EditValue;
                }
                object mDUE_DATE = Get_Management_Value("DUE_DATE");
                if (iString.ISNull(mDUE_DATE) == string.Empty)
                {
                    mDUE_DATE = Get_Management_Value("TR_EXPIRATION_DATE");
                }
                object mDEPT_ID = DEPT_ID.EditValue;
                object mDEPT_NAME = DEPT_NAME.EditValue;

                FCMF0525_BILL vFCMF0525_BILL = new FCMF0525_BILL(isAppInterfaceAdv1.AppInterface, mDEPT_ID, mDEPT_NAME
                                                                    , mBILL_CLASS, mBILL_NUM, mBILL_AMOUNT
                                                                    , mVENDOR_CODE, mBANK_CODE
                                                                    , mVAT_ISSUE_DATE, mISSUE_DATE, mDUE_DATE
                                                                    , PERSON_ID.EditValue, PERSON_NAME.EditValue);
                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0525_BILL, isAppInterfaceAdv1.AppInterface);
                dlgResult = vFCMF0525_BILL.ShowDialog();
                if (dlgResult == DialogResult.OK)
                {
                    //어음금액
                    GL_AMOUNT.EditValue = vFCMF0525_BILL.Get_BILL_AMOUNT;
                    //거래처.
                    Set_Management_Value("CUSTOMER", vFCMF0525_BILL.Get_VENDOR_CODE, vFCMF0525_BILL.Get_VENDOR_NAME);
                    //은행
                    Set_Management_Value("BANK", vFCMF0525_BILL.Get_BANK_CODE, vFCMF0525_BILL.Get_BANK_NAME);
                    //세금계산서발행일
                    Set_Management_Value("VAT_ISSUE_DATE", vFCMF0525_BILL.Get_VAT_ISSUE_DATE, null);
                    //발행일자
                    Set_Management_Value("ISSUE_DATE", vFCMF0525_BILL.Get_ISSUE_DATE, null);
                    //만기일자
                    Set_Management_Value("DUE_DATE", vFCMF0525_BILL.Get_DUE_DATE, null);
                    //만기일자
                    Set_Management_Value("TR_EXPIRATION_DATE", vFCMF0525_BILL.Get_DUE_DATE, null);  
                    //어음번호.
                    Set_Management_Value("RECEIVABLE_BILL", vFCMF0525_BILL.Get_BILL_NUM, String.Format("{0:###,###,###,###,###,###}", vFCMF0525_BILL.Get_BILL_AMOUNT));

                    Init_DR_CR_Amount();    // 차대금액 생성 //
                    Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                }
                vFCMF0525_BILL.Dispose();
            }
            else if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "PAYABLE_BILL".ToString())
            {//지급어음
                object mBILL_CLASS = "1";  // 어음구분.
                object mBILL_NUM = Get_Management_Value("PAYABLE_BILL");
                object mBILL_AMOUNT = GL_AMOUNT.EditValue;
                object mVENDOR_CODE = Get_Management_Value("CUSTOMER");
                object mBANK_CODE = Get_Management_Value("BANK");
                object mVAT_ISSUE_DATE = Get_Management_Value("VAT_ISSUE_DATE");
                object mISSUE_DATE = Get_Management_Value("ISSUE_DATE");
                if (iString.ISNull(mISSUE_DATE) == string.Empty)
                {
                    mISSUE_DATE = GL_DATE.EditValue;
                }
                object mDUE_DATE = Get_Management_Value("DUE_DATE");
                if (iString.ISNull(mDUE_DATE) == string.Empty)
                {
                    mDUE_DATE = Get_Management_Value("TR_EXPIRATION_DATE");
                }
                object mDEPT_ID = DEPT_ID.EditValue;
                object mDEPT_NAME = DEPT_NAME.EditValue;

                FCMF0525_BILL vFCMF0525_BILL = new FCMF0525_BILL(isAppInterfaceAdv1.AppInterface, mDEPT_ID, mDEPT_NAME
                                                                    , mBILL_CLASS, mBILL_NUM, mBILL_AMOUNT
                                                                    , mVENDOR_CODE, mBANK_CODE
                                                                    , mVAT_ISSUE_DATE, mISSUE_DATE, mDUE_DATE
                                                                    , PERSON_ID.EditValue, PERSON_NAME.EditValue); 
                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0525_BILL, isAppInterfaceAdv1.AppInterface);
                dlgResult = vFCMF0525_BILL.ShowDialog();
                if (dlgResult == DialogResult.OK)
                {
                    //어음금액
                    GL_AMOUNT.EditValue = vFCMF0525_BILL.Get_BILL_AMOUNT;
                    //거래처.
                    Set_Management_Value("CUSTOMER", vFCMF0525_BILL.Get_VENDOR_CODE, vFCMF0525_BILL.Get_VENDOR_NAME);
                    //은행
                    Set_Management_Value("BANK", vFCMF0525_BILL.Get_BANK_CODE, vFCMF0525_BILL.Get_BANK_NAME);
                    //세금계산서발행일
                    Set_Management_Value("VAT_ISSUE_DATE", vFCMF0525_BILL.Get_VAT_ISSUE_DATE, null);
                    //발행일자
                    Set_Management_Value("ISSUE_DATE", vFCMF0525_BILL.Get_ISSUE_DATE, null);
                    //만기일자
                    Set_Management_Value("DUE_DATE", vFCMF0525_BILL.Get_DUE_DATE, null);
                    //만기일자
                    Set_Management_Value("TR_EXPIRATION_DATE", vFCMF0525_BILL.Get_DUE_DATE, null);  
                    //어음번호.
                    Set_Management_Value("PAYABLE_BILL", vFCMF0525_BILL.Get_BILL_NUM, String.Format("{0:###,###,###,###,###,###}", vFCMF0525_BILL.Get_BILL_AMOUNT));

                    Init_DR_CR_Amount();    // 차대금액 생성 //
                    Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                }
                vFCMF0525_BILL.Dispose();
            }
            //else if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "AP_VAT".ToString())
            //{
            //    S_SUPPLY_AMOUNT.EditValue = Get_Management_Value("SUPPLY_AMOUNT");   //공급가액 설정.
            //    S_VAT_AMOUNT.EditValue = Get_Management_Value("VAT_AMOUNT");      //세액 설정.

            //    //서브판넬 
            //    Init_Sub_Panel(true, "AP_VAT");
            //}
            //else if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "DEAL".ToString())
            //{//외화획득용 구매(공급) 확인서
            //    object mISSUE_NUM = Get_Management_Value("PC_ISSUE_NO");
            //    object mVENDOR_CODE = Get_Management_Value("CUSTOMER");
            //    object mBANK_CODE = Get_Management_Value("BANK");
            //    object mISSUE_DATE = Get_Management_Value("ISSUE_DATE");
            //    object mCURRENCY_CODE = CURRENCY_CODE.EditValue;

            //    FCMF0525_ITEM_DEAL vFCMF0525_ITEM_DEAL = new FCMF0525_ITEM_DEAL(isAppInterfaceAdv1.AppInterface, mISSUE_NUM, mCURRENCY_CODE
            //                                                                    , mVENDOR_CODE, mBANK_CODE, mISSUE_DATE);

            //    dlgResult = vFCMF0525_ITEM_DEAL.ShowDialog();
            //    if (dlgResult == DialogResult.OK)
            //    {
            //        //거래처.4
            //        Set_Management_Value("CUSTOMER", vFCMF0525_ITEM_DEAL.Get_VENDOR_CODE, vFCMF0525_ITEM_DEAL.Get_VENDOR_NAME);

            //        //구매(공급)확인번호
            //        Set_Management_Value("PC_ISSUE_NO", vFCMF0525_ITEM_DEAL.Get_ISSUE_NUM, DBNull.Value);

            //        Set_Management_Value("BANK", vFCMF0525_ITEM_DEAL.Get_BANK_CODE, vFCMF0525_ITEM_DEAL.Get_BANK_NAME);

            //        Set_Management_Value("ISSUE_DATE", vFCMF0525_ITEM_DEAL.Get_ISSUE_DATE, DBNull.Value);
            //    }
            //    vFCMF0525_ITEM_DEAL.Dispose();
            //}
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_OFFSET_ACCOUNT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Init_Offset_Account(igrSLIP_LINE.RowIndex);
        }

        private void BTN_SLIP_TRANS_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vBATCH_CLOSED_YN = GET_BATCH_CLOSED_YN(BATCH_HEADER_ID.EditValue);
            if (vBATCH_CLOSED_YN == "Y" || vBATCH_CLOSED_YN == "F")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10408"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_BATCH_HEADER.Update();

            if (iString.ISNull(BATCH_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}",Get_Edit_Prompt(BATCH_HEADER_ID))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult vdlgResult;
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10303"), "Questiong", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;

            isDataTransaction1.BeginTran();
            IDC_SLIP_TRANS.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_SLIP_TRANS.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_SLIP_TRANS.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
            if(IDC_SLIP_TRANS.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                MessageBoxAdv.Show(mMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            isDataTransaction1.Commit();

            Search_DB();
        }

        private void BTN_SLIP_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BATCH_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BATCH_HEADER_ID))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult vdlgResult;
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Questiong", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
            
            isDataTransaction1.BeginTran();
            IDC_SLIP_CANCEL.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_SLIP_CANCEL.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_SLIP_CANCEL.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
            if (IDC_SLIP_CANCEL.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                MessageBoxAdv.Show(mMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            isDataTransaction1.Commit();

            Search_DB();
        }

        #endregion

        #region ----- Lookup Event ----- 
        
        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaSLIP_NUM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
        }

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", " VALUE1 <> 'BL'", "Y");
        }

        private void ilaSLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SLIP_TYPE_SLIP_DOCU.SetLookupParamValue("P_ENABLED_FLAG", "Y"); 
        }

        private void ilaREQ_PAYABLE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("PAYABLE_TYPE", "Y");
        }

        private void ilaBUDGET_DEPT_L_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBUDGET_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildBUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", GL_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaACCOUNT_DR_CR_SelectedRowData(object pSender)
        {
            Init_DR_CR_Amount();
            Init_Total_GL_Amount();
            GetSubForm();
        }

        private void ilaACCOUNT_CONTROL_SelectedRowData(object pSender)
        {
            Init_Currency_Code("Y");
            Set_Control_Item_Prompt();
            Init_Control_Management_Value();
            Init_Set_Item_Prompt(IDA_BATCH_LINE.CurrentRow);
            Init_Set_Item_Need(IDA_BATCH_LINE.CurrentRow);
            if (IDA_BATCH_LINE.CurrentRow.RowState != DataRowState.Modified)
            {
                Init_Default_Value();
            }
            Init_Default_Management("DEPT");
            Init_Default_Management("TAX_CODE");
            GetSubForm();
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_SelectedRowData(object pSender)
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) != string.Empty)
            {
                Init_Currency_Amount();
                if (CURRENCY_CODE.EditValue.ToString() != mCurrency_Code.ToString())
                {
                    idcEXCHANGE_RATE.ExecuteNonQuery();
                    EXCHANGE_RATE.EditValue = idcEXCHANGE_RATE.GetCommandParamValue("X_EXCHANGE_RATE");

                    Init_GL_Amount();
                    EXCHANGE_RATE.Focus();
                }
            }
        }

        private void ILA_SLIP_REMARK_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_SLIP_REMARK.SetLookupParamValue("P_ENABLED_FLAG", "Y");
            ILD_SLIP_REMARK.SetLookupParamValue("P_ENABLED_DATE", GL_DATE.EditValue);
        }

        private void ILA_SLIP_REMARK_SelectedRowData(object pSender)
        {
            if (iString.ISNull(REMARK.EditValue) != string.Empty)
            {
                REMARK.TextSelectionStart = iString.ISNull(REMARK.EditValue).Length;
                REMARK.Focus();
            }
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT1_ID", "Y", igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT2_ID", "Y", igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
        }

        private void ilaREFER1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER1_ID", "Y", igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
        }

        private void ilaREFER2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER2_ID", "Y", igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
        }

        private void ilaREFER3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER3_ID", "Y", igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
        }

        private void ilaREFER4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER4_ID", "Y", igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
        }

        private void ilaREFER5_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER5_ID", "Y", igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
        }

        private void ilaREFER6_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER6_ID", "Y", igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
        }

        private void ilaREFER7_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER7_ID", "Y", igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
        }

        private void ilaREFER8_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER8_ID", "Y", igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
        }


        private void ilaMANAGEMENT1_SelectedRowData(object pSender)
        {// 관리항목1 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT1");

            //관리항목 동기화// 
            //거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "DUE_DATE", MANAGEMENT1.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "PAYMENT_METHOD", MANAGEMENT1.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("MANAGEMENT1", "CREDIT_CARD", "DUE_DATE", MANAGEMENT1.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("MANAGEMENT1", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", MANAGEMENT1.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaMANAGEMENT2_SelectedRowData(object pSender)
        {// 관리항목2 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT2");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("MANAGEMENT2", "CUSTOMER", "DUE_DATE", MANAGEMENT2.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("MANAGEMENT2", "CUSTOMER", "PAYMENT_METHOD", MANAGEMENT2.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("MANAGEMENT2", "CREDIT_CARD", "DUE_DATE", MANAGEMENT2.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("MANAGEMENT2", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", MANAGEMENT2.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER1_SelectedRowData(object pSender)
        {// 관리항목3 선택시 적용.
            Init_SELECT_LOOKUP("REFER1");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER1", "CUSTOMER", "DUE_DATE", REFER1.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER1", "CUSTOMER", "PAYMENT_METHOD", REFER1.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER1", "CREDIT_CARD", "DUE_DATE", REFER1.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER1", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER1.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER2_SelectedRowData(object pSender)
        {// 관리항목4 선택시 적용.
            Init_SELECT_LOOKUP("REFER2");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER2", "CUSTOMER", "DUE_DATE", REFER2.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER2", "CUSTOMER", "PAYMENT_METHOD", REFER2.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER2", "CREDIT_CARD", "DUE_DATE", REFER2.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER2", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER2.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER3_SelectedRowData(object pSender)
        {// 관리항목5 선택시 적용.
            Init_SELECT_LOOKUP("REFER3");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER3", "CUSTOMER", "DUE_DATE", REFER3.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER3", "CUSTOMER", "PAYMENT_METHOD", REFER3.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER3", "CREDIT_CARD", "DUE_DATE", REFER3.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER3", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER3.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER4_SelectedRowData(object pSender)
        {// 관리항목6 선택시 적용.
            Init_SELECT_LOOKUP("REFER4");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER4", "CUSTOMER", "DUE_DATE", REFER4.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER4", "CUSTOMER", "PAYMENT_METHOD", REFER4.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER4", "CREDIT_CARD", "DUE_DATE", REFER4.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER4", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER4.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER5_SelectedRowData(object pSender)
        {// 관리항목7 선택시 적용.
            Init_SELECT_LOOKUP("REFER5");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER5", "CUSTOMER", "DUE_DATE", REFER5.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER5", "CUSTOMER", "PAYMENT_METHOD", REFER5.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER5", "CREDIT_CARD", "DUE_DATE", REFER5.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER5", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER5.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER6_SelectedRowData(object pSender)
        {// 관리항목8 선택시 적용.
            Init_SELECT_LOOKUP("REFER6");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER6", "CUSTOMER", "DUE_DATE", REFER6.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER6", "CUSTOMER", "PAYMENT_METHOD", REFER6.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER6", "CREDIT_CARD", "DUE_DATE", REFER6.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER6", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER6.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER7_SelectedRowData(object pSender)
        {// 관리항목9 선택시 적용.
            Init_SELECT_LOOKUP("REFER7");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER7", "CUSTOMER", "DUE_DATE", REFER7.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER7", "CUSTOMER", "PAYMENT_METHOD", REFER7.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER7", "CREDIT_CARD", "DUE_DATE", REFER7.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER7", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER7.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER8_SelectedRowData(object pSender)
        {// 관리항목10 선택시 적용.
            Init_SELECT_LOOKUP("REFER8");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("REFER8", "CUSTOMER", "DUE_DATE", REFER8.EditValue);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value("REFER8", "CUSTOMER", "PAYMENT_METHOD", REFER8.EditValue);
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER8", "CREDIT_CARD", "DUE_DATE", REFER8.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER8", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER8.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        #endregion       

        #region ----- Adapter Event -----

        private void idaSLIP_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["SLIP_TYPE"]) == string.Empty)
            {// 전표유형
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10116"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (iString.ISNull(e.Row["BATCH_DATE"]) == string.Empty)
            {// 기표일자.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10117"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["GL_DATE"]) == string.Empty)
            {// 전표일자.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            // 전표번호 채번//
            if (iString.ISNull(BATCH_NUM.EditValue) == string.Empty) // || iString.ISNull(e.Row["BATCH_DATE"]).Substring(0, 7) != iString.ISNull(e.Row["OLD_GL_DATE"], e.Row["GL_DATE"]).Substring(0, 7))
            {
                GetBatchNum();
            }

            //// 전표번호 채번//
            //if (iString.ISNull(GL_NUM.EditValue) == string.Empty || iString.ISNull(e.Row["GL_DATE"]).Substring(0, 7) != iString.ISNull(e.Row["OLD_GL_DATE"], e.Row["GL_DATE"]).Substring(0, 7))
            //{
            //    GetBatchNum();
            //}

            if (iString.ISNull(e.Row["BATCH_NUM"]) == string.Empty)
            {// 기표번호
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10118"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {// 발의부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10119"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {// 발의부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10121"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
        }

        private void idaSLIP_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                if (e.Row["CLOSED_YN"].ToString() == "Y".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaSLIP_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ACCOUNT_DR_CR"]) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {// 계정과목.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CODE"]) == string.Empty)
            {// 계정과목
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {// 통화
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_ENABLED_FLAG"]) == "Y".ToString())
            {// 외화 계좌.
                if (mCurrency_Code.ToString() != e.Row["CURRENCY_CODE"].ToString() && iString.ISDecimaltoZero(e.Row["EXCHANGE_RATE"]) == Convert.ToInt32(0))
                {// 입력통화와 기본 통화가 다를경우 환율입력 체크.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10125"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (iString.ISNull(e.Row["MANAGEMENT1"]) == string.Empty && iString.ISNull(e.Row["MANAGEMENT1_YN"], "N") == "Y".ToString())
            {// 관리항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["MANAGEMENT1_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["MANAGEMENT2"]) == string.Empty && iString.ISNull(e.Row["MANAGEMENT2_YN"], "N") == "Y".ToString())
            {// 관리항목2 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["MANAGEMENT2_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER1"]) == string.Empty && iString.ISNull(e.Row["REFER1_YN"], "N") == "Y".ToString())
            {// 참고항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER1_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER2"]) == string.Empty && iString.ISNull(e.Row["REFER2_YN"], "N") == "Y".ToString())
            {// 참고항목2 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER2_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER3"]) == string.Empty && iString.ISNull(e.Row["REFER3_YN"], "N") == "Y".ToString())
            {// 참고항목3 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER3_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER4"]) == string.Empty && iString.ISNull(e.Row["REFER4_YN"], "N") == "Y".ToString())
            {// 참고항목4 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER4_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER5"]) == string.Empty && iString.ISNull(e.Row["REFER5_YN"], "N") == "Y".ToString())
            {// 참고항목5 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER5_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER6"]) == string.Empty && iString.ISNull(e.Row["REFER6_YN"], "N") == "Y".ToString())
            {// 참고항목6 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER6_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER7"]) == string.Empty && iString.ISNull(e.Row["REFER7_YN"], "N") == "Y".ToString())
            {// 참고항목7 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER7_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REFER8"]) == string.Empty && iString.ISNull(e.Row["REFER8_YN"], "N") == "Y".ToString())
            {// 참고항목8 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", e.Row["REFER8_NAME"])), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaSLIP_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            try
            {
                if (e.Row.RowState != DataRowState.Added)
                {
                    if (e.Row["CLOSED_YN"].ToString() == "Y".ToString())
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        return;
                    }
                }
            }
            catch(Exception ex)
            {
                IDA_BATCH_LINE.MoveFirst(this.Name);
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private void idaSLIP_HEADER_UpdateCompleted(object pSender)
        {
            string vGL_NUM = iString.ISNull(GL_NUM.EditValue); // igrSLIP_LIST.GetCellValue("GL_NUM"));
            int vIDX_GL_NUM = IGR_BATCH_LIST.GetColumnToIndex("GL_NUM");
            Search_DB();

            // 기존 위치 이동 : 없을 경우.
            for (int r = 0; r < IGR_BATCH_LIST.RowCount; r++)
            {
                if (vGL_NUM == iString.ISNull(IGR_BATCH_LIST.GetCellValue(r, vIDX_GL_NUM)))
                {
                    IGR_BATCH_LIST.CurrentCellMoveTo(r, vIDX_GL_NUM);
                    IGR_BATCH_LIST.CurrentCellActivate(r, vIDX_GL_NUM);
                }
            }
            SLIP_TYPE_NAME.Focus();
        }

        private void idaSLIP_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Currency_Code("Y");
            Init_Currency_Amount();
            GetSubForm();
            Init_Total_GL_Amount();
        }

        private void idaSLIP_LINE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Set_Item_Prompt(pBindingManager.DataRow);
            Init_Set_Item_Need(pBindingManager.DataRow);
        }


        #endregion
    }
}