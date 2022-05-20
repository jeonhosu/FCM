using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0522
{
    public partial class FCMF0522_ADD : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        object mCurrency_Code;
        object mBudget_Control_YN;

        #endregion;

        #region ----- Constructor -----

        public FCMF0522_ADD(ISAppInterface pAppInterface, object pBALANCE_DATE, object pACCOUNT_CONTROL_ID, object pACCOUNT_CODE, object pACCOUNT_DESC)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            BALANCE_DATE.EditValue = pBALANCE_DATE;
            ACCOUNT_CONTROL_ID_0.EditValue = pACCOUNT_CONTROL_ID;
            ACCOUNT_CODE_0.EditValue = pACCOUNT_CODE;
            ACCOUNT_DESC_0.EditValue = pACCOUNT_DESC;
        }

        #endregion;

        #region ----- Private Methods ----

        private void GetAccountBook()
        {
            idcACCOUNT_BOOK.ExecuteNonQuery();
            mAccount_Book_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = idcACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = idcACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE");
            mBudget_Control_YN = idcACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private Boolean CheckData()
        {
            //if (iString.ISNull(LAST_CLOSED_DATE.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    GL_DATE.Focus();
            //    return false;
            //}
            if (iString.ISNull(BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BALANCE_DATE.Focus();
                return false;
            }
            //if (Convert.ToDateTime(LAST_CLOSED_DATE.EditValue) >= Convert.ToDateTime(GL_DATE.EditValue))
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    GL_DATE.Focus();
            //    return false;
            //}
            return true;
        }

        private void SEARCH_DB()
        {
            if (iString.ISNull(BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BALANCE_DATE.Focus();
                return;
            }
            IDA_BALANCE_STATEMENT_ADD.Fill();
            CURRENCY_CODE.Focus();
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            CURR_REMAIN_AMOUNT.Insertable = false;
            CURR_REMAIN_AMOUNT.Updatable = false;
            GL_DATE.Insertable = false;
            GL_DATE.Updatable = false;
            SLIP_REMARK.Insertable = false;
            SLIP_REMARK.Updatable = false;

            IDA_ITEM_PROMPT.Fill();
            if (IDA_ITEM_PROMPT.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            object mENABLED_FLAG;       // 사용(표시)여부.
            
            // 외화금액 - 통화관리 하는 경우 적용.
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["CONTROL_CURRENCY_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "Y".ToString())
            {
                CURR_REMAIN_AMOUNT.Insertable = true;
                CURR_REMAIN_AMOUNT.Updatable = true;
            }
            CURR_REMAIN_AMOUNT.Refresh();

            // 전표일자
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "Y".ToString())
            {
                GL_DATE.EditValue = BALANCE_DATE.EditValue;
                GL_DATE.Insertable = true;
                GL_DATE.Updatable = true;
            }
            GL_DATE.Refresh();

            // 적요
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "Y".ToString())
            {
                SLIP_REMARK.Insertable = true;
                SLIP_REMARK.Updatable = true;
            }
            SLIP_REMARK.Refresh();
        }

        //관리항목 LOOKUP 선택시 처리.
        private void Init_SELECT_LOOKUP(object pManagement_Type)
        {
            string mMANAGEMENT = iString.ISNull(pManagement_Type);
        }

        private void SetManagementParameter(string pManagement_Field, string pEnabled_YN, object pLookup_Type)
        {
            if (iString.ISNull(pLookup_Type) == "COSTCENTER".ToString())
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
                if (iString.ISNull(BALANCE_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = BALANCE_DATE.DateTimeValue.ToShortDateString();
                }
                else if (iString.ISNull(BALANCE_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = BALANCE_DATE.DateTimeValue.ToShortDateString();
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
            if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT1.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT2.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER1_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER1_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER1.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER2_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER2_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER2.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER3_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER3_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER3.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER4_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER4_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER4.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER5_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER5_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER5.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER6_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER6_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER6.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER7_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER7_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER7.EditValue;
            }
            else if (iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER8_LOOKUP_TYPE")) != string.Empty
                && iString.ISNull(IGR_BALANCE_DETAIL.GetCellValue("REFER8_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER8.EditValue;
            }
            else
            {
                mLookup_Value = null;
            }
            return mLookup_Value;
        }

        private void Init_Set_Item_Prompt(DataRow pDataRow)
        {// edit 데이터 형식, 사용여부 변경.
            if (pDataRow == null)
            {
                return;
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            CURRENCY_CODE.Nullable = true;
            if (iString.ISNull(pDataRow["CONTROL_CURRENCY_YN"], "N") == "Y".ToString())
            {
                CURRENCY_CODE.Nullable = false;
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT1.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["MANAGEMENT1_YN"], "F") == "F".ToString())
            {
                MANAGEMENT1.Nullable = true;
                MANAGEMENT1.ReadOnly = true;
                MANAGEMENT1.Insertable = false;
                MANAGEMENT1.Updatable = false;
                MANAGEMENT1.TabStop = false;
            }
            else
            {
                MANAGEMENT1.ReadOnly = false;
                MANAGEMENT1.Insertable = true;
                MANAGEMENT1.Updatable = true;
                MANAGEMENT1.TabStop = true;
                if (iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]) == "RATE".ToString())
                {
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT1.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]) == "DATE".ToString())
                {
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["MANAGEMENT1_YN"], "N") == "Y".ToString())
                {
                    MANAGEMENT1.ReadOnly = false;
                }
            }
            MANAGEMENT1.Refresh();

            MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT2.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["MANAGEMENT2_YN"], "F") == "F".ToString())
            {
                MANAGEMENT2.Nullable = true;
                MANAGEMENT2.ReadOnly = true;
                MANAGEMENT2.Insertable = false;
                MANAGEMENT2.Updatable = false;
                MANAGEMENT2.TabStop = false;
            }
            else
            {
                MANAGEMENT2.ReadOnly = false;
                MANAGEMENT2.Insertable = true;
                MANAGEMENT2.Updatable = true;
                MANAGEMENT2.TabStop = true;
                if (iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]) == "RATE".ToString())
                {
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT2.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]) == "DATE".ToString())
                {
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                else
                    if (iString.ISNull(pDataRow["MANAGEMENT2_YN"], "N") == "Y".ToString())
                    {
                        MANAGEMENT2.ReadOnly = false;
                    }
            }
            MANAGEMENT2.Refresh();

            REFER1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER1.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER1_YN"], "F") == "F".ToString())
            {
                REFER1.Nullable = true;
                REFER1.ReadOnly = true;
                REFER1.Insertable = false;
                REFER1.Updatable = false;
                REFER1.TabStop = false;
            }
            else
            {
                REFER1.ReadOnly = false;
                REFER1.Insertable = true;
                REFER1.Updatable = true;
                REFER1.TabStop = true;
                if (iString.ISNull(pDataRow["REFER1_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER1_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER1.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER1_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER1_YN"], "N") == "Y".ToString())
                {
                    REFER1.ReadOnly = false;
                }
            }
            REFER1.Refresh();

            REFER2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER2.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER2_YN"], "F") == "F".ToString())
            {
                REFER2.Nullable = true;
                REFER2.ReadOnly = true;
                REFER2.Insertable = false;
                REFER2.Updatable = false;
                REFER2.TabStop = false;
            }
            else
            {
                REFER2.ReadOnly = false;
                REFER2.Insertable = true;
                REFER2.Updatable = true;
                REFER2.TabStop = true;
                if (iString.ISNull(pDataRow["REFER2_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER2_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER2.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER2_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER2_YN"], "N") == "Y".ToString())
                {
                    REFER2.ReadOnly = false;
                }
            }
            REFER2.Refresh();

            REFER3.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER3.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER3_YN"], "F") == "F".ToString())
            {
                REFER3.Nullable = true;
                REFER3.ReadOnly = true;
                REFER3.Insertable = false;
                REFER3.Updatable = false;
                REFER3.TabStop = false;
            }
            else
            {
                REFER3.ReadOnly = false;
                REFER3.Insertable = true;
                REFER3.Updatable = true;
                REFER3.TabStop = true;
                if (iString.ISNull(pDataRow["REFER3_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER3_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER3.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER3_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER3_YN"], "N") == "Y".ToString())
                {
                    REFER3.ReadOnly = false;
                }
            }
            REFER3.Refresh();

            REFER4.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER4.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER4_YN"], "F") == "F".ToString())
            {
                REFER4.Nullable = true;
                REFER4.ReadOnly = true;
                REFER4.Insertable = false;
                REFER4.Updatable = false;
                REFER4.TabStop = false;
            }
            else
            {
                REFER4.ReadOnly = false;
                REFER4.Insertable = true;
                REFER4.Updatable = true;
                REFER4.TabStop = true;
                if (iString.ISNull(pDataRow["REFER4_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER4_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER4.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER4_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER4_YN"], "N") == "Y".ToString())
                {
                    REFER4.ReadOnly = false;
                }
            }
            REFER4.Refresh();

            REFER5.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER5.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER5_YN"], "F") == "F".ToString())
            {
                REFER5.Nullable = true;
                REFER5.ReadOnly = true;
                REFER5.Insertable = false;
                REFER5.Updatable = false;
                REFER5.TabStop = false;
            }
            else
            {
                REFER5.ReadOnly = false;
                REFER5.Insertable = true;
                REFER5.Updatable = true;
                REFER5.TabStop = true;
                if (iString.ISNull(pDataRow["REFER5_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER5_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER5.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER5_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER5_YN"], "N") == "Y".ToString())
                {
                    REFER5.ReadOnly = false;
                }
            }
            REFER5.Refresh();

            REFER6.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER6.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER6_YN"], "F") == "F".ToString())
            {
                REFER6.Nullable = true;
                REFER6.ReadOnly = true;
                REFER6.Insertable = false;
                REFER6.Updatable = false;
                REFER6.TabStop = false;
            }
            else
            {
                REFER6.ReadOnly = false;
                REFER6.Insertable = true;
                REFER6.Updatable = true;
                REFER6.TabStop = true;
                if (iString.ISNull(pDataRow["REFER6_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER6_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER6.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER6_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER6_YN"], "N") == "Y".ToString())
                {
                    REFER6.ReadOnly = false;
                }
            }
            REFER6.Refresh();

            REFER7.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER7.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER7_YN"], "F") == "F".ToString())
            {
                REFER7.Nullable = true;
                REFER7.ReadOnly = true;
                REFER7.Insertable = false;
                REFER7.Updatable = false;
                REFER7.TabStop = false;
            }
            else
            {
                REFER7.ReadOnly = false;
                REFER7.Insertable = true;
                REFER7.Updatable = true;
                REFER7.TabStop = true;
                if (iString.ISNull(pDataRow["REFER7_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER7_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER7.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER7_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER7_YN"], "N") == "Y".ToString())
                {
                    REFER7.ReadOnly = false;
                }
            }
            REFER7.Refresh();

            REFER8.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            REFER8.NumberDecimalDigits = 0;
            if (iString.ISNull(pDataRow["REFER8_YN"], "F") == "F".ToString())
            {
                REFER8.Nullable = true;
                REFER8.ReadOnly = true;
                REFER8.Insertable = false;
                REFER8.Updatable = false;
                REFER8.TabStop = false;
            }
            else
            {
                REFER8.ReadOnly = false;
                REFER8.Insertable = true;
                REFER8.Updatable = true;
                REFER8.TabStop = true;
                if (iString.ISNull(pDataRow["REFER8_DATA_TYPE"]) == "NUMBER".ToString())
                {
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                }
                else if (iString.ISNull(pDataRow["REFER8_DATA_TYPE"]) == "RATE".ToString())
                {
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER8.NumberDecimalDigits = 4;
                }
                else if (iString.ISNull(pDataRow["REFER8_DATA_TYPE"]) == "DATE".ToString())
                {
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                }
                if (iString.ISNull(pDataRow["REFER8_YN"], "N") == "Y".ToString())
                {
                    REFER8.ReadOnly = false;
                }
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
            object mDATA_TYPE;
            object mDR_CR_YN = "N";
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            //--1
            mDATA_VALUE = MANAGEMENT1.EditValue;
            MANAGEMENT1.Nullable = true;
            mDATA_TYPE = pDataRow["MANAGEMENT1_DATA_TYPE"];
            mDR_CR_YN = pDataRow["MANAGEMENT1_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                MANAGEMENT1.Nullable = false;
            }
            MANAGEMENT1.EditValue = mDATA_VALUE;
            MANAGEMENT1.Refresh();
            //--2
            mDATA_VALUE = MANAGEMENT2.EditValue;
            MANAGEMENT2.Nullable = true;
            mDATA_TYPE = pDataRow["MANAGEMENT2_DATA_TYPE"];
            mDR_CR_YN = pDataRow["MANAGEMENT2_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                MANAGEMENT2.Nullable = false;
            }
            MANAGEMENT2.Refresh();
            MANAGEMENT2.EditValue = mDATA_VALUE;
            //--3
            mDATA_VALUE = REFER1.EditValue;
            REFER1.Nullable = true;
            mDATA_TYPE = pDataRow["REFER1_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER1_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER1.Nullable = false;
            }
            REFER1.Refresh();
            REFER1.EditValue = mDATA_VALUE;
            //--4
            mDATA_VALUE = REFER2.EditValue;
            REFER2.Nullable = true;
            mDATA_TYPE = pDataRow["REFER2_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER2_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER2.Nullable = false;
            }
            REFER2.Refresh();
            REFER2.EditValue = mDATA_VALUE;
            //--5
            mDATA_VALUE = REFER3.EditValue;
            REFER3.Nullable = true;
            mDATA_TYPE = pDataRow["REFER3_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER3_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER3.Nullable = false;
            }
            REFER3.Refresh();
            REFER3.EditValue = mDATA_VALUE;
            //--6
            mDATA_VALUE = REFER4.EditValue;
            REFER4.Nullable = true;
            mDATA_TYPE = pDataRow["REFER4_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER4_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER4.Nullable = false;
            }
            REFER4.Refresh();
            REFER4.EditValue = mDATA_VALUE;
            //--7
            mDATA_VALUE = REFER5.EditValue;
            REFER5.Nullable = true;
            mDATA_TYPE = pDataRow["REFER5_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER5_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER5.Nullable = false;
            }
            REFER5.Refresh();
            REFER5.EditValue = mDATA_VALUE;
            //--8
            mDATA_VALUE = REFER6.EditValue;
            REFER6.Nullable = true;
            mDATA_TYPE = pDataRow["REFER6_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER6_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER6.Nullable = false;
            }
            REFER6.Refresh();
            REFER6.EditValue = mDATA_VALUE;
            //--9
            mDATA_VALUE = REFER7.EditValue;
            REFER7.Nullable = true;
            mDATA_TYPE = pDataRow["REFER7_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER7_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER7.Nullable = false;
            }
            REFER7.Refresh();
            REFER7.EditValue = mDATA_VALUE;
            //--10
            mDATA_VALUE = REFER8.EditValue;
            REFER8.Nullable = true;
            mDATA_TYPE = pDataRow["REFER8_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER8_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = IGR_BALANCE_DETAIL.GetCellValue("REFER8_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = IGR_BALANCE_DETAIL.GetCellValue("REFER8_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER8.Nullable = false;
            }
            REFER8.Refresh();
            REFER8.EditValue = mDATA_VALUE;
        }

        private void Set_Control_Item_Prompt()
        {
            idaCONTROL_ITEM_PROMPT.Fill();
            if (idaCONTROL_ITEM_PROMPT.OraSelectData.Rows.Count > 0)
            {
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_NAME"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_NAME"]);

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_YN"]);

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_YN"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_YN"]);

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_TYPE"]);

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_DATA_TYPE"]);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_DATA_TYPE"]);
            }
            else
            {
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_NAME", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_NAME", null);

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_YN", "F");
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_YN", "F");

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_LOOKUP_YN", "N");
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_LOOKUP_YN", "N");

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_LOOKUP_TYPE", null);
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_LOOKUP_TYPE", null);

                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER1_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER2_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER3_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER4_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER5_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER6_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER7_DATA_TYPE", "VARCHAR2");
                IGR_BALANCE_DETAIL.SetCellValue("REFER8_DATA_TYPE", "VARCHAR2");
            }
        }

        private void Init_Control_Management_Value()
        {            
            IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1", null);
            IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT1_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2", null);
            IGR_BALANCE_DETAIL.SetCellValue("MANAGEMENT2_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER1", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER1_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER2", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER2_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER3", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER3_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER4", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER4_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER5", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER5_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER6", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER6_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER7", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER7_DESC", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER8", null);
            IGR_BALANCE_DETAIL.SetCellValue("REFER8_DESC", null);
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private object GetTerritory()
        {

            object vTerritory = "Default";
            vTerritory = isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage;
            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            try
            {
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
            }
            catch
            {
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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
            }
        }

        #endregion;
        
        #region ----- From Event -----

        private void FCMF0522_ADD_Load(object sender, EventArgs e)
        {
            IDA_BALANCE_STATEMENT_ADD.FillSchema();
        }

        private void FCMF0522_ADD_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();

            GetAccountBook();
            INIT_MANAGEMENT_COLUMN();
            Set_Control_Item_Prompt();
            Init_Control_Management_Value();
            Init_Set_Item_Prompt(IDA_BALANCE_STATEMENT_ADD.CurrentRow);
            Init_Set_Item_Need(IDA_BALANCE_STATEMENT_ADD.CurrentRow);
            CURRENCY_CODE.Focus();
        }

        private void ibtnUPDATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_BALANCE_STATEMENT_ADD.Update();
        }

        private void ibtnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_BALANCE_STATEMENT_ADD.Cancel();

            this.DialogResult = DialogResult.Cancel;
            this.Close();            
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT1_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT2_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
        }

        private void ilaREFER1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER1_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER1_LOOKUP_TYPE"));
        }

        private void ilaREFER2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER2_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER2_LOOKUP_TYPE"));
        }

        private void ilaREFER3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER3_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER3_LOOKUP_TYPE"));
        }

        private void ilaREFER4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER4_ID", "N", IGR_BALANCE_DETAIL.GetCellValue("REFER4_LOOKUP_TYPE"));
        }

        private void ilaREFER5_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER5_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER5_LOOKUP_TYPE"));
        }

        private void ilaREFER6_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER6_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER6_LOOKUP_TYPE"));
        }

        private void ilaREFER7_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER7_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER7_LOOKUP_TYPE"));
        }

        private void ilaREFER8_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER8_ID", "Y", IGR_BALANCE_DETAIL.GetCellValue("REFER8_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT1_SelectedRowData(object pSender)
        {// 관리항목1 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT1");
        }

        private void ilaMANAGEMENT2_SelectedRowData(object pSender)
        {// 관리항목2 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT2");
        }

        private void ilaREFER1_SelectedRowData(object pSender)
        {// 관리항목3 선택시 적용.
            Init_SELECT_LOOKUP("REFER1");
        }

        private void ilaREFER2_SelectedRowData(object pSender)
        {// 관리항목4 선택시 적용.
            Init_SELECT_LOOKUP("REFER2");
        }

        private void ilaREFER3_SelectedRowData(object pSender)
        {// 관리항목5 선택시 적용.
            Init_SELECT_LOOKUP("REFER3");
        }

        private void ilaREFER4_SelectedRowData(object pSender)
        {// 관리항목6 선택시 적용.
            Init_SELECT_LOOKUP("REFER4");
        }

        private void ilaREFER5_SelectedRowData(object pSender)
        {// 관리항목7 선택시 적용.
            Init_SELECT_LOOKUP("REFER5");
        }

        private void ilaREFER6_SelectedRowData(object pSender)
        {// 관리항목8 선택시 적용.
            Init_SELECT_LOOKUP("REFER6");
        }

        private void ilaREFER7_SelectedRowData(object pSender)
        {// 관리항목9 선택시 적용.
            Init_SELECT_LOOKUP("REFER7");
        }

        private void ilaREFER8_SelectedRowData(object pSender)
        {// 관리항목10 선택시 적용.
            Init_SELECT_LOOKUP("REFER8");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_BALANCE_STATEMENT_ADD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BALANCE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (iString.ISNull(ACCOUNT_CONTROL_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_CODE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 

            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CURRENCY_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISDecimaltoZero(e.Row["REMAIN_AMOUNT"]) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REMAIN_AMOUNT))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["GL_DATE"]) == string.Empty && iString.ISNull(e.Row["GL_DATE_YN"], "N") == "Y".ToString())
            {// 관리항목1 필수 입력 체크
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(GL_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["CONTROL_CURRENCY_YN"]) == "Y".ToString())
            {// 외화 계좌.
                if (mCurrency_Code.ToString() != e.Row["CURRENCY_CODE"].ToString() && iString.ISDecimaltoZero(e.Row["CURR_REMAIN_AMOUNT"]) == Convert.ToInt32(0))
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

        private void IDA_BALANCE_STATEMENT_ADD_UpdateCompleted(object pSender)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void IDA_BALANCE_STATEMENT_ADD_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
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