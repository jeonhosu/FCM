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

using System.IO;
using Syncfusion.GridExcelConverter;

namespace FCMF0316
{
    public partial class FCMF0316_ACCOUNT : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
         
        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mBase_Currency_Code; 
        object mBudget_Control_YN; 

        object mDEPT_ID;
        object mDEPT_NAME;
        object mCURRENCY_CODE;
        object mEXCHANGE_RATE;
        object mGL_CURRENCY_AMOUNT;
        object mGL_AMOUNT;

        object mVENDOR_CODE;
        object mVENDOR_NAME;
        object mVAT_TAX_TYPE;
        object mVAT_TAX_TYPE_NAME;

        #endregion;

        #region ----- Constructor -----

        public FCMF0316_ACCOUNT()
        {
            InitializeComponent();
        }

        public FCMF0316_ACCOUNT(Form pMainForm, ISAppInterface pAppInterface, object pSALE_HEADER_ID, string pSLIP_LINE_TYPE,
                                object pSLIP_DATE, object pDEPT_ID, object pDEPT_NAME, object pBUDGET_DEPT_ID, object pUSE_DEPT_ID, 
                                object pCURRENCY_CODE, object pEXCHANGE_RATE, object pGL_CURR_AMOUNT, object pGL_AMOUNT,
                                object pVENDOR_CODE, object pVENDOR_NAME,
                                object pVAT_TAX_TYPE, object pVAT_TAX_TYPE_NAME)
        {
            InitializeComponent(); 
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_SALE_HEADER_ID.EditValue = pSALE_HEADER_ID;
            V_SLIP_LINE_TYPE.EditValue = pSLIP_LINE_TYPE;
            V_SLIP_DATE.EditValue = pSLIP_DATE; 

            mDEPT_ID = pDEPT_ID;
            mDEPT_NAME = pDEPT_NAME;
            mCURRENCY_CODE = pCURRENCY_CODE;
            mEXCHANGE_RATE = pEXCHANGE_RATE;
            mGL_CURRENCY_AMOUNT = pGL_CURR_AMOUNT;
            mGL_AMOUNT = pGL_AMOUNT;

            mVENDOR_CODE = pVENDOR_CODE;
            mVENDOR_NAME = pVENDOR_NAME;
            mVAT_TAX_TYPE = pVAT_TAX_TYPE;
            mVAT_TAX_TYPE_NAME = pVAT_TAX_TYPE_NAME;
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
            mBase_Currency_Code = iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = idcACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");            
        }

        private void Set_Control_Item_Prompt(DataRowState pRowState)
        {
            try
            {
                //기존 관리항목 타입 저장 - 수정시 기존입력된 값 유지 위해 -- 
                string vMANAGEMENT1_LOOKUP_TYPE = string.Empty;
                string vMANAGEMENT2_LOOKUP_TYPE = string.Empty;
                string vREFER1_LOOKUP_TYPE = string.Empty;
                string vREFER2_LOOKUP_TYPE = string.Empty;
                string vREFER3_LOOKUP_TYPE = string.Empty;
                string vREFER4_LOOKUP_TYPE = string.Empty;
                string vREFER5_LOOKUP_TYPE = string.Empty;
                string vREFER6_LOOKUP_TYPE = string.Empty;
                string vREFER7_LOOKUP_TYPE = string.Empty;
                string vREFER8_LOOKUP_TYPE = string.Empty;
                if (pRowState == DataRowState.Modified)
                {
                    vMANAGEMENT1_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"), "/");
                    vMANAGEMENT2_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"), "/");
                    vREFER1_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"), "/");
                    vREFER2_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"), "/");
                    vREFER3_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"), "/");
                    vREFER4_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"), "/");
                    vREFER5_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"), "/");
                    vREFER6_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"), "/");
                    vREFER7_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"), "/");
                    vREFER8_LOOKUP_TYPE = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"), "/");
                }

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

                if (pRowState == DataRowState.Modified)
                {
                    if (vMANAGEMENT1_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("MANAGEMENT1", null);
                        igrSLIP_LINE.SetCellValue("MANAGEMENT1_DESC", null);
                    }
                    if (vMANAGEMENT2_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("MANAGEMENT2", null);
                        igrSLIP_LINE.SetCellValue("MANAGEMENT2_DESC", null);
                    }
                    if (vREFER1_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER1", null);
                        igrSLIP_LINE.SetCellValue("REFER1_DESC", null);
                    }
                    if (vREFER2_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER2", null);
                        igrSLIP_LINE.SetCellValue("REFER2_DESC", null);
                    }
                    if (vREFER3_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER3", null);
                        igrSLIP_LINE.SetCellValue("REFER3_DESC", null);
                    }
                    if (vREFER4_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER4", null);
                        igrSLIP_LINE.SetCellValue("REFER4_DESC", null);
                    }
                    if (vREFER5_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER5", null);
                        igrSLIP_LINE.SetCellValue("REFER5_DESC", null);
                    }
                    if (vREFER6_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER6", null);
                        igrSLIP_LINE.SetCellValue("REFER6_DESC", null);
                    }
                    if (vREFER7_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER7", null);
                        igrSLIP_LINE.SetCellValue("REFER7_DESC", null);
                    }
                    if (vREFER8_LOOKUP_TYPE != iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")))
                    {
                        igrSLIP_LINE.SetCellValue("REFER8", null);
                        igrSLIP_LINE.SetCellValue("REFER8_DESC", null);
                    }
                }
                else
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
            }
            catch
            {
                //
            }
        }

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

        private void SetManagementParameter(string pManagement_Field, string pEnabled_YN, object pLookup_Type)
        {
            string mLookup_Type = iString.ISNull(pLookup_Type);
            
            if (mLookup_Type == "VAT_TAX_TYPE")
            {//세무구분
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", ACCOUNT_CODE.EditValue);
            }
            else if (mLookup_Type == "VAT_REASON")
            {//부가세사유
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", Get_Management_Value("VAT_TAX_TYPE"));
            }
            else if (mLookup_Type == "DEPT".ToString())
            {
                
            }
            else if (mLookup_Type == "COSTCENTER".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", Get_Management_Value("DEPT"));
            }
            else if (mLookup_Type == "BANK_ACCOUNT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", Get_Management_Value("BANK_SITE"));
            }
            else if (mLookup_Type == "BANK_SITE".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", Get_Management_Value("CUSTOMER"));
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
                string vGL_DATE = null;
                if (iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = V_SLIP_DATE.DateTimeValue.ToShortDateString();
                }
                else if (iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = V_SLIP_DATE.DateTimeValue.ToShortDateString();
                }
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", vGL_DATE);
            }
            else
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", null);
            }
            ildMANAGEMENT.SetLookupParamValue("W_MANAGEMENT_FIELD", pManagement_Field);
            ildMANAGEMENT.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
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
            string vLookup_Type = Get_Lookup_Type(pManagement);
            if (vLookup_Type != pLookup_Type)
            {
                return;
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
            int mPreviousRowPosition = IDA_ASSET_SALE_SLIP.CurrentRowPosition() - 1;
            object mPrevious_Code;
            object mPrevious_Name;
            string mData_Type;
            string mLookup_Type;

            if (mPreviousRowPosition > -1
                && iString.ISNull(REMARK.EditValue) == string.Empty
                && iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REMARK"]) != string.Empty)
            {//REMARK.
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REMARK"];
                REMARK.EditValue = mPrevious_Name;
            }
            
            //1
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT1.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT1.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_LOOKUP_TYPE"]))
            {//MANAGEMENT1_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["MANAGEMENT1"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_DESC"];

                MANAGEMENT1.EditValue = mPrevious_Code;
                MANAGEMENT1_DESC.EditValue = mPrevious_Name;
            }
            //2
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT2.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT2.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_LOOKUP_TYPE"]))
            {//MANAGEMENT2_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["MANAGEMENT2"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_DESC"];

                MANAGEMENT2.EditValue = mPrevious_Code;
                MANAGEMENT2_DESC.EditValue = mPrevious_Name;
            }
            //3
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER1_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER1_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER1.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER1.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER1_LOOKUP_TYPE"]))
            {//REFER1_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER1"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER1_DESC"];

                REFER1.EditValue = mPrevious_Code;
                REFER1_DESC.EditValue = mPrevious_Name;
            }
            //4
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER2_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER2_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER2.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER2.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER2_LOOKUP_TYPE"]))
            {//REFER2_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER2"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER2_DESC"];

                REFER2.EditValue = mPrevious_Code;
                REFER2_DESC.EditValue = mPrevious_Name;
            }
            //5
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER3_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER3_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER3.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER3.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER3_LOOKUP_TYPE"]))
            {//REFER3_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER3"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER3_DESC"];

                REFER3.EditValue = mPrevious_Code;
                REFER3_DESC.EditValue = mPrevious_Name;
            }
            //6
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER4_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER4_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER4.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER4.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER4_LOOKUP_TYPE"]))
            {//REFER4_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER4"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER4_DESC"];

                REFER4.EditValue = mPrevious_Code;
                REFER4_DESC.EditValue = mPrevious_Name;
            }
            //7
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER5_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER5_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER5.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER5.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER5_LOOKUP_TYPE"]))
            {//REFER5_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER5"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER5_DESC"];

                REFER5.EditValue = mPrevious_Code;
                REFER5_DESC.EditValue = mPrevious_Name;
            }
            //8
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER6_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER6_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER6.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER6.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER6_LOOKUP_TYPE"]))
            {//REFER6_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER6"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER6_DESC"];

                REFER6.EditValue = mPrevious_Code;
                REFER6_DESC.EditValue = mPrevious_Name;
            }
            //9
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER7_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER7_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER7.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER7.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER7_LOOKUP_TYPE"]))
            {//REFER7_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER7"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER7_DESC"];

                REFER7.EditValue = mPrevious_Code;
                REFER7_DESC.EditValue = mPrevious_Name;
            }
            //10
            mData_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER8_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRow["REFER8_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER8.EditValue) == string.Empty && iString.ISNull(V_SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER8.EditValue = iDate.ISGetDate(V_SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER8_LOOKUP_TYPE"]))
            {//REFER8_LOOKUP_TYPE
                mPrevious_Code = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER8"];
                mPrevious_Name = IDA_ASSET_SALE_SLIP.CurrentRows[mPreviousRowPosition]["REFER8_DESC"];

                REFER8.EditValue = mPrevious_Code;
                REFER8_DESC.EditValue = mPrevious_Name;
            }
        }

        private void Init_Currency_Code(string pInit_YN)
        {
            
        } 

        //관리항목 LOOKUP 선택시 처리.
        private void Init_SELECT_LOOKUP(object pManagement_Type)
        {
            string mMANAGEMENT = iString.ISNull(pManagement_Type);
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
         
        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                     
                }
                //전표 행 위치 보정 위해 주석 
                //else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                //{
                //    if (idaSLIP_LINE.IsFocused)
                //    {
                //        if (Check_Sub_Panel() == false)
                //        {
                //            return;
                //        } 

                //        idaSLIP_LINE.AddOver();
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
                //            idaSLIP_HEADER.SetSelectParamValue("W_SLIP_HEADER_ID", 0);
                //            idaSLIP_HEADER.Fill();

                //            if (Check_Sub_Panel() == false)
                //            {
                //                return;
                //            } 

                //            idaSLIP_HEADER.AddOver();
                //            idaSLIP_LINE.AddOver();
                //            InsertSlipHeader();
                //            InsertSlipLine();

                //            SLIP_DATE.Focus();
                //        }
                //    }
                //}
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    ;
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                     
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0316_ACCOUNT_Load(object sender, EventArgs e)
        {             
            // 회계장부 정보 설정.
            GetAccountBook();
            IDA_ASSET_SALE_SLIP.FillSchema();
        }

        private void FCMF0316_ACCOUNT_Shown(object sender, EventArgs e)
        {
            igrSLIP_LINE.LastConfirmChanges();
            IDA_ASSET_SALE_SLIP.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_SLIP.Refillable = true;


            string vADD_FLAG = "N";
            IDA_ASSET_SALE_SLIP.SetSelectParamValue("P_SALE_HEADER_ID", V_SALE_HEADER_ID.EditValue);
            IDA_ASSET_SALE_SLIP.SetSelectParamValue("P_SLIP_LINE_TYPE", V_SLIP_LINE_TYPE.EditValue);
            IDA_ASSET_SALE_SLIP.Fill();
            if (IDA_ASSET_SALE_SLIP.CurrentRows.Count < 1)
            {
                IDA_ASSET_SALE_SLIP.AddUnder();
                vADD_FLAG = "Y";

                BUDGET_DEPT_ID_L.EditValue = mDEPT_ID;
                BUDGET_DEPT_NAME_L.EditValue = mDEPT_NAME;
                CURRENCY_CODE.EditValue = mCURRENCY_CODE;
                CURRENCY_DESC.EditValue = mCURRENCY_CODE;
                EXCHANGE_RATE.EditValue = mEXCHANGE_RATE;
                GL_CURRENCY_AMOUNT.EditValue = mGL_CURRENCY_AMOUNT;
                GL_AMOUNT.EditValue = mGL_AMOUNT;
                if (iString.ISNull(V_SLIP_LINE_TYPE.EditValue) == "VAT")
                {
                    //부가세 계정 세팅
                    IDC_GET_VAT_ACCOUNT_CODE.SetCommandParamValue("W_VAT_TAX_TYPE", mVAT_TAX_TYPE);
                    IDC_GET_VAT_ACCOUNT_CODE.ExecuteNonQuery();

                    ACCOUNT_CONTROL_ID.EditValue = IDC_GET_VAT_ACCOUNT_CODE.GetCommandParamValue("O_ACCOUNT_CONTROL_ID");
                    ACCOUNT_CODE.EditValue = IDC_GET_VAT_ACCOUNT_CODE.GetCommandParamValue("O_ACCOUNT_CODE");
                    ACCOUNT_DESC.EditValue = IDC_GET_VAT_ACCOUNT_CODE.GetCommandParamValue("O_ACCOUNT_DESC");
                    ACCOUNT_DR_CR.EditValue = IDC_GET_VAT_ACCOUNT_CODE.GetCommandParamValue("O_ACCOUNT_DR_CR");
                    ACCOUNT_DR_CR_NAME.EditValue = IDC_GET_VAT_ACCOUNT_CODE.GetCommandParamValue("O_ACCOUNT_DR_CR_NAME");

                    Set_Control_Item_Prompt(IDA_ASSET_SALE_SLIP.CurrentRow.RowState);

                    GL_CURRENCY_AMOUNT.EditValue = 0;  //외화금액 0원처리.
                }
            }
            if (iString.ISNull(V_SLIP_LINE_TYPE.EditValue) == "VAT")
            {
                //부가세 계정 세팅 
                ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", ACCOUNT_CODE.EditValue);
            }
            else
            {
                ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", null);
            }
            
             
            Init_Set_Item_Prompt(IDA_ASSET_SALE_SLIP.CurrentRow);
            Init_Set_Item_Need(IDA_ASSET_SALE_SLIP.CurrentRow);

            //기본값 설정//
            if (vADD_FLAG == "Y")
            {
                Application.DoEvents();

                //고객사//
                Set_Management_Value("CUSTOMER", mVENDOR_CODE, mVENDOR_NAME);

                //부가세유형//
                Set_Management_Value("VAT_TAX_TYPE", mVAT_TAX_TYPE, mVAT_TAX_TYPE_NAME);

                //부가세유형//
                Set_Management_Value("VAT_ISSUE_DATE", V_SLIP_DATE.EditValue, null);

                //공급가액 동기화//
                Set_Ref_Management_Value("MANAGEMENT1", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", MANAGEMENT1.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("MANAGEMENT2", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", MANAGEMENT2.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER1", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER1.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER2", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER2.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER3", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER3.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER4", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER4.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER5", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER5.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER6", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER6.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER7", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER7.EditValue, null, null, GL_AMOUNT.EditValue);
                Set_Ref_Management_Value("REFER8", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER8.EditValue, null, null, GL_AMOUNT.EditValue);

                //만기일자 동기화//
                Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "DUE_DATE", MANAGEMENT1.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("MANAGEMENT2", "CUSTOMER", "DUE_DATE", MANAGEMENT2.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER1", "CUSTOMER", "DUE_DATE", REFER1.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER2", "CUSTOMER", "DUE_DATE", REFER2.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER3", "CUSTOMER", "DUE_DATE", REFER3.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER4", "CUSTOMER", "DUE_DATE", REFER4.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER5", "CUSTOMER", "DUE_DATE", REFER5.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER6", "CUSTOMER", "DUE_DATE", REFER6.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER7", "CUSTOMER", "DUE_DATE", REFER7.EditValue, null, V_SLIP_DATE.EditValue, null);
                Set_Ref_Management_Value("REFER8", "CUSTOMER", "DUE_DATE", REFER8.EditValue, null, V_SLIP_DATE.EditValue, null);
            }
            BUDGET_DEPT_NAME_L.Focus();
        }
                
        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_SALE_SLIP.Cancel();
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void BTN_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_SALE_SLIP.Delete();
            IDA_ASSET_SALE_SLIP.Update();            
        }

        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_SALE_SLIP.Update();
        }


        private void MANAGEMENT1_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("MANAGEMENT1", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void MANAGEMENT2_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("MANAGEMENT2", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER1_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER1", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER2_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER2", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER3_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER3", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER4_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER4", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER5_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER5", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER6_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER6", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER7_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER7", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        private void REFER8_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //부가세 세금유형을 선택하면 부가세이유를 CLEAR 
            Set_Validate_Management_Value("REFER8", "VAT_TAX_TYPE", "VAT_REASON", null, null);
        }

        #endregion

        #region ----- Lookup Event ----- 
        
        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", " VALUE1 <> 'BL'", "Y");
        }
         
        private void ilaSLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", " VALUE1 <> 'BL'", "Y");
        }

        private void ilaREQ_PAYABLE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("PAYABLE_TYPE", "Y");
        }
           
        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaBUDGET_DEPT_L_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBUDGET_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildBUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", V_SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", V_SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "A");
        }

        private void ilaACCOUNT_DR_CR_SelectedRowData(object pSender)
        {
            //전호수주석 : 관리항목 변경.
            //Set_Control_Item_Prompt();
            //Init_Control_Management_Value();
            //Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);
            //Init_Set_Item_Need(idaSLIP_LINE.CurrentRow);
            //Init_Default_Value(); 
        }

        private void ilaACCOUNT_CONTROL_SelectedRowData(object pSender)
        {
            Set_Control_Item_Prompt(IDA_ASSET_SALE_SLIP.CurrentRow.RowState);
            Init_Set_Item_Prompt(IDA_ASSET_SALE_SLIP.CurrentRow);
            Init_Set_Item_Need(IDA_ASSET_SALE_SLIP.CurrentRow);
            if (iString.ISNull(V_SLIP_LINE_TYPE.EditValue) == "AR")
            {
                //고객사//
                Set_Management_Value("CUSTOMER", mVENDOR_CODE, mVENDOR_NAME);

                //1.거래처 선택시 만기일자 있으면 만기일자 설정//
                Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "DUE_DATE", MANAGEMENT1.EditValue);
                Set_Ref_Management_Value("MANAGEMENT2", "CUSTOMER", "DUE_DATE", MANAGEMENT2.EditValue);
                Set_Ref_Management_Value("REFER1", "CUSTOMER", "DUE_DATE", REFER1.EditValue);
                Set_Ref_Management_Value("REFER2", "CUSTOMER", "DUE_DATE", REFER2.EditValue);
                Set_Ref_Management_Value("REFER3", "CUSTOMER", "DUE_DATE", REFER3.EditValue);
                Set_Ref_Management_Value("REFER4", "CUSTOMER", "DUE_DATE", REFER4.EditValue);
                Set_Ref_Management_Value("REFER5", "CUSTOMER", "DUE_DATE", REFER5.EditValue);
                Set_Ref_Management_Value("REFER6", "CUSTOMER", "DUE_DATE", REFER6.EditValue);
                Set_Ref_Management_Value("REFER7", "CUSTOMER", "DUE_DATE", REFER7.EditValue);
                Set_Ref_Management_Value("REFER8", "CUSTOMER", "DUE_DATE", REFER8.EditValue);  
            } 
        }
         
        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_SLIP_REMARK_SelectedRowData(object pSender)
        {
            if (iString.ISNull(REMARK.EditValue) != string.Empty)
            {
                REMARK.TextSelectionStart = iString.ISNull(REMARK.EditValue).Length;
                REMARK.Focus();
            }
        }

        private void ilaMANAGEMENT1_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("MANAGEMENT1_ID", "Y", igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("MANAGEMENT2_ID", "Y", igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
        }

        private void ilaREFER1_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER1_ID", "Y", igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
        }

        private void ilaREFER2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER2_ID", "Y", igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
        }

        private void ilaREFER3_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER3_ID", "Y", igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
        }

        private void ilaREFER4_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER4_ID", "Y", igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
        }

        private void ilaREFER5_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER5_ID", "Y", igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
        }

        private void ilaREFER6_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER6_ID", "Y", igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
        }

        private void ilaREFER7_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER7_ID", "Y", igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
        }

        private void ilaREFER8_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetManagementParameter("REFER8_ID", "Y", igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT1_SelectedRowData(object pSender)
        {// 관리항목1 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT1");

            //관리항목 동기화// 
            //1.거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "DUE_DATE", MANAGEMENT1.EditValue);
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
            //신용카드 결재일자//
            Set_Ref_Management_Value("REFER8", "CREDIT_CARD", "DUE_DATE", REFER8.EditValue);
            //공급가액 동기화//
            Set_Ref_Management_Value("REFER8", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER8.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaVAT_ASSET_GB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_ASSET_GB", "Y");
        }

        #endregion       

        #region ----- Adapter Event -----
         
        private void idaSLIP_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {// 예산부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_NAME_L))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
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
                if (mBase_Currency_Code.ToString() != e.Row["CURRENCY_CODE"].ToString() && iString.ISDecimaltoZero(e.Row["EXCHANGE_RATE"]) == Convert.ToInt32(0))
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
             
        }

        private void IDA_ASSET_SALE_SLIP_UpdateCompleted(object pSender)
        {
            if (IDA_ASSET_SALE_SLIP.UpdateChangedRowCount != 0 || IDA_ASSET_SALE_SLIP.ChangedRowCount == 0)
            {
                DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        #endregion


    }
}