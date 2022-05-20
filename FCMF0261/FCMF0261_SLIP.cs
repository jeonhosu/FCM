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

namespace FCMF0261
{
    public partial class FCMF0261_SLIP : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
          
        string mStatus = "";

        object mSys_Date;
        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mCurrency_Code;
        object mBudget_Control_YN;
        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ---- Public Value ----

        public object Get_Slip_Num
        {
            get
            {
                return SLIP_NUM.EditValue;
            }
        }
        
        public object Get_Slip_Date
        {
            get
            {
                return SLIP_DATE.EditValue;
            }
        }

        public object Get_Slip_Header_ID
        {
            get
            {
                return SLIP_INTERFACE_HEADER_ID.EditValue;
            }

        }

        #endregion

        #region ----- Constructor -----

        public FCMF0261_SLIP()
        {
            InitializeComponent();
        }

        public FCMF0261_SLIP(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public FCMF0261_SLIP(Form pMainForm, ISAppInterface pAppInterface, string pStatus, Object pSESSION_ID, Object pSYS_DATE)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mStatus = pStatus;
            mSession_ID = pSESSION_ID;
            mSys_Date = pSYS_DATE;
        }

        #endregion;

        #region ----- Private Methods -----

        private void GetAccountBook()
        {
            idcACCOUNT_BOOK.ExecuteNonQuery(); 
            mAccount_Book_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = idcACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = idcACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = iConv.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = idcACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
            if (iConv.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_TEMP_SLIP_REMARK_FLAG")) == "Y")
            {
                REMARK.LookupAdapter = ILA_TEMP_SLIP_REMARK;
            }
            else
            {
                REMARK.LookupAdapter = null;
            }
        }

        private void Search_DB()
        { 
            try
            { 
                IDA_CARD_SLIP_GROUP_APPR.SetSelectParamValue("W_SYS_DATE", mSys_Date);
                IDA_CARD_SLIP_GROUP_APPR.SetSelectParamValue("W_SESSION_ID", mSession_ID);
                IDA_CARD_SLIP_GROUP_APPR.SetSelectParamValue("W_SLIP_FLAG", mStatus);
                IDA_CARD_SLIP_GROUP_APPR.Fill();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }

            //헤더 적용등 조회//
            IDC_GET_CARD_SLIP_P.SetCommandParamValue("P_SYS_DATE", mSys_Date);
            IDC_GET_CARD_SLIP_P.ExecuteNonQuery();
            HEADER_REMARK.EditValue = IDC_GET_CARD_SLIP_P.GetCommandParamValue("O_HEADER_REMARK");

            IGR_CARD_SLIP_GROUP_APPR.Focus();
        }

        private void Set_Control_Item_Prompt(DataRowState pRowState)
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
                vMANAGEMENT1_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"), "/");
                vMANAGEMENT2_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"), "/");
                vREFER1_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"), "/");
                vREFER2_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"), "/");
                vREFER3_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"), "/");
                vREFER4_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"), "/");
                vREFER5_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"), "/");
                vREFER6_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"), "/");
                vREFER7_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"), "/");
                vREFER8_LOOKUP_TYPE = iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"), "/");
            }

            idaCONTROL_ITEM_PROMPT.Fill();
            if (idaCONTROL_ITEM_PROMPT.CurrentRows.Count > 0)
            {
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_NAME"]);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER1_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER2_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER3_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER4_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER5_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER6_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER7_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_NAME"]);
                IGR_SLIP_LINE.SetCellValue("REFER8_NAME", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_NAME"]);

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_YN"]);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER1_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER2_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER3_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER4_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER5_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER6_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER7_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER8_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_YN"]);

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER1_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER2_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER3_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER4_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER5_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER6_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER7_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_YN"]);
                IGR_SLIP_LINE.SetCellValue("REFER8_LOOKUP_YN", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_YN"]);

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER1_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER2_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER3_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER4_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER5_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER6_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER7_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_LOOKUP_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER8_LOOKUP_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_LOOKUP_TYPE"]);

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER1_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER2_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER3_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER4_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER5_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER6_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER7_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_DATA_TYPE"]);
                IGR_SLIP_LINE.SetCellValue("REFER8_DATA_TYPE", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_DATA_TYPE"]);
            }
            else
            {
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_NAME", null);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER1_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER2_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER3_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER4_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER5_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER6_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER7_NAME", null);
                IGR_SLIP_LINE.SetCellValue("REFER8_NAME", null);

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_YN", "F");
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER1_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER2_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER3_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER4_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER5_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER6_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER7_YN", "F");
                IGR_SLIP_LINE.SetCellValue("REFER8_YN", "F");

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER1_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER2_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER3_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER4_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER5_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER6_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER7_LOOKUP_YN", "N");
                IGR_SLIP_LINE.SetCellValue("REFER8_LOOKUP_YN", "N");

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER1_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER2_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER3_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER4_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER5_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER6_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER7_LOOKUP_TYPE", null);
                IGR_SLIP_LINE.SetCellValue("REFER8_LOOKUP_TYPE", null);

                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER1_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER2_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER3_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER4_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER5_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER6_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER7_DATA_TYPE", "VARCHAR2");
                IGR_SLIP_LINE.SetCellValue("REFER8_DATA_TYPE", "VARCHAR2");
            }

            if (pRowState == DataRowState.Modified)
            {
                if (vMANAGEMENT1_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", null);
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_DESC", null);
                }
                if (vMANAGEMENT2_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", null);
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_DESC", null);
                }
                if (vREFER1_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER1", null);
                    IGR_SLIP_LINE.SetCellValue("REFER1_DESC", null);
                }
                if (vREFER2_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER2", null);
                    IGR_SLIP_LINE.SetCellValue("REFER2_DESC", null);
                }
                if (vREFER3_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER3", null);
                    IGR_SLIP_LINE.SetCellValue("REFER3_DESC", null);
                }
                if (vREFER4_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER4", null);
                    IGR_SLIP_LINE.SetCellValue("REFER4_DESC", null);
                }
                if (vREFER5_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER5", null);
                    IGR_SLIP_LINE.SetCellValue("REFER5_DESC", null);
                }
                if (vREFER6_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER6", null);
                    IGR_SLIP_LINE.SetCellValue("REFER6_DESC", null);
                }
                if (vREFER7_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER7", null);
                    IGR_SLIP_LINE.SetCellValue("REFER7_DESC", null);
                }
                if (vREFER8_LOOKUP_TYPE != iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")))
                {
                    IGR_SLIP_LINE.SetCellValue("REFER8", null);
                    IGR_SLIP_LINE.SetCellValue("REFER8_DESC", null);
                }
            }
            else
            {
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", null);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_DESC", null);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", null);
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER1", null);
                IGR_SLIP_LINE.SetCellValue("REFER1_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER2", null);
                IGR_SLIP_LINE.SetCellValue("REFER2_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER3", null);
                IGR_SLIP_LINE.SetCellValue("REFER3_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER4", null);
                IGR_SLIP_LINE.SetCellValue("REFER4_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER5", null);
                IGR_SLIP_LINE.SetCellValue("REFER5_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER6", null);
                IGR_SLIP_LINE.SetCellValue("REFER6_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER7", null);
                IGR_SLIP_LINE.SetCellValue("REFER7_DESC", null);
                IGR_SLIP_LINE.SetCellValue("REFER8", null);
                IGR_SLIP_LINE.SetCellValue("REFER8_DESC", null);
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
            string mLookup_Type = iConv.ISNull(pLookup_Type);

            if (mLookup_Type == "VAT_TAX_TYPE")
            {//세무구분
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", ACCOUNT_CODE.EditValue);
            }
            else if (mLookup_Type == "VAT_REASON")
            {//부가세사유
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("VAT_TAX_TYPE"));
            }
            else if (mLookup_Type == "DEPT".ToString())
            {
                object vDEPT_CODE = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("BUDGET_DEPT_CODE");
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", vDEPT_CODE);
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
                if (iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    vSLIP_DATE = SLIP_DATE.DateTimeValue.ToShortDateString();
                }
                else if (iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    vSLIP_DATE = SLIP_DATE.DateTimeValue.ToShortDateString();
                }
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", vSLIP_DATE);
            }
            else
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", null);
            }
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
            if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT1.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT2.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER1_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER1_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER1.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER2_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER2_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER2.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER3_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER3_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER3.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER4_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER4_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER4.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER5_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER5_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER5.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER6_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER6_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER6.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER7_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER7_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER7.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER8_LOOKUP_TYPE"]) != string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER8_LOOKUP_TYPE"]) == iConv.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER8.EditValue;
            }
            else
            {
                mLookup_Value = null;
            }
            return mLookup_Value;
        }

        //private void GetSlipNum()
        //{
        //    if (iString.ISNull(DOCUMENT_TYPE.EditValue) == string.Empty)
        //    {
        //        return;
        //    }
        //    idcSLIP_NUM.SetCommandParamValue("W_DOCUMENT_TYPE", DOCUMENT_TYPE.EditValue);
        //    idcSLIP_NUM.ExecuteNonQuery();
        //    SLIP_NUM.EditValue = idcSLIP_NUM.GetCommandParamValue("O_DOCUMENT_NUM");
        //    GL_NUM.EditValue = SLIP_NUM.EditValue;
        //}

        private void GetSubForm()
        {
            ibtSUB_FORM.Visible = false;
            ACCOUNT_CLASS_YN.EditValue = null;
            ACCOUNT_CLASS_TYPE.EditValue = null;
            string vBTN_CAPTION = null;

            if (iConv.ISNull(ACCOUNT_CONTROL_ID.EditValue) == string.Empty || iConv.ISNull(ACCOUNT_DR_CR.EditValue) == string.Empty)   
            {
                return;
            }
            idcGET_SUB_FORM.ExecuteNonQuery();
            ACCOUNT_CLASS_YN.EditValue = idcGET_SUB_FORM.GetCommandParamValue("O_ACCOUNT_CLASS_YN");
            ACCOUNT_CLASS_TYPE.EditValue = idcGET_SUB_FORM.GetCommandParamValue("O_ACCOUNT_CLASS_TYPE");
            vBTN_CAPTION = iConv.ISNull(idcGET_SUB_FORM.GetCommandParamValue("O_BTN_CAPTION"));
            if (iConv.ISNull(ACCOUNT_CLASS_YN.EditValue, "N") == "N".ToString())
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
            if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                MANAGEMENT1.EditValue = pManagement_Value;
                MANAGEMENT1_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                MANAGEMENT2.EditValue = pManagement_Value;
                MANAGEMENT2_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                REFER1.EditValue = pManagement_Value;
                REFER1_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                REFER2.EditValue = pManagement_Value;
                REFER2_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                REFER3.EditValue = pManagement_Value;
                REFER3_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                REFER4.EditValue = pManagement_Value;
                REFER4_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                REFER5.EditValue = pManagement_Value;
                REFER5_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                REFER6.EditValue = pManagement_Value;
                REFER6_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                REFER7.EditValue = pManagement_Value;
                REFER7_DESC.EditValue = pManagement_Desc;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                REFER8.EditValue = pManagement_Value;
                REFER8_DESC.EditValue = pManagement_Desc;
            }
        }

        private void Set_Ref_Management_Value(string pManagement, string pLookup_Type, string pRef_Lookup_Type, object pManagement_Value, object pManagement_Desc)
        {
            if (pManagement == "MANAGEMENT1" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "MANAGEMENT2" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER1" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER2" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER3" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER4" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER5" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER6" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER7" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER8" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
        }

        private object Get_Management_Value(string pLookup_Type)
        {
            object vManagement_Value = null;
            if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                vManagement_Value = MANAGEMENT1.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                vManagement_Value = MANAGEMENT2.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                vManagement_Value = REFER1.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                vManagement_Value = REFER2.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                vManagement_Value = REFER3.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                vManagement_Value = REFER4.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                vManagement_Value = REFER5.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                vManagement_Value = REFER6.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                vManagement_Value = REFER7.EditValue;
            }
            else if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                vManagement_Value = REFER8.EditValue;
            }
            return vManagement_Value;
        }

        private void Set_Validate_Management_Value(string pManagement, string pLookup_Type, string pRef_Lookup_Type, object pManagement_Value, object pManagement_Desc)
        {
            if (pManagement == "MANAGEMENT1" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목1
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "MANAGEMENT2" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목2
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER1" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목3
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER2" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목4
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER3" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목5
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER4" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목6
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER5" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목7
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER6" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목8
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER7" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목9
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
            else if (pManagement == "REFER8" &&
                iConv.ISNull(IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")).ToUpper() == pLookup_Type.ToUpper())
            {//관리항목10
                Set_Management_Value(pRef_Lookup_Type, pManagement_Value, pManagement_Desc);
            }
        }

        #endregion;

        #region ----- Initialize Event -----
        
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

                igbSLIP_HEADER.Enabled = false;
                igbCONFIRM_INFOMATION.Enabled = false;
                igbACCOUNT_LINE.Enabled = false;
                igbSLIP_LINE.Enabled = false;
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

                igbSLIP_HEADER.Enabled = true;
                igbCONFIRM_INFOMATION.Enabled = true;
                igbACCOUNT_LINE.Enabled = true;
                igbSLIP_LINE.Enabled = true;
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
        
        private Boolean Check_SlipHeader_Added()
        {
            Boolean Row_Added_Status = false;
            //헤더 체크 
            for (int r = 0; r < IDA_CARD_SLIP_GROUP_APPR.SelectRows.Count; r++)
            {
                if (IDA_CARD_SLIP_GROUP_APPR.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_CARD_SLIP_GROUP_APPR.SelectRows[r].RowState == DataRowState.Modified)
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
                for (int r = 0; r < IDA_SLIP_LINE.SelectRows.Count; r++)
                {
                    if (IDA_SLIP_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_SLIP_LINE.SelectRows[r].RowState == DataRowState.Modified)
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

        //private void InsertSlipHeader()
        //{
        //    SLIP_DATE.EditValue = iDate.ISGetDate(mGrid.GetCellValue("SLIP_DATE"));
        //    H_REMARK.EditValue = mGrid.GetCellValue("REMARK");

        //    idcDV_SLIP_TYPE.SetCommandParamValue("W_WHERE", "GROUP_CODE = 'SLIP_TYPE' AND VALUE1 = 'CC' ");
        //    idcDV_SLIP_TYPE.ExecuteNonQuery();
        //    SLIP_TYPE.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_CODE");
        //    SLIP_TYPE_NAME.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_CODE_NAME");
        //    SLIP_TYPE_CLASS.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_VALUE1");
        //    DOCUMENT_TYPE.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_VALUE2");
             
        //    idcUSER_INFO.ExecuteNonQuery();
        //    DEPT_NAME.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_NAME");
        //    DEPT_CODE.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_CODE");
        //    DEPT_ID.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_ID");
        //    PERSON_NAME.EditValue = idcUSER_INFO.GetCommandParamValue("O_PERSON_NAME");
        //    PERSON_ID.EditValue = isAppInterfaceAdv1.PERSON_ID;

        //    //헤더 예산부서
        //    BUDGET_DEPT_NAME.EditValue = DEPT_NAME.EditValue;
        //    BUDGET_DEPT_CODE.EditValue = DEPT_CODE.EditValue;
        //    BUDGET_DEPT_ID.EditValue = DEPT_ID.EditValue;
        //}

        private void InsertSlipLine()
        {
            Set_Slip_Line_Seq();    //LINE SEQ 채번//

            SLIP_LINE_TYPE.EditValue = "EXP";
            CURRENCY_CODE.EditValue = mCurrency_Code;
            CURRENCY_DESC.EditValue = mCurrency_Code;
            Init_Currency_Amount();
            Init_Budget_Dept();
            GL_AMOUNT.EditValue = 0;
            GL_CURRENCY_AMOUNT.EditValue = 0;

            BUDGET_DEPT_NAME_L.Focus(); 
        }
        
        private void Set_Slip_Line_Seq()
        {
            //LINE SEQ 채번//
            decimal mSLIP_LINE_SEQ = 0;
            decimal vPre_Line_Seq = 0;
            decimal vNext_Line_Seq = 0;
            try
            {
                int mPreviousRowPosition = IDA_SLIP_LINE.CurrentRowPosition() - 1;

                //현재 이전 line seq 
                if (mPreviousRowPosition > -1)
                {
                    vPre_Line_Seq = iConv.ISDecimaltoZero(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["SLIP_LINE_SEQ"], 1);
                }
                else
                {
                    vPre_Line_Seq = 0;
                }

                //현재 다음 line seq                 
                if ((IDA_SLIP_LINE.CurrentRowPosition() + 1) == IDA_SLIP_LINE.CurrentRows.Count)
                {
                    vNext_Line_Seq = 0;
                }
                else
                {
                    int mNextRowPosition = IDA_SLIP_LINE.CurrentRowPosition() + 1;
                    vNext_Line_Seq = iConv.ISDecimaltoZero(IDA_SLIP_LINE.CurrentRows[mNextRowPosition]["SLIP_LINE_SEQ"], 1);
                }

                //실재 Slip Line Seq 채번//
                if (vNext_Line_Seq == 0)
                {
                    mSLIP_LINE_SEQ = Math.Truncate(vPre_Line_Seq) + 1;
                }
                else
                {
                    decimal vAvg = Math.Round(((vNext_Line_Seq - vPre_Line_Seq) / 2), 10);
                    mSLIP_LINE_SEQ = vPre_Line_Seq + vAvg;
                }
            }
            catch
            {

            }
            IGR_SLIP_LINE.SetCellValue("SLIP_LINE_SEQ", mSLIP_LINE_SEQ);
        }

        private void Set_Insert_Slip_Line()
        {
            IDA_BALANCE_SLIP_BUDGET.Fill();
            if (IDA_BALANCE_SLIP_BUDGET.SelectRows.Count < 1)
            {
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Not found data, Check data");
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int Row_Count = IGR_SLIP_LINE.RowCount;
            IGR_SLIP_LINE.BeginUpdate();
            try
            {
                for (int i = 0; i < IDA_BALANCE_SLIP_BUDGET.SelectRows.Count; i++)
                {
                    IDA_SLIP_LINE.AddUnder();
                    for (int c = 0; c < IGR_SLIP_LINE.GridAdvExColElement.Count; c++)
                    {
                        if (IGR_SLIP_LINE.GridAdvExColElement[c].DataColumn.ToString() != "HEADER_ID")
                        {
                            IGR_SLIP_LINE.SetCellValue(i + Row_Count, c, IDA_BALANCE_SLIP_BUDGET.OraDataSet().Rows[i][c]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                IGR_SLIP_LINE.EndUpdate();
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            IGR_SLIP_LINE.EndUpdate();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }
         
        private void Init_GL_Amount()
        {
            if (iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue) == 0)
            {
                return;
            }
            else if (iConv.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue) == 0)
            {
                return;
            }
            decimal mGL_AMOUNT = iConv.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue) * iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue);
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
            Init_DR_CR_Amount();    // 차대금액 생성 //
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
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
            decimal vNew_GL_Amount = iConv.ISDecimaltoZero(pNew_Exchange_Rate) * iConv.ISDecimaltoZero(pCurrency_Amount) ;

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
                IDA_SLIP_LINE.AddUnder();
                for (int c = 0; c < IGR_SLIP_LINE.ColCount; c++)
                {
                    IGR_SLIP_LINE.SetCellValue(IGR_SLIP_LINE.RowIndex, c, IGR_SLIP_LINE.GetCellValue(pCurrent_Row_Index, c));
                }
                // 반제 SLIP_LINE_ID.
                //igrSLIP_LINE.SetCellValue(igrSLIP_LINE.RowIndex, igrSLIP_LINE.GetColumnToIndex("UNLIQUIDATE_SLIP_LINE_ID"), );
                ACCOUNT_DR_CR.EditValue = vAccount_DR_CR;
                ACCOUNT_DR_CR_NAME.EditValue = vAccount_DR_CR_Name;
                ACCOUNT_CONTROL_ID.EditValue = vAccount_ID;
                ACCOUNT_CODE.EditValue = vAccount_Code;
                ACCOUNT_DESC.EditValue = vAccount_Desc;
                EXCHANGE_RATE.EditValue = iConv.ISDecimaltoZero(pOld_Exchange_Rate);
                GL_CURRENCY_AMOUNT.EditValue = iConv.ISDecimaltoZero(pCurrency_Amount);
                GL_AMOUNT.EditValue = Math.Abs(iConv.ISDecimaltoZero(vExchange_Profit_Loss_Amount));

                //참고항목 동기화.                
                //Set_Control_Item_Prompt(IDA_BALANCE_SLIP_BUDGET);
                Init_Set_Item_Prompt(IDA_SLIP_LINE.CurrentRow);

                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                mExchange_Profit_Loss = true;
            }
            return mExchange_Profit_Loss;
        }

        private void Init_DR_CR_Amount()
        {
            mStatus = "NON-QUERY";
            isAppInterfaceAdv1.OnAppMessage(string.Empty);

            if (IGR_SLIP_LINE.RowCount < 1)
            {
                return;
            }
            try
            {
                int vIDX_ROW_CURR = IGR_SLIP_LINE.RowIndex;
                if (IDA_SLIP_LINE.CurrentRowPosition() != vIDX_ROW_CURR)
                {
                    return;
                }

                int vIDX_COL_GL_AMOUNT = IGR_SLIP_LINE.GetColumnToIndex("GL_AMOUNT");
                int vIDX_COL_DR = IGR_SLIP_LINE.GetColumnToIndex("DR_AMOUNT");
                int vIDX_COL_CR = IGR_SLIP_LINE.GetColumnToIndex("CR_AMOUNT");

                if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["ACCOUNT_DR_CR"], "1") == "1".ToString())
                {
                    IGR_SLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, IDA_SLIP_LINE.CurrentRow["GL_AMOUNT"]);
                    IGR_SLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, 0);
                }
                else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["ACCOUNT_DR_CR"], "1") == "2".ToString())
                {
                    IGR_SLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, 0);
                    IGR_SLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, IDA_SLIP_LINE.CurrentRow["GL_AMOUNT"]);
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private void Init_Total_GL_Amount()
        {
            mStatus = "NON-QUERY";

            decimal vDR_Amount = Convert.ToDecimal(0);
            decimal vCR_Amount = Convert.ToDecimal(0);
            decimal vCurrency_DR_Amount = Convert.ToInt32(0);

            foreach (DataRow vRow in IDA_SLIP_LINE.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Deleted)
                {
                    if (iConv.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "1".ToString())
                    {
                        vDR_Amount = vDR_Amount + iConv.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                        vCurrency_DR_Amount = vCurrency_DR_Amount + iConv.ISDecimaltoZero(vRow["GL_CURRENCY_AMOUNT"]);
                    }
                    else if (iConv.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "2".ToString())
                    {
                        vCR_Amount = vCR_Amount + iConv.ISDecimaltoZero(vRow["GL_AMOUNT"]); ;
                    }
                }
            }
            if(VAT_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                vDR_Amount = vDR_Amount + 
                            iConv.ISDecimaltoZero(BASE_VAT_AMOUNT.EditValue, 0);
            }

            vCR_Amount = vCR_Amount +
                        iConv.ISDecimaltoZero(BASE_TOTAL_AMOUNT.EditValue ,0);

            TOTAL_DR_AMOUNT.EditValue = iConv.ISDecimaltoZero(vDR_Amount);
            TOTAL_CR_AMOUNT.EditValue = iConv.ISDecimaltoZero(vCR_Amount);
            MARGIN_AMOUNT.EditValue = -(System.Math.Abs(iConv.ISDecimaltoZero(vDR_Amount) - iConv.ISDecimaltoZero(vCR_Amount))); ;
        }

        private void Init_Control_Management_Value()
        {
            IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", null);
            IGR_SLIP_LINE.SetCellValue("MANAGEMENT1_DESC", null);
            IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", null);
            IGR_SLIP_LINE.SetCellValue("MANAGEMENT2_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER1", null);
            IGR_SLIP_LINE.SetCellValue("REFER1_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER2", null);
            IGR_SLIP_LINE.SetCellValue("REFER2_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER3", null);
            IGR_SLIP_LINE.SetCellValue("REFER3_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER4", null);
            IGR_SLIP_LINE.SetCellValue("REFER4_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER5", null);
            IGR_SLIP_LINE.SetCellValue("REFER5_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER6", null);
            IGR_SLIP_LINE.SetCellValue("REFER6_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER7", null);
            IGR_SLIP_LINE.SetCellValue("REFER7_DESC", null);
            IGR_SLIP_LINE.SetCellValue("REFER8", null);
            IGR_SLIP_LINE.SetCellValue("REFER8_DESC", null);
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
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) == "R" || iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) == "S")
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
                if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) == "S")
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
            if (iConv.ISNull(pDataRow["CURRENCY_ENABLED_FLAG"], "N") == "Y".ToString())
            {
                CURRENCY_DESC.Nullable = false;
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            string mDATA_TYPE = "VARCHAR2";
            object mValue;
            mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
            MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("MANAGEMENT1");
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT1.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("MANAGEMENT1");
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT1.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("MANAGEMENT1");
                    MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    MANAGEMENT1.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                MANAGEMENT1.ReadOnly = true;
                MANAGEMENT1.Insertable = false;
                MANAGEMENT1.Updatable = false;
                MANAGEMENT1.TabStop = false;
            }
            MANAGEMENT1.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
            MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("MANAGEMENT2");
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT2.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("MANAGEMENT2");
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    MANAGEMENT2.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("MANAGEMENT2");
                    MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    MANAGEMENT2.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                MANAGEMENT2.ReadOnly = true;
                MANAGEMENT2.Insertable = false;
                MANAGEMENT2.Updatable = false;
                MANAGEMENT2.TabStop = false;
            }
            MANAGEMENT2.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER1_DATA_TYPE"]);
            REFER1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER1");
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER1.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER1", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER1");
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER1.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER1", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER1");
                    REFER1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER1.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER1", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER1.ReadOnly = true;
                REFER1.Insertable = false;
                REFER1.Updatable = false;
                REFER1.TabStop = false;
            }
            REFER1.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER2_DATA_TYPE"]);
            REFER2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER2");
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER2.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER2", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER2");
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER2.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER2", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER2");
                    REFER2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER2.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER2", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER2.ReadOnly = true;
                REFER2.Insertable = false;
                REFER2.Updatable = false;
                REFER2.TabStop = false;
            }
            REFER2.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER3_DATA_TYPE"]);
            REFER3.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER3");
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER3.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER3", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER3");
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER3.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER3", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER3");
                    REFER3.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER3.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER3", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER3.ReadOnly = true;
                REFER3.Insertable = false;
                REFER3.Updatable = false;
                REFER3.TabStop = false;
            }
            REFER3.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER4_DATA_TYPE"]);
            REFER4.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER4");
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER4.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER4", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER4");
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER4.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER4", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER4");
                    REFER4.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER4.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER4", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER4.ReadOnly = true;
                REFER4.Insertable = false;
                REFER4.Updatable = false;
                REFER4.TabStop = false;
            }
            REFER4.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER5_DATA_TYPE"]);
            REFER5.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER5");
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER5.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER5", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER5");
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER5.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER5", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER5");
                    REFER5.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER5.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER5", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER5.ReadOnly = true;
                REFER5.Insertable = false;
                REFER5.Updatable = false;
                REFER5.TabStop = false;
            }
            REFER5.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER6_DATA_TYPE"]);
            REFER6.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER6");
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER6.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER6", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER6");
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER6.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER6", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER6");
                    REFER6.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER6.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER6", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER6.ReadOnly = true;
                REFER6.Insertable = false;
                REFER6.Updatable = false;
                REFER6.TabStop = false;
            }
            REFER6.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER7_DATA_TYPE"]);
            REFER7.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER7");
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER7.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER7", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER7");
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER7.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER7", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER7");
                    REFER7.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER7.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER7", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER7.ReadOnly = true;
                REFER7.Insertable = false;
                REFER7.Updatable = false;
                REFER7.TabStop = false;
            }
            REFER7.Refresh();

            mDATA_TYPE = iConv.ISNull(pDataRow["REFER8_DATA_TYPE"]);
            REFER8.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
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
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER8");
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER8.NumberDecimalDigits = 0;
                    IGR_SLIP_LINE.SetCellValue("REFER8", mValue);
                }
                else if (mDATA_TYPE == "RATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER8");
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                    REFER8.NumberDecimalDigits = 4;
                    IGR_SLIP_LINE.SetCellValue("REFER8", mValue);
                }
                else if (mDATA_TYPE == "DATE".ToString())
                {
                    mValue = IGR_SLIP_LINE.GetCellValue("REFER8");
                    REFER8.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
                    REFER8.DateFormat = "yyyy-MM-dd";
                    IGR_SLIP_LINE.SetCellValue("REFER8", mValue);
                }
            }
            if (iConv.ISNull(pDataRow["REF_SLIP_FLAG"]) != string.Empty)
            {
                REFER8.ReadOnly = true;
                REFER8.Insertable = false;
                REFER8.Updatable = false;
                REFER8.TabStop = false;
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
            if (MANAGEMENT1.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("MANAGEMENT1");
                MANAGEMENT1.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["MANAGEMENT1_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    MANAGEMENT1.ReadOnly = true;
                    MANAGEMENT1.Nullable = false;
                    MANAGEMENT1.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT1", mDATA_VALUE);
                MANAGEMENT1.Refresh();
            }

            //--2
            if (MANAGEMENT2.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("MANAGEMENT2");
                MANAGEMENT2.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["MANAGEMENT2_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    MANAGEMENT2.ReadOnly = true;
                    MANAGEMENT2.Nullable = false;
                    MANAGEMENT2.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("MANAGEMENT2", mDATA_VALUE);
                MANAGEMENT2.Refresh();
            }

            //--3
            if (REFER1.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER1");
                REFER1.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER1_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER1_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER1.ReadOnly = true;
                    REFER1.Nullable = false;
                    REFER1.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER1", mDATA_VALUE);
                REFER1.Refresh();
            }

            //--4
            if (REFER2.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER2");
                REFER2.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER2_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER2_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER2.ReadOnly = true;
                    REFER2.Nullable = false;
                    REFER2.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER2", mDATA_VALUE);
                REFER2.Refresh();
            }

            //--5
            if (REFER3.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER3");
                REFER3.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER3_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER3_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER3.ReadOnly = true;
                    REFER3.Nullable = false;
                    REFER3.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER3", mDATA_VALUE);
                REFER3.Refresh();
            }

            //--6
            if (REFER4.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER4");
                REFER4.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER4_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER4_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER4.ReadOnly = true;
                    REFER4.Nullable = false;
                    REFER4.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER4", mDATA_VALUE);
                REFER4.Refresh();
            }

            //--7
            if (REFER5.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER5");
                REFER5.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER5_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER5_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER5.ReadOnly = true;
                    REFER5.Nullable = false;
                    REFER5.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER5", mDATA_VALUE);
                REFER5.Refresh();
            }

            //--8
            if (REFER6.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER6");
                REFER6.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER6_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER6_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER6.ReadOnly = true;
                    REFER6.Nullable = false;
                    REFER6.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER6", mDATA_VALUE);
                REFER6.Refresh();
            }

            //--9
            if (REFER7.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER7");
                REFER7.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER7_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER7_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER7.ReadOnly = true;
                    REFER7.Nullable = false;
                    REFER7.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER7", mDATA_VALUE);
                REFER7.Refresh();
            }

            //--10
            if (REFER8.ReadOnly == false)
            {
                mDATA_VALUE = IGR_SLIP_LINE.GetCellValue("REFER8");
                REFER8.Nullable = true;
                mDATA_TYPE = iConv.ISNull(pDataRow["REFER8_DATA_TYPE"]);
                mDR_CR_YN = iConv.ISNull(pDataRow["REFER8_YN"]);
                if (mDATA_TYPE == "VARCHAR2" && mDR_CR_YN == "Y")
                {
                    REFER8.ReadOnly = true;
                    REFER8.Nullable = false;
                    REFER8.ReadOnly = false;
                }
                IGR_SLIP_LINE.SetCellValue("REFER8", mDATA_VALUE);
                REFER8.Refresh();
            }
        } 

        private void Init_Default_Value()
        {
            int mPreviousRowPosition = IDA_SLIP_LINE.CurrentRowPosition() - 1;
            object mPrevious_Code;
            object mPrevious_Name;
            string mData_Type;
            string mLookup_Type;

            if (mPreviousRowPosition > -1
                && iConv.ISNull(REMARK.EditValue) == string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"]) != string.Empty)
            {//REMARK.
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"];
                REMARK.EditValue = mPrevious_Name;
            }

            //1
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(MANAGEMENT1.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT1.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_LOOKUP_TYPE"]))
            {//MANAGEMENT1_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_DESC"];

                MANAGEMENT1.EditValue = mPrevious_Code;
                MANAGEMENT1_DESC.EditValue = mPrevious_Name;
            }
            //2
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(MANAGEMENT2.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT2.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_LOOKUP_TYPE"]))
            {//MANAGEMENT2_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_DESC"];

                MANAGEMENT2.EditValue = mPrevious_Code;
                MANAGEMENT2_DESC.EditValue = mPrevious_Name;
            }
            //3
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER1_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER1_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER1.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER1.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_LOOKUP_TYPE"]))
            {//REFER1_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_DESC"];

                REFER1.EditValue = mPrevious_Code;
                REFER1_DESC.EditValue = mPrevious_Name;
            }
            //4
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER2_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER2_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER2.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER2.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_LOOKUP_TYPE"]))
            {//REFER2_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_DESC"];

                REFER2.EditValue = mPrevious_Code;
                REFER2_DESC.EditValue = mPrevious_Name;
            }
            //5
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER3_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER3_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER3.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER3.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_LOOKUP_TYPE"]))
            {//REFER3_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_DESC"];

                REFER3.EditValue = mPrevious_Code;
                REFER3_DESC.EditValue = mPrevious_Name;
            }
            //6
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER4_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER4_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER4.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER4.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_LOOKUP_TYPE"]))
            {//REFER4_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_DESC"];

                REFER4.EditValue = mPrevious_Code;
                REFER4_DESC.EditValue = mPrevious_Name;
            }
            //7
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER5_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER5_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER5.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER5.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_LOOKUP_TYPE"]))
            {//REFER5_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_DESC"];

                REFER5.EditValue = mPrevious_Code;
                REFER5_DESC.EditValue = mPrevious_Name;
            }
            //8
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER6_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER6_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER6.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER6.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_LOOKUP_TYPE"]))
            {//REFER6_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_DESC"];

                REFER6.EditValue = mPrevious_Code;
                REFER6_DESC.EditValue = mPrevious_Name;
            }
            //9
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER7_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER7_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER7.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER7.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_LOOKUP_TYPE"]))
            {//REFER7_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_DESC"];

                REFER7.EditValue = mPrevious_Code;
                REFER7_DESC.EditValue = mPrevious_Name;
            }
            //10
            mData_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER8_DATA_TYPE"]);
            mLookup_Type = iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REFER8_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iConv.ISNull(REFER8.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER8.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_LOOKUP_TYPE"]))
            {//REFER8_LOOKUP_TYPE
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_DESC"];

                REFER8.EditValue = mPrevious_Code;
                REFER8_DESC.EditValue = mPrevious_Name;
            }
        }

        //private void Init_Default_Value()
        //{
        //    int mPreviousRowPosition = idaSLIP_LINE.CurrentRowPosition() - 1;
        //    object mPrevious_Code;
        //    object mPrevious_Name;
        //    string mData_Type;
        //    string mLookup_Type;

        //    if (mPreviousRowPosition > -1
        //        && iConv.ISNull(REMARK.EditValue) == string.Empty
        //        && iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"]) != string.Empty)
        //    {//REMARK.
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"];
        //        REMARK.EditValue = mPrevious_Name;
        //    }

        //    //1
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(MANAGEMENT1.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            MANAGEMENT1.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_LOOKUP_TYPE"]))
        //    {//MANAGEMENT1_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_DESC"];

        //        MANAGEMENT1.EditValue = mPrevious_Code;
        //        MANAGEMENT1_DESC.EditValue = mPrevious_Name;
        //    }
        //    //2
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(MANAGEMENT2.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            MANAGEMENT2.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_LOOKUP_TYPE"]))
        //    {//MANAGEMENT2_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_DESC"];

        //        MANAGEMENT2.EditValue = mPrevious_Code;
        //        MANAGEMENT2_DESC.EditValue = mPrevious_Name;
        //    }
        //    //3
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER1_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER1.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER1.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_LOOKUP_TYPE"]))
        //    {//REFER1_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_DESC"];

        //        REFER1.EditValue = mPrevious_Code;
        //        REFER1_DESC.EditValue = mPrevious_Name;
        //    }
        //    //4
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER2_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER2.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER2.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_LOOKUP_TYPE"]))
        //    {//REFER2_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_DESC"];

        //        REFER2.EditValue = mPrevious_Code;
        //        REFER2_DESC.EditValue = mPrevious_Name;
        //    }
        //    //5
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER3_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER3.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER3.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_LOOKUP_TYPE"]))
        //    {//REFER3_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_DESC"];

        //        REFER3.EditValue = mPrevious_Code;
        //        REFER3_DESC.EditValue = mPrevious_Name;
        //    }
        //    //6
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER4_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER4.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER4.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_LOOKUP_TYPE"]))
        //    {//REFER4_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_DESC"];

        //        REFER4.EditValue = mPrevious_Code;
        //        REFER4_DESC.EditValue = mPrevious_Name;
        //    }
        //    //7
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER5_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER5.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER5.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_LOOKUP_TYPE"]))
        //    {//REFER5_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_DESC"];

        //        REFER5.EditValue = mPrevious_Code;
        //        REFER5_DESC.EditValue = mPrevious_Name;
        //    }
        //    //8
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER6_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER6.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER6.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_LOOKUP_TYPE"]))
        //    {//REFER6_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_DESC"];

        //        REFER6.EditValue = mPrevious_Code;
        //        REFER6_DESC.EditValue = mPrevious_Name;
        //    }
        //    //9
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER7_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER7.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER7.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_LOOKUP_TYPE"]))
        //    {//REFER7_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_DESC"];

        //        REFER7.EditValue = mPrevious_Code;
        //        REFER7_DESC.EditValue = mPrevious_Name;
        //    }
        //    //10
        //    mData_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER8_DATA_TYPE"));
        //    mLookup_Type = iConv.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
        //    if (mData_Type == "NUMBER".ToString())
        //    {
        //    }
        //    else if (mData_Type == "RATE".ToString())
        //    {
        //    }
        //    else if (mData_Type == "DATE".ToString())
        //    {
        //        if (iConv.ISNull(REFER8.EditValue) == string.Empty && iConv.ISNull(SLIP_DATE.EditValue) != string.Empty)
        //        {
        //            REFER8.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
        //        }
        //    }
        //    if (mPreviousRowPosition > -1
        //        && mLookup_Type != string.Empty
        //        && mLookup_Type == iConv.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_LOOKUP_TYPE"]))
        //    {//REFER8_LOOKUP_TYPE
        //        mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8"];
        //        mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_DESC"];

        //        REFER8.EditValue = mPrevious_Code;
        //        REFER8_DESC.EditValue = mPrevious_Name;
        //    }
        //}

        private void Init_Currency_Code(string pInit_YN)
        {
            //if (iConv.ISNull(idaSLIP_LINE.CurrentRow["CURRENCY_ENABLED_FLAG"], "N") == "Y")
            //{
            //    CURRENCY_DESC.ReadOnly = false;
            //    CURRENCY_DESC.Insertable = true;
            //    CURRENCY_DESC.Updatable = true;
            //    CURRENCY_DESC.TabStop = true;
            //}
            //else
            //{
            //    CURRENCY_DESC.ReadOnly = true;
            //    CURRENCY_DESC.Insertable = false;
            //    CURRENCY_DESC.Updatable = false;
            //    CURRENCY_DESC.TabStop = false;
            //    if (pInit_YN == "Y")
            //    {
            //        CURRENCY_CODE.EditValue = mCurrency_Code;
            //        CURRENCY_DESC.EditValue = mCurrency_Code;
            //        Init_Currency_Amount();
            //    }
            //}
            //CURRENCY_CODE.Invalidate();
            //CURRENCY_DESC.Invalidate();
        }

        private void Init_Currency_Amount()
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) == string.Empty || iConv.ISNull(CURRENCY_CODE.EditValue) == mCurrency_Code)
            {
                if (iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue) != Convert.ToDecimal(0))
                {
                    EXCHANGE_RATE.EditValue = null;
                }
                if (iConv.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    GL_CURRENCY_AMOUNT.EditValue = null;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                GL_CURRENCY_AMOUNT.ReadOnly = true;
                GL_CURRENCY_AMOUNT.Insertable = false;
                GL_CURRENCY_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                GL_CURRENCY_AMOUNT.TabStop = false;
            }
            else
            {
                if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REF_SLIP_FLAG"]) != string.Empty)
                {
                    EXCHANGE_RATE.ReadOnly = true;
                    EXCHANGE_RATE.Insertable = false;
                    EXCHANGE_RATE.Updatable = false;
                    EXCHANGE_RATE.TabStop = false;

                    //원전표인 경우 금액수정 불가//
                    if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["REF_SLIP_FLAG"]) == "S")
                    {
                        GL_CURRENCY_AMOUNT.ReadOnly = true;
                        GL_CURRENCY_AMOUNT.Insertable = false;
                        GL_CURRENCY_AMOUNT.Updatable = false;
                        GL_CURRENCY_AMOUNT.TabStop = false;
                    }
                }
                else
                {
                    EXCHANGE_RATE.ReadOnly = false;
                    EXCHANGE_RATE.Insertable = true;
                    EXCHANGE_RATE.Updatable = true;
                    EXCHANGE_RATE.TabStop = true;

                    GL_CURRENCY_AMOUNT.ReadOnly = false;
                    GL_CURRENCY_AMOUNT.Insertable = true;
                    GL_CURRENCY_AMOUNT.Updatable = true;
                    GL_CURRENCY_AMOUNT.TabStop = true;
                }
            }
            EXCHANGE_RATE.Refresh();
            GL_CURRENCY_AMOUNT.Refresh();
        }

        // 부가세 관련 설정 제어 - 세액/공급가액(세액 * 10)
        private void Init_VAT_Amount()
        {
            object mVAT_ENABLED_FLAG = IDA_SLIP_LINE.CurrentRow["VAT_ENABLED_FLAG"];
            if (iConv.ISNull(mVAT_ENABLED_FLAG, "N") != "Y")
            {
                return;
            }

            IDC_GET_ACCOUNT_DEFAULT_VALUE.ExecuteNonQuery();
            decimal vVAT_RATE = iConv.ISDecimaltoZero(IDC_GET_ACCOUNT_DEFAULT_VALUE.GetCommandParamValue("O_VAT_RATE"));

            decimal mGL_AMOUNT = iConv.ISDecimaltoZero(GL_AMOUNT.EditValue);
            decimal mSUPPLY_AMOUNT = mGL_AMOUNT * vVAT_RATE; //공급가액 설정.

            Set_Management_Value("SUPPLY_AMOUNT", mSUPPLY_AMOUNT, null);
            Set_Management_Value("VAT_AMOUNT", mGL_AMOUNT, null);
        }

        //예산부서 동기화
        private void Init_Budget_Dept()
        {
            int mPreviousRowPosition = IDA_SLIP_LINE.CurrentRowPosition() - 1;
            object mPrevious_ID;
            object mPrevious_Code;
            object mPrevious_Name;

            if (mPreviousRowPosition > -1
                && iConv.ISNull(BUDGET_DEPT_ID_L.EditValue) == string.Empty
                && iConv.ISNull(IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_ID"]) != string.Empty)
            {//budget dept
                mPrevious_ID = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_ID"];
                mPrevious_Code = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_CODE"];
                mPrevious_Name = IDA_SLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_NAME"];

                BUDGET_DEPT_NAME_L.EditValue = mPrevious_Name;
                BUDGET_DEPT_CODE_L.EditValue = mPrevious_Code;
                BUDGET_DEPT_ID_L.EditValue = mPrevious_ID;
            }
            else
            {
                BUDGET_DEPT_NAME_L.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("BUDGET_DEPT_NAME");
                BUDGET_DEPT_CODE_L.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("BUDGET_DEPT_CODE"); ;
                BUDGET_DEPT_ID_L.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("BUDGET_DEPT_ID");
            }
        }

        //부서 
        private void Init_Dept()
        {
            if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == "DEPT" && 
                iConv.ISNull(MANAGEMENT1.EditValue) == String.Empty)
            {
                MANAGEMENT1_DESC.EditValue = BUDGET_DEPT_NAME_L.EditValue;
                MANAGEMENT1.EditValue = BUDGET_DEPT_CODE_L.EditValue;
            }
            else if (iConv.ISNull(IDA_SLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == "DEPT")
            {
                //2014-02-18 전호수 추가 : 신전무님 요청사항 
                //사용부서가 있어도 예산부서로 강제 변경
                MANAGEMENT1_DESC.EditValue = BUDGET_DEPT_NAME_L.EditValue;
                MANAGEMENT1.EditValue = BUDGET_DEPT_CODE_L.EditValue;
            }
        }

        //관리항목 LOOKUP 선택시 처리.
        private void Init_SELECT_LOOKUP(object pManagement_Type)
        {
            string mMANAGEMENT = iConv.ISNull(pManagement_Type);
        }

        private bool Init_DPR_ASSET_SUM_AMOUNT()
        {
            decimal mSUPPLY_AMOUNT = 0;
            decimal mVAT_AMOUNT = 0;
            decimal mCOUNT = 0;
            decimal mSUM_SUPPLY_AMOUNT = 0;
            decimal mSUM_VAT_AMOUNT = 0;
            decimal mSUM_COUNT = 0;

            int mIDX_ITEM_CONTENTS = igrDPR_ASSET.GetColumnToIndex("ITEM_CONTENTS");
            int mIDX_VAT_ASSET_GB = igrDPR_ASSET.GetColumnToIndex("VAT_ASSET_GB");
            int mIDX_SUPPLY_AMOUNT = igrDPR_ASSET.GetColumnToIndex("SUPPLY_AMOUNT");
            int mIDX_VAT_AMOUNT = igrDPR_ASSET.GetColumnToIndex("VAT_AMOUNT");
            int mIDX_COUNT = igrDPR_ASSET.GetColumnToIndex("ASSET_COUNT");
            for (int r = 0; r < igrDPR_ASSET.RowCount; r++)
            {
                mSUPPLY_AMOUNT = mSUPPLY_AMOUNT + iConv.ISDecimaltoZero(igrDPR_ASSET.GetCellValue(r, mIDX_SUPPLY_AMOUNT));
                mVAT_AMOUNT = mVAT_AMOUNT + iConv.ISDecimaltoZero(igrDPR_ASSET.GetCellValue(r, mIDX_VAT_AMOUNT));
                mCOUNT = mCOUNT + iConv.ISDecimaltoZero(igrDPR_ASSET.GetCellValue(r, mIDX_COUNT));

                if ((mSUPPLY_AMOUNT + mVAT_AMOUNT) != 0 && iConv.ISNull(igrDPR_ASSET.GetCellValue(r, mIDX_ITEM_CONTENTS)) == string.Empty)
                {//공급가액, 부가세 등록했는데 품목 등록 안함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10523"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (mSUPPLY_AMOUNT == 0 && iConv.ISNull(igrDPR_ASSET.GetCellValue(r, mIDX_ITEM_CONTENTS)) != string.Empty)
                {//공급가액 등록 않했는데 품목 등록함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (mVAT_AMOUNT == 0 && iConv.ISNull(igrDPR_ASSET.GetCellValue(r, mIDX_ITEM_CONTENTS)) != string.Empty)
                {//부가세 등록 않했는데 품목 등록함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10281"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if ((mSUPPLY_AMOUNT + mVAT_AMOUNT) != 0 && mCOUNT == 0)
                {//공급가액, 부가세 등록했는데 수량 등록 안함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10206"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (mSUPPLY_AMOUNT == 0 && mCOUNT != 0)
                {//공급가액, 부가세 등록했는데 수량 등록 안함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10523"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (mVAT_AMOUNT == 0 && mCOUNT != 0)
                {//공급가액, 부가세 등록했는데 수량 등록 안함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                mSUM_SUPPLY_AMOUNT = mSUM_SUPPLY_AMOUNT + mSUPPLY_AMOUNT;
                mSUM_VAT_AMOUNT = mSUM_VAT_AMOUNT + mVAT_AMOUNT;
                mSUM_COUNT = mSUM_COUNT + mCOUNT;
            }
            S_SUM_SUPPLY_AMOUNT.EditValue = mSUM_SUPPLY_AMOUNT;
            S_SUM_VAT_AMOUNT.EditValue = mSUM_VAT_AMOUNT;
            S_SUM_COUNT.EditValue = mSUM_COUNT;

            return true;
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

            string vREAD_FLAG = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_READ_FLAG"));
            string vUSER_TYPE = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_USER_TYPE"));
            string vPRINT_FLAG = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_PRINT_FLAG"));

            decimal vASSEMBLY_INFO_ID = iConv.ISDecimaltoZero(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_INFO_ID"));
            string vASSEMBLY_ID = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_ID"));
            string vASSEMBLY_NAME = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_NAME"));
            string vASSEMBLY_FILE_NAME = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_FILE_NAME"));

            string vASSEMBLY_VERSION = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_VERSION"));
            string vDIR_FULL_PATH = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_DIR_FULL_PATH"));

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
                        if (iConv.ISNull(vRow["REPORT_FILE_NAME"]) != string.Empty)
                        {
                            vReportFileName = iConv.ISNull(vRow["REPORT_FILE_NAME"]);
                            vReportFileNameTarget = string.Format("_{0}", vReportFileName);
                        }
                        if (iConv.ISNull(vRow["REPORT_PATH_FTP"]) != string.Empty)
                        {
                            vPathReportFTP = iConv.ISNull(vRow["REPORT_PATH_FTP"]);
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

        private void XLPrinting_Main(string pOutput_Type)
        {
            object vSlip_Header_id;
            object vSlip_Date;
            object vSlip_Num;

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

             
            vSlip_Header_id = SLIP_INTERFACE_HEADER_ID.EditValue;
            vSlip_Date = SLIP_DATE.EditValue;
            vSlip_Num = SLIP_NUM.EditValue;
        
            AssmblyRun_Manual("FCMF0212", vSlip_Header_id, vSlip_Date, vSlip_Num);
            //IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", SLIP_DATE.EditValue);
            //IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0206");
            //IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            //string vREPORT_TYPE = iConv.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            //if (vREPORT_TYPE.ToUpper() == "BSK")
            //{
            //    XLPrinting_BSK(pOutput_Type);
            //}
            //else
            //{
            //    XLPrinting(pOutput_Type);
            //}

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
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
                    Search_DB();
                }
                //전표 행 위로 추가 사용 안함 ==> 라인 SEQ 제어문제때문에//
                //else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                //{
                //    if (idaSLIP_LINE.IsFocused)
                //    {
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
                //            idaSLIP_HEADER.SetSelectParamValue("W_HEADER_ID", 0);
                //            idaSLIP_HEADER.Fill();

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
                    if (IDA_SLIP_LINE.IsFocused)
                    {
                        Insert_Silp_Line();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    ACCOUNT_CODE.Focus();

                    Init_DR_CR_Amount();    // 차대금액 생성 //
                    Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //

                    if (iConv.ISDecimaltoZero(TOTAL_DR_AMOUNT.EditValue) != iConv.ISDecimaltoZero(TOTAL_CR_AMOUNT.EditValue))
                    {// 차대금액 일치 여부 체크.
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10134"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    IDA_CARD_SLIP_GROUP_APPR.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    mStatus = "QUERY";
                    if (IDA_CARD_SLIP_GROUP_APPR.IsFocused)
                    {
                        IDA_SLIP_LINE.Cancel();
                        IDA_CARD_SLIP_GROUP_APPR.Cancel();
                    }
                    else if (IDA_SLIP_LINE.IsFocused)
                    {
                        IDA_SLIP_LINE.Cancel();
                        Init_Total_GL_Amount();  //합계 금액 재 계산 //
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CARD_SLIP_GROUP_APPR.IsFocused)
                    {
                        for (int r = 0; r < IGR_SLIP_LINE.RowCount; r++)
                        {
                            IDA_SLIP_LINE.Delete();
                        }
                        IDA_CARD_SLIP_GROUP_APPR.Delete();
                    }
                    else if (IDA_SLIP_LINE.IsFocused)
                    {
                        IDA_SLIP_LINE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main("EXCEL");
                }
            }
        }

        private void Insert_Silp_Line()
        {
            IDA_SLIP_LINE.AddUnder();
            InsertSlipLine();
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0261_SLIP_Load(object sender, EventArgs e)
        {           
            // 회계장부 정보 설정.
            GetAccountBook();

            //서브판넬 
            Init_Sub_Panel(false, "ALL");

            //전표 복사 버튼 맨 앞으로 가져오기
            //btnGET_BALANCE_STATEMENT.BringToFront();
            BUDGET_REMAIN_AMOUNT.BringToFront(); 

            // 콤퍼넌트 동기화.
            //Init_Currency_Code();
            ibtSUB_FORM.Visible = false;

            IDA_CARD_SLIP_GROUP_APPR.FillSchema();
        }

        private void FCMF0261_SLIP_Shown(object sender, EventArgs e)
        {
            BTN_LINE_INSERT.BringToFront();
            BTN_LINE_DELETE.BringToFront();
            BTN_SAVE.BringToFront();

            Search_DB();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void CURRENCY_DESC_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_Currency_Amount();
        }

        private void ibtSUB_FORM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("ACCOUNT_DR_CR")) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_DR_CR.Focus();
                return;
            }
            if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("ACCOUNT_CONTROL_ID")) == string.Empty)
            {// 계정과목.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_CODE.Focus();
                return;
            }
            if (iConv.ISNull(IGR_SLIP_LINE.GetCellValue("CURRENCY_CODE")) == string.Empty)
            {// 통화
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CURRENCY_DESC.Focus();
                return;
            }
            if (mCurrency_Code.ToString() != IGR_SLIP_LINE.GetCellValue("CURRENCY_CODE").ToString() 
                  && iConv.ISDecimaltoZero(IGR_SLIP_LINE.GetCellValue("EXCHANGE_RATE")) == Convert.ToInt32(0))
            {// 입력통화와 기본 통화가 다를경우 환율입력 체크.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10125"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                EXCHANGE_RATE.Focus();
                return;
            }
            if (mCurrency_Code.ToString() != IGR_SLIP_LINE.GetCellValue("CURRENCY_CODE").ToString() 
                  && iConv.ISDecimaltoZero(IGR_SLIP_LINE.GetCellValue("GL_CURRENCY_AMOUNT")) == Convert.ToInt32(0))
            {// 입력통화와 기본 통화가 다를경우 외화금액 체크.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10127"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                GL_CURRENCY_AMOUNT.Focus();
                return;
            }
                         
            System.Windows.Forms.DialogResult dlgResult;
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor; 

            if (iConv.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "AP_VAT".ToString())
            {
                S_SUPPLY_AMOUNT.EditValue = Get_Management_Value("SUPPLY_AMOUNT");   //공급가액 설정.
                S_VAT_AMOUNT.EditValue = Get_Management_Value("VAT_AMOUNT");      //세액 설정.

                //서브판넬 
                Init_Sub_Panel(true, "AP_VAT");
            } 
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (IDA_SLIP_LINE.CurrentRow != null && IDA_SLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_GL_Amount();
            }
        }

        private void GL_CURRENCY_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (IDA_SLIP_LINE.CurrentRow != null && IDA_SLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_GL_Amount();
            }
        }

        private void GL_AMOUNT_EditValueChanged(object pSender)
        {
            if (IDA_SLIP_LINE.CurrentRow != null && IDA_SLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_VAT_Amount();
            }
        }

        private void GL_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
        }

        private void BTN_TRANSFER_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (iString.ISNull(HEADER_ID.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10118"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
             
            IDC_SET_SLIP_TRANSFER.SetCommandParamValue("W_SYS_DATE", mSys_Date);
            IDC_SET_SLIP_TRANSFER.SetCommandParamValue("W_SESSION_ID", mSession_ID);
            IDC_SET_SLIP_TRANSFER.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_SLIP_TRANSFER.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_SLIP_TRANSFER.GetCommandParamValue("O_MESSAGE"));
            object vSLIP_NUM = IDC_SET_SLIP_TRANSFER.GetCommandParamValue("O_SLIP_NUM");
            object vSLIP_INTERFACE_HEADER_ID = IDC_SET_SLIP_TRANSFER.GetCommandParamValue("O_INTERFACE_HEADER_ID");

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_SLIP_TRANSFER.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_SLIP_TRANSFER.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            SLIP_NUM.EditValue = vSLIP_NUM;
            SLIP_INTERFACE_HEADER_ID.EditValue = vSLIP_INTERFACE_HEADER_ID;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BTN_REQ_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (iString.ISNull(HEADER_ID.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10118"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_APPROVAL_REQUEST_CANCEL.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_CANCEL.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_APPROVAL_REQUEST_CANCEL.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_APPROVAL_REQUEST_CANCEL.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_APPROVAL_REQUEST_CANCEL.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Search_DB();
        }

        private void BTN_CLOSED_ButtonClick_1(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
            IDA_SLIP_LINE.Cancel();
            IDA_CARD_SLIP_GROUP_APPR.Cancel();

            DialogResult = DialogResult.Cancel;
        
            this.Close();
        }

        private void S_BTN_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "AP_VAT".ToString())
            {
                idaDPR_ASSET.AddUnder();

                int vIDX_ASSET_GB_DESC = igrDPR_ASSET.GetColumnToIndex("VAT_ASSET_GB_DESC");
                igrDPR_ASSET.CurrentCellMoveTo(vIDX_ASSET_GB_DESC);
                igrDPR_ASSET.CurrentCellActivate(vIDX_ASSET_GB_DESC);
                igrDPR_ASSET.Focus();
            }
        }

        private void S_BTN_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Delete();
        }

        private void S_BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
        }

        private void S_BTN_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            S_SUPPLY_AMOUNT.Focus();
            if (Init_DPR_ASSET_SUM_AMOUNT() == false)
            {
                return;
            }

            //서브판넬 
            Init_Sub_Panel(false, "AP_VAT");

            MANAGEMENT1.Focus();    //focus 이동 
        }

        private void BTN_LINE_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //부가세 항목일 경우 부가세 유형 선택 필수//
            if(VAT_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                if(iConv.ISNull(VAT_TAX_TYPE.EditValue)  == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(VAT_TAX_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
             
            Insert_Silp_Line();
        }

        private void BTN_LINE_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_SLIP_LINE.Delete();
        }

        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CARD_SLIP_GROUP_APPR.Update(); 
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
            //서브판넬 
            Init_Sub_Panel(false, "AP_VAT");

            MANAGEMENT1.Focus();    //focus 이동 
        }

        private void VAT_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if(iConv.ISNull(IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_VAT_FLAG")) == "N")
            {
                return;
            }

            if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                BASE_APPR_AMOUNT.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_BASE_APPR_AMOUNT");
                BASE_VAT_AMOUNT.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_BASE_VAT_AMOUNT");
                BASE_TOTAL_AMOUNT.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_BASE_TOTAL_AMOUNT");

                if (iConv.ISNull(VAT_TAX_TYPE.EditValue) == String.Empty)
                {
                    VAT_TAX_TYPE.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_VAT_TAX_TYPE");
                    VAT_TAX_TYPE_NAME.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_VAT_TAX_TYPE_NAME");
                }

                VAT_TAX_TYPE.ReadOnly = false;
                VAT_TAX_TYPE.Insertable = true;
                VAT_TAX_TYPE.Updatable = true;
            }
            else
            {
                VAT_TAX_TYPE.EditValue = null;
                VAT_TAX_TYPE_NAME.EditValue = null;

                BASE_APPR_AMOUNT.EditValue = iConv.ISDecimaltoZero(IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_BASE_APPR_AMOUNT"), 0) +
                                                iConv.ISDecimaltoZero(IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_BASE_VAT_AMOUNT"), 0);
                BASE_VAT_AMOUNT.EditValue = 0;
                BASE_TOTAL_AMOUNT.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_BASE_TOTAL_AMOUNT");

                VAT_TAX_TYPE.ReadOnly = true;
                VAT_TAX_TYPE.Insertable = false;
                VAT_TAX_TYPE.Updatable = false;
            }
            VAT_TAX_TYPE.Refresh();
            Init_Total_GL_Amount();
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
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_DEPT_ID", isAppInterfaceAdv1.DEPT_ID);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaSLIP_NUM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
        }

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", "VALUE1 = 'GE'", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaSLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", "VALUE1 = 'GE'", "Y");
        }

        private void ilaREQ_PAYABLE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("PAYABLE_TYPE", "Y");
        }

        private void ilaREQ_BANK_ACCOUNT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildREQ_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BUDGET_DEPT_H_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", SLIP_DATE.EditValue);
        }

        private void ILA_VAT_TAX_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VAT_TAX_TYPE_AP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaBUDGET_DEPT_L_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", SLIP_DATE.EditValue);
        }

        private void ilaBUDGET_DEPT_L_SelectedRowData(object pSender)
        {
            Init_Dept();
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaVAT_ASSET_GB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_ASSET_GB", "Y");
        }

        private void ilaACCOUNT_DR_CR_SelectedRowData(object pSender)
        {
            //전호수주석 : 관리항목 변경.
            //Set_Control_Item_Prompt();
            //Init_Control_Management_Value();
            //Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);
            //Init_Set_Item_Need(idaSLIP_LINE.CurrentRow);
            //Init_Default_Value();
            Init_DR_CR_Amount();    // 차대금액 생성 //
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
            GetSubForm();
        }

        private void ilaACCOUNT_CONTROL_SelectedRowData(object pSender)
        {
            Init_Currency_Code("Y");
            Set_Control_Item_Prompt(IDA_SLIP_LINE.CurrentRow.RowState);
            Init_Set_Item_Prompt(IDA_SLIP_LINE.CurrentRow);
            Init_Set_Item_Need(IDA_SLIP_LINE.CurrentRow);
            if (IDA_SLIP_LINE.CurrentRow.RowState == DataRowState.Added)
            {
                Init_Default_Value();
                //거래처 셋팅.
                Set_Management_Value("CUSTOMER", IGR_CARD_SLIP_GROUP_APPR.GetCellValue("VENDOR_CODE"), IGR_CARD_SLIP_GROUP_APPR.GetCellValue("MANAGEMENT_NAME"));
                //부서 셋팅.
                Set_Management_Value("DEPT", IGR_CARD_SLIP_GROUP_APPR.GetCellValue("BUDGET_DEPT_CODE"), IGR_CARD_SLIP_GROUP_APPR.GetCellValue("BUDGET_DEPT_NAME"));
            }
            Init_Dept();
            GetSubForm();
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_SelectedRowData(object pSender)
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) != string.Empty)
            {
                Init_Currency_Amount();
                if (CURRENCY_CODE.EditValue.ToString() != mCurrency_Code.ToString())
                {
                    idcEXCHANGE_RATE.ExecuteNonQuery();
                    EXCHANGE_RATE.EditValue = idcEXCHANGE_RATE.GetCommandParamValue("X_EXCHANGE_RATE");

                    Init_GL_Amount();
                }
            }
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_DEPT_ID", BUDGET_DEPT_ID_L.EditValue);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT1_ID", "Y", IGR_SLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
        }

        private void ilaMANAGEMENT2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("MANAGEMENT2_ID", "Y", IGR_SLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
        }

        private void ilaREFER1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER1_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
        }

        private void ilaREFER2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER2_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
        }

        private void ilaREFER3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER3_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
        }

        private void ilaREFER4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER4_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
        }

        private void ilaREFER5_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER5_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
        }

        private void ilaREFER6_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER6_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
        }

        private void ilaREFER7_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER7_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
        }

        private void ilaREFER8_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetManagementParameter("REFER8_ID", "Y", IGR_SLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
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

        private void IDA_CARD_SLIP_GROUP_APPR_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            if (iConv.ISNull(pBindingManager.DataRow["ORI_VAT_FLAG"]) == "N")
            {
                VAT_FLAG.ReadOnly = true;
                VAT_TAX_TYPE.ReadOnly = true;
                VAT_TAX_TYPE.Insertable = false;
                VAT_TAX_TYPE.Updatable = false;
            }
            else if (iConv.ISNull(pBindingManager.DataRow["VAT_FLAG"]) == "Y")
            {
                VAT_FLAG.ReadOnly = false;
                if(iConv.ISNull(VAT_TAX_TYPE.EditValue) == String.Empty)
                {
                    VAT_TAX_TYPE.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_VAT_TAX_TYPE");
                    VAT_TAX_TYPE_NAME.EditValue = IGR_CARD_SLIP_GROUP_APPR.GetCellValue("ORI_VAT_TAX_TYPE_NAME");
                }
                VAT_TAX_TYPE.ReadOnly = false;
                VAT_TAX_TYPE.Insertable = true;
                VAT_TAX_TYPE.Updatable = true;
            }
            else
            {
                VAT_FLAG.ReadOnly = false;
                VAT_TAX_TYPE.ReadOnly = true;
                VAT_TAX_TYPE.Insertable = false;
                VAT_TAX_TYPE.Updatable = false;
            }
            VAT_TAX_TYPE.Refresh(); 

        }

        private void IDA_CARD_SLIP_GROUP_APPR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["USE_PERSON_ID"]) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("{0}{1}", "&&FIELD_NAME:=", Get_Edit_Prompt(USE_PERSON_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaSLIP_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //if (iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            //{// 예산부서
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iConv.ISNull(e.Row["ACCOUNT_DR_CR"]) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {// 계정과목.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ACCOUNT_CODE"]) == string.Empty)
            {// 계정과목
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //예산관리 계정에 대해서 예산부서 검증.
            if (iConv.ISNull(e.Row["BUDGET_ENABLED_FLAG"]) == "Y" && iConv.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10458"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {// 통화
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CURRENCY_ENABLED_FLAG"]) == "Y".ToString())
            {// 외화 계좌.
                if (mCurrency_Code.ToString() != e.Row["CURRENCY_CODE"].ToString() && iConv.ISDecimaltoZero(e.Row["EXCHANGE_RATE"]) == Convert.ToInt32(0))
                {// 입력통화와 기본 통화가 다를경우 환율입력 체크.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10125"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
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

        private void idaSLIP_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            try
            {
                if (e.Row.RowState != DataRowState.Added)
                {
                    if (e.Row["CONFIRM_YN"].ToString() == "Y".ToString())
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10408"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_SLIP_LINE.MoveFirst(this.Name);
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
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

            //Init_Currency_Code("N");
            //Init_Currency_Amount();
            //GetSubForm();
            //if (SLIP_QUERY_STATUS.EditValue.ToString() != "QUERY".ToString())
            //{
            //    Init_DR_CR_Amount();    // 차대금액 생성 //
            //}            
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

        private void idaDPR_ASSET_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_DPR_ASSET_SUM_AMOUNT();
        }
         
        private void idaDPR_ASSET_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["VAT_ASSET_GB"]) == string.Empty)
            {// 자산구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Asset Type(자산구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}