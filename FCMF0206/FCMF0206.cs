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

namespace FCMF0206
{
    public partial class FCMF0206 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private ISFileTransferAdv mFileTransfer;
        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mPassive = "N";
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = string.Empty;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.
        private bool mSave_Appr_Status = false;
        private string mATT_FILE_YN = "N";                         //첨부파일 사용여부.
        private string mAPPROVAL_YN = "N";                         //승인단계 사용여부.

        bool mIsClickInquiryDetail = false;
        int mInquiryDetailPreX, mInquiryDetailPreY; //마우스 이동 제어.

        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mCurrency_Code;
        object mBudget_Control_YN;
        bool mSUB_SHOW_FLAG = false;
        string gBTN_Req_YN = "N";          // 버튼 승인 여부 //
        string mAuto_Search_Flag = "N";
        string mREMARK_COPY_FLAG = "N";

        object mOffset_Account_Control_ID;
        object mOffset_Account_Code;
        object mOffset_Account_Desc;
        string mOffset_Account_DR_CR;
        object mOffset_Account_DR_CR_Name;

        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #endregion;

        #region ----- Constructor -----

        public FCMF0206()
        {
            InitializeComponent();
        }

        public FCMF0206(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public FCMF0206(Form pMainForm, ISAppInterface pAppInterface, object pGL_Date_FR, object pGL_Date_TO, object pSlip_Num)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iString.ISNull(pSlip_Num) != string.Empty)
            {
                SLIP_DATE_FR_0.EditValue = pGL_Date_FR;
                SLIP_DATE_TO_0.EditValue = pGL_Date_TO;
                SLIP_NUM_0.EditValue = pSlip_Num; 
                mAuto_Search_Flag = "Y";
            }
        }

        public FCMF0206(Form pMainForm, ISAppInterface pAppInterface, object pGL_Date_FR, object pGL_Date_TO,
                        object pSlip_Num, object pAccount_Control_ID, object pAccount_Code, object pAccount_Name)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iString.ISNull(pAccount_Control_ID) != string.Empty)
            {
                SLIP_DATE_FR_0.EditValue = pGL_Date_FR;
                SLIP_DATE_TO_0.EditValue = pGL_Date_TO;
                SLIP_NUM_0.EditValue = pSlip_Num;
                ACCOUNT_CONTROL_ID_0.EditValue = pAccount_Control_ID;
                ACCOUNT_CODE_0.EditValue = pAccount_Code;
                ACCOUNT_DESC_0.EditValue = pAccount_Name;
                if (iString.ISNull(pSlip_Num) != string.Empty)
                {
                    mAuto_Search_Flag = "Y";
                }
                else
                {
                    mAuto_Search_Flag = "L";
                }
            }
        }

        public FCMF0206(Form pMainForm, ISAppInterface pAppInterface, object pGL_Date_FR, object pGL_Date_TO,
                        object pSlip_Num, object pAccount_Control_ID, object pAccount_Code, object pAccount_Name, 
                        object pDept_Name, object pDept_ID)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iString.ISNull(pAccount_Control_ID) != string.Empty)
            {
                SLIP_DATE_FR_0.EditValue = pGL_Date_FR;
                SLIP_DATE_TO_0.EditValue = pGL_Date_TO;
                SLIP_NUM_0.EditValue = pSlip_Num;
                ACCOUNT_CONTROL_ID_0.EditValue = pAccount_Control_ID;
                ACCOUNT_CODE_0.EditValue = pAccount_Code;
                ACCOUNT_DESC_0.EditValue = pAccount_Name;
                DEPT_NAME_0.EditValue = pDept_Name;
                DEPT_ID_0.EditValue = pDept_ID;

                if (iString.ISNull(pSlip_Num) != string.Empty)
                {
                    mAuto_Search_Flag = "Y";
                }
                else
                {
                    mAuto_Search_Flag = "L";
                }
            }
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
            if (iString.ISNull(mBudget_Control_YN) == "Y")
            {
                BUDGET_REMAIN_AMOUNT.Visible = true;
            }
            else
            {
                BUDGET_REMAIN_AMOUNT.Visible = false;
            }

            if (iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_TEMP_BAL_STATEMENT_FLAG")) == "Y")
            {
                btnGET_BALANCE_STATEMENT.Visible = true;
                REF_SLIP_FLAG.Visible = true;
            }
            else
            {
                btnGET_BALANCE_STATEMENT.Visible = false;
                REF_SLIP_FLAG.Visible = false;
            }
            if (iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_TEMP_SLIP_REMARK_FLAG")) == "Y")
            {
                REMARK.LookupAdapter = ILA_TEMP_SLIP_REMARK;
            }
            else
            {
                REMARK.LookupAdapter = null;
            }
            //적요 복사 여부//
            IDC_SLIP_REMARK_COPY_P.SetCommandParamValue("W_ASSEMBLY_ID", "FCMF0206");
            IDC_SLIP_REMARK_COPY_P.SetCommandParamValue("W_SLIP_DATE", DBNull.Value);
            IDC_SLIP_REMARK_COPY_P.ExecuteNonQuery();
            mREMARK_COPY_FLAG = iString.ISNull(IDC_SLIP_REMARK_COPY_P.GetCommandParamValue("O_REMARK_COPY_FLAG"));

            //상쇄계정 존재여부.
            IDC_TEMP_OFFSET_ACCOUNT_P.ExecuteNonQuery();
            mOffset_Account_Control_ID = IDC_TEMP_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_CONTROL_ID");
            mOffset_Account_Code = IDC_TEMP_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_CODE");
            mOffset_Account_Desc = IDC_TEMP_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_DESC");
            mOffset_Account_DR_CR = iString.ISNull(IDC_TEMP_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_DR_CR"));
            mOffset_Account_DR_CR_Name = IDC_TEMP_OFFSET_ACCOUNT_P.GetCommandParamValue("O_ACCOUNT_DR_CR_DESC");

            if (iString.ISNull(mOffset_Account_Control_ID) == string.Empty)
            {
                BTN_OFFSET_ACCOUNT.Visible = false;
            }
            else
            {
                BTN_OFFSET_ACCOUNT.Visible = true;
            }

            //예산사용신청서번호//
            IDC_BUDGET_APPLY_NUM_SHOW_P.ExecuteNonQuery();
            string vAPPLY_NUM_SHOW_YN = iString.ISNull(IDC_BUDGET_APPLY_NUM_SHOW_P.GetCommandParamValue("O_APPLY_NUM_SHOW_YN"));
            if (vAPPLY_NUM_SHOW_YN == "Y")
            {
                BUDGET_APPLY_NUM.Visible = true;
            }
            else
            {
                BUDGET_APPLY_NUM.Visible = false;
            }

            //전표 승인단계 관리 여부.
            IDC_GET_SLIP_CONFIG_P.ExecuteNonQuery();
            mATT_FILE_YN = iString.ISNull(IDC_GET_SLIP_CONFIG_P.GetCommandParamValue("O_ATT_FILE_YN"));
            mAPPROVAL_YN = iString.ISNull(IDC_GET_SLIP_CONFIG_P.GetCommandParamValue("O_APPROVAL_YN"));
            if(mATT_FILE_YN == "N")
            {
                BTN_DOC_ATT_L.Visible = false;
                BTN_FILE_ATTACH.Visible = false;
                CB_DOC_ATT_FLAG.Visible = false;
                igrSLIP_BUDGET_LIST.GridAdvExColElement[igrSLIP_BUDGET_LIST.GetColumnToIndex("DOC_ATT_FLAG")].Visible = 0;
            }
            if (mAPPROVAL_YN == "N")
            {
                BTN_APPR_STEP.Visible = false;
                TP_APPR.TabVisible = false;
            }
            igrSLIP_BUDGET_LIST.ResetDraw = true;
        }

        private void Search_DB()
        {
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            gBTN_Req_YN = "F";
            if (itbSLIP.SelectedTab.TabIndex == 2)
            {
                Search_DB_DETAIL(HEADER_ID.EditValue);
            }
            else
            {
                if (iString.ISNull(SLIP_DATE_FR_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    SLIP_DATE_FR_0.Focus();
                    return;
                }

                if (iString.ISNull(SLIP_DATE_TO_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    SLIP_DATE_TO_0.Focus();
                    return;
                }

                if (Convert.ToDateTime(SLIP_DATE_FR_0.EditValue) > Convert.ToDateTime(SLIP_DATE_TO_0.EditValue))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    SLIP_DATE_FR_0.Focus();
                    return;
                }

                igrSLIP_BUDGET_LIST.LastConfirmChanges();
                idaSLIP_BUDGET_LIST.OraSelectData.AcceptChanges();
                idaSLIP_BUDGET_LIST.Refillable = true; 

                string vGL_NUM = iString.ISNull(igrSLIP_BUDGET_LIST.GetCellValue("GL_NUM"));
                int vCOL_IDX = igrSLIP_BUDGET_LIST.GetColumnToIndex("GL_NUM");
                idaSLIP_BUDGET_LIST.Fill();
                if (iString.ISNull(vGL_NUM) != string.Empty)
                {
                    for (int i = 0; i < igrSLIP_BUDGET_LIST.RowCount; i++)
                    {
                        if (vGL_NUM == iString.ISNull(igrSLIP_BUDGET_LIST.GetCellValue(i, vCOL_IDX)))
                        {
                            igrSLIP_BUDGET_LIST.CurrentCellMoveTo(i, vCOL_IDX);
                            igrSLIP_BUDGET_LIST.CurrentCellActivate(i, vCOL_IDX);
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

                gBTN_Req_YN = "F";
                itbSLIP.SelectedIndex = 1;
                itbSLIP.SelectedTab.Focus();

                igrSLIP_LINE.LastConfirmChanges();
                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;

                IGR_APPROVAL_PERSON.LastConfirmChanges();
                IDA_APPROVAL_PERSON.OraSelectData.AcceptChanges();
                IDA_APPROVAL_PERSON.Refillable = true;

                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;

                idaSLIP_HEADER.SetSelectParamValue("W_HEADER_ID", pSLIP_HEADER_ID);
                try
                {
                    idaSLIP_HEADER.Fill();
                    SLIP_DATE.Focus();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                    return;
                }
                //첨부파일 여부.//
                DOC_ATT_FLAG();

                IGR_APPROVAL_PERSON.LastConfirmChanges();
                IDA_APPROVAL_PERSON.OraSelectData.AcceptChanges();
                IDA_APPROVAL_PERSON.Refillable = true;

                igrSLIP_LINE.LastConfirmChanges();
                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;

                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;

                Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);
                Init_Set_Item_Need(idaSLIP_LINE.CurrentRow);

                igrSLIP_LINE.LastConfirmChanges();
                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;

                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;
            }
        }

        private void Search_REF_LIST()
        {
            IDA_REF_SLIP_LIST.SetSelectParamValue("W_SLIP_FLAG", "N");
            IDA_REF_SLIP_LIST.Fill();

            Init_Sub_Panel(true, "REF_LIST");
            IGR_REF_SLIP_LIST.Focus();
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

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT1_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["MANAGEMENT2_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER1_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER1_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER2_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER2_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER3_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER3_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER4_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER4_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER5_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER5_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER6_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER6_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER7_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER7_DECIMAL_DIGITS"]);
                igrSLIP_LINE.SetCellValue("REFER8_DECIMAL_DIGITS", idaCONTROL_ITEM_PROMPT.CurrentRow["REFER8_DECIMAL_DIGITS"]);
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

                igrSLIP_LINE.SetCellValue("MANAGEMENT1_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("MANAGEMENT2_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER1_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER2_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER3_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER4_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER5_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER6_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER7_DECIMAL_DIGITS", 0);
                igrSLIP_LINE.SetCellValue("REFER8_DECIMAL_DIGITS", 0);
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
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", DEPT_CODE.EditValue);
            }
            else if (mLookup_Type == "COSTCENTER".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", Get_Management_Value("DEPT"));
            }
            else if (mLookup_Type == "BANK_ACCOUNT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", Get_Management_Value("BANK_SITE"));
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
                if (iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    vSLIP_DATE = SLIP_DATE.DateTimeValue.ToShortDateString();
                }
                else if (iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
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

        private void GetSlipNum()
        {
            if (iString.ISNull(DOCUMENT_TYPE.EditValue) == string.Empty)
            {
                return;
            }
            idcSLIP_NUM.SetCommandParamValue("W_DOCUMENT_TYPE", DOCUMENT_TYPE.EditValue);
            idcSLIP_NUM.ExecuteNonQuery();
            SLIP_NUM.EditValue = idcSLIP_NUM.GetCommandParamValue("O_DOCUMENT_NUM");
            GL_NUM.EditValue = SLIP_NUM.EditValue;
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
            ibtSUB_FORM.Left = 777;
            ibtSUB_FORM.Top = 93;
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

        private void Set_Ref_Management(string pManagement, object pManagement_Value, object pVarchar2, object pDate, object pNumber)
        {
            //관리항목 동기화// 
            //거래처 선택시 만기일자 있으면 만기일자 설정//
            Set_Ref_Management_Value(pManagement, "CUSTOMER", "DUE_DATE", pManagement_Value);
            //거래처 선택시 지급방법 설정//
            Set_Ref_Management_Value(pManagement, "CUSTOMER", "PAYMENT_METHOD", pManagement_Value);
            //신용카드 결재일자//
            Set_Ref_Management_Value(pManagement, "CREDIT_CARD", "DUE_DATE", pManagement_Value);
            //공급가액 동기화// 
            Set_Ref_Management_Value(pManagement, "VAT_TAX_TYPE", "SUPPLY_AMOUNT", pManagement_Value, pVarchar2, pDate, pNumber);
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

        //계정잔액명세서 SHOW
        private void Get_Balance_Statement()
        {
            if (iString.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (CONFIRM_YN.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (mSUB_SHOW_FLAG == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //delete temp data : 계정잔액 대상 산출과 전표 저장 완료시 변경//
            Delete_Balance_Remain_TP();

            DialogResult vRESULT;
            FCMF0206_SET vFCMF0206_SET = new FCMF0206_SET(isAppInterfaceAdv1.AppInterface, mSession_ID, mCurrency_Code, "R"
                                                        , ACCOUNT_CONTROL_ID.EditValue, ACCOUNT_CODE.EditValue, ACCOUNT_DESC.EditValue
                                                        , SLIP_DATE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0206_SET, isAppInterfaceAdv1.AppInterface);
            vRESULT = vFCMF0206_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                Set_Insert_Slip_Line();
                Init_Currency_Code("Y");
                Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);
                Init_Total_GL_Amount();
            }
            vFCMF0206_SET.Dispose(); 
            BUDGET_DEPT_NAME_L.Focus();  
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

                        GB_AP_VAT.BringToFront();
                        GB_AP_VAT.Visible = true;
                    }
                    else if (pSub_Panel == "COPY_SLIP")
                    {
                        GB_COPY_DOCUMENT.Left = 100;
                        GB_COPY_DOCUMENT.Top = 15;

                        GB_COPY_DOCUMENT.Width = 540;
                        GB_COPY_DOCUMENT.Height = 145;

                        GB_COPY_DOCUMENT.Border3DStyle = Border3DStyle.Bump;
                        GB_COPY_DOCUMENT.BorderStyle = BorderStyle.Fixed3D;

                        GB_COPY_DOCUMENT.BringToFront();
                        GB_COPY_DOCUMENT.Visible = true;
                    }
                    else if (pSub_Panel == "REF_LIST")
                    {
                        GB_REF_LIST.Left = 433;
                        GB_REF_LIST.Top = 30;

                        GB_REF_LIST.Width = 503;
                        GB_REF_LIST.Height = 140;

                        GB_REF_LIST.Border3DStyle = Border3DStyle.Bump;
                        GB_REF_LIST.BorderStyle = BorderStyle.Fixed3D;

                        GB_REF_LIST.BringToFront();
                        GB_REF_LIST.Visible = true;
                    }
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Left = 65;
                        GB_APPR.Top = 115;

                        GB_APPR.Width = 900;
                        GB_APPR.Height = 240;
                        
                        GB_APPR.Border3DStyle = Border3DStyle.Bump;
                        GB_APPR.BorderStyle = BorderStyle.Fixed3D;

                        //GroupBox 이동//
                        GB_APPR.Controls[0].MouseDown += GB_APPR_MouseDown;
                        GB_APPR.Controls[0].MouseMove += GB_APPR_MouseMove;
                        GB_APPR.Controls[0].MouseUp += GB_APPR_MouseUp;
                        GB_APPR.Controls[1].MouseDown += GB_APPR_MouseDown;
                        GB_APPR.Controls[1].MouseMove += GB_APPR_MouseMove;
                        GB_APPR.Controls[1].MouseUp += GB_APPR_MouseUp;

                        GB_APPR.BringToFront();
                        GB_APPR.Visible = true;
                    } 
                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                itpSLIP_LIST.Enabled = false;
                igbSLIP_HEADER.Enabled = false;
                igbCONFIRM_INFOMATION.Enabled = false;
                igbACCOUNT_LINE.Enabled = false;
                igbSLIP_LINE.Enabled = false;
                TB_INFO.Enabled = false;
            }
            else
            {                
                try
                {
                    if (pSub_Panel == "ALL")
                    {
                        GB_AP_VAT.Visible = false;
                        GB_COPY_DOCUMENT.Visible = false;
                        GB_REF_LIST.Visible = false;
                        GB_APPR.Visible = false; 
                        TB_INFO.Enabled = false;
                    }
                    else if (pSub_Panel == "AP_VAT")
                    {
                        GB_AP_VAT.Visible = false;
                    }
                    else if (pSub_Panel == "COPY_SLIP")
                    {
                        GB_COPY_DOCUMENT.Visible = false;
                    }
                    else if (pSub_Panel == "REF_LIST")
                    {
                        GB_REF_LIST.Visible = false;
                    }
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Visible = false;
                    } 
                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                itpSLIP_LIST.Enabled = true;
                igbSLIP_HEADER.Enabled = true;
                igbCONFIRM_INFOMATION.Enabled = true;
                igbACCOUNT_LINE.Enabled = true;
                igbSLIP_LINE.Enabled = true;
                TB_INFO.Enabled = true;
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
            for (int r = 0; r < idaSLIP_HEADER.SelectRows.Count; r++)
            {
                if (idaSLIP_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    idaSLIP_HEADER.SelectRows[r].RowState == DataRowState.Modified)
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
                for (int r = 0; r < idaSLIP_LINE.SelectRows.Count; r++)
                {
                    if (idaSLIP_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        idaSLIP_LINE.SelectRows[r].RowState == DataRowState.Modified)
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
            itbSLIP.SelectedIndex = 1;
            itbSLIP.SelectedTab.Focus();
            
            SLIP_DATE.EditValue = iDate.ISGetDate(DateTime.Today);

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

            //헤더 예산부서
            BUDGET_DEPT_NAME.EditValue = DEPT_NAME.EditValue;
            BUDGET_DEPT_CODE.EditValue = DEPT_CODE.EditValue;
            BUDGET_DEPT_ID.EditValue = DEPT_ID.EditValue;

            //승인단계 설정.
            Init_SLIP_APPR();
        }

        private void InsertSlipLine()
        {
            Set_Slip_Line_Seq();    //LINE SEQ 채번//

            USE_DATE.EditValue = SLIP_DATE.EditValue;
            CURRENCY_CODE.EditValue = mCurrency_Code;
            CURRENCY_DESC.EditValue = mCurrency_Code;
            Init_Currency_Amount();
            Init_Budget_Dept();
            GL_AMOUNT.EditValue = 0;
            GL_CURRENCY_AMOUNT.EditValue = 0;
            
            if(mREMARK_COPY_FLAG == "Y")
            {
                REMARK.EditValue = H_REMARK.EditValue;
            }
            BUDGET_DEPT_NAME_L.Focus();
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
                mPreviousRowPosition = idaSLIP_LINE.CurrentRowPosition() - 1;
            }
            catch
            {
                mPreviousRowPosition = 0;
            }

            //현재 이전 line seq 
            if (mPreviousRowPosition > -1)
            {
                vPre_Line_Seq = iString.ISDecimaltoZero(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["SLIP_LINE_SEQ"], 1);
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
                    mNextRowPosition = idaSLIP_LINE.CurrentRowPosition() + 1;
                }
                catch
                {
                    mNextRowPosition = 0;
                }

                if (mNextRowPosition == idaSLIP_LINE.CurrentRows.Count)
                {
                    vNext_Line_Seq = 0;
                }
                else
                {
                    vNext_Line_Seq = iString.ISDecimaltoZero(idaSLIP_LINE.CurrentRows[mNextRowPosition]["SLIP_LINE_SEQ"], 1);
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
            igrSLIP_LINE.SetCellValue("SLIP_LINE_SEQ", mSLIP_LINE_SEQ);
        }

        private void Set_Insert_Slip_Line()
        {
            IDA_BALANCE_SLIP_BUDGET.SetSelectParamValue("P_SLIP_DATE", SLIP_DATE.EditValue);
            IDA_BALANCE_SLIP_BUDGET.SetSelectParamValue("P_SESSION_ID", mSession_ID);
            IDA_BALANCE_SLIP_BUDGET.Fill();
            if (IDA_BALANCE_SLIP_BUDGET.SelectRows.Count < 1)
            {
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Not found data, Check data");
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vMOVE_LAST_FLAG = "F";
            igrSLIP_LINE.BeginUpdate();
            int Row_Count = igrSLIP_LINE.RowIndex;
            try
            {
                for (int i = 0; i < IDA_BALANCE_SLIP_BUDGET.SelectRows.Count; i++)
                {
                    if (iString.ISNull(ACCOUNT_CONTROL_ID.EditValue) == string.Empty ||
                        iString.ISDecimaltoZero(GL_AMOUNT.EditValue) == 0)
                    {
                        vMOVE_LAST_FLAG = "Y";
                    }
                    else
                    {
                        if (vMOVE_LAST_FLAG == "F")
                        {
                            idaSLIP_LINE.MoveLast(igrSLIP_LINE.Name);
                            Row_Count = igrSLIP_LINE.RowIndex;
                            Row_Count++;
                            vMOVE_LAST_FLAG = "N";
                        }
                        idaSLIP_LINE.AddUnder();
                    }
                    for (int c = 0; c < igrSLIP_LINE.GridAdvExColElement.Count; c++)
                    {
                        if (igrSLIP_LINE.GridAdvExColElement[c].DataColumn.ToString() != "HEADER_ID")
                        {
                            igrSLIP_LINE.SetCellValue(i + Row_Count, c, IDA_BALANCE_SLIP_BUDGET.OraDataSet().Rows[i][c]);
                        }
                    }
                    Set_Slip_Line_Seq();    //LINE SEQ 채번//

                    //마지막으로 이동.
                    if (vMOVE_LAST_FLAG == "Y")
                    {
                        idaSLIP_LINE.MoveLast(igrSLIP_LINE.Name);
                        Row_Count = igrSLIP_LINE.RowIndex;
                        vMOVE_LAST_FLAG = "N";
                    }
                }
            }
            catch (Exception ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                igrSLIP_LINE.EndUpdate();
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            igrSLIP_LINE.EndUpdate();
             
            //동일한 반제전표 존재 체크//
            int Rec_Cnt = 0;
            decimal vBALANCE_STATEMENT_ID = 0;
            foreach (System.Data.DataRow vRow in idaSLIP_LINE.CurrentRows)
            {
                if (vRow.RowState == DataRowState.Deleted)
                {
                    //
                }
                else if (iString.ISDecimaltoZero(vRow["UNLIQUIDATE_SLIP_LINE_ID"], 0) != 0)
                {
                    vBALANCE_STATEMENT_ID = iString.ISDecimaltoZero(vRow["UNLIQUIDATE_SLIP_LINE_ID"], 0);
                    foreach (System.Data.DataRow vS_Row in idaSLIP_LINE.CurrentRows)
                    {
                        if (vS_Row.RowState == DataRowState.Deleted)
                        {
                            //
                        } 
                        else if (vRow != vS_Row && iString.ISDecimaltoZero(vS_Row["UNLIQUIDATE_SLIP_LINE_ID"], 0) == vBALANCE_STATEMENT_ID)
                        {
                            Rec_Cnt++;
                        }
                    }
                }
            }
            if (Rec_Cnt != 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10588"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void Delete_Balance_Remain_TP()
        {
            //IDA_BATCH_LINE.MoveFirst(igrSLIP_LINE.Name);
            string mSTATUS = "F";
            string mMESSAGE = null;
            try
            {
                IDC_DELETE_BALANCE_REMAIN_TP.SetCommandParamValue("P_SESSION_ID", mSession_ID);
                IDC_DELETE_BALANCE_REMAIN_TP.ExecuteNonQuery();
                mSTATUS = iString.ISNull(IDC_DELETE_BALANCE_REMAIN_TP.GetCommandParamValue("O_STATU"));
                mMESSAGE = iString.ISNull(IDC_DELETE_BALANCE_REMAIN_TP.GetCommandParamValue("O_MESSAGE"));
                if (IDC_DELETE_BALANCE_REMAIN_TP.ExcuteError || mSTATUS == "F")
                {
                    isAppInterfaceAdv1.OnAppMessage(mMESSAGE);
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }
        
        private void Init_GL_Amount()
        {
            if (iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue) == 0)
            {
                return;
            }
            else if (iString.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue) == 0)
            {
                return;
            }
            if (iString.ISNull(REF_SLIP_FLAG.EditValue) == "R" || iString.ISNull(REF_SLIP_FLAG.EditValue) == "S")
            {
                return;
            }

            try
            {
                IDC_BASE_AMOUNT_P.SetCommandParamValue("W_CURRENCY_TO", mCurrency_Code);
                IDC_BASE_AMOUNT_P.SetCommandParamValue("W_CURR_AMOUNT", GL_CURRENCY_AMOUNT.EditValue);
                IDC_BASE_AMOUNT_P.SetCommandParamValue("W_EXCHANGE_RATE", EXCHANGE_RATE.EditValue);
                IDC_BASE_AMOUNT_P.SetCommandParamValue("W_EXCH_TYPE", "SLIP");
                IDC_BASE_AMOUNT_P.ExecuteNonQuery();
                GL_AMOUNT.EditValue = IDC_BASE_AMOUNT_P.GetCommandParamValue("O_AMOUNT");
            }
            catch
            {
                GL_AMOUNT.EditValue = Math.Round(iString.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue, 0) *
                                                iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue, 0), 0);
            }
            Init_DR_CR_Amount();    // 차대금액 생성 //
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
        } 

        private bool Init_Exchange_Profit_Loss(int pCurrent_Row_Index)
        {
            int vIDX_ACCOUNT_DR_CR = igrSLIP_LINE.GetColumnToIndex("ACCOUNT_DR_CR");
            int vIDX_GL_CURRENCY_AMOUNT = igrSLIP_LINE.GetColumnToIndex("GL_CURRENCY_AMOUNT");
            int vIDX_GL_AMOUNT = igrSLIP_LINE.GetColumnToIndex("GL_AMOUNT");
            int vIDX_CURRENCY_CODE = igrSLIP_LINE.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_EXCHANGE_RATE = igrSLIP_LINE.GetColumnToIndex("EXCHANGE_RATE");

            decimal vDR_Curr_GL_Amount = 0;
            decimal vCR_Curr_GL_Amount = 0;
            decimal vDR_GL_Amount = 0;
            decimal vCR_GL_Amount = 0;
            decimal vExchange_Profit_Loss_Amount = 0;

            string vCURRENCY_CODE = iString.ISNull(mCurrency_Code);
            decimal vExchange_Rate = 0;

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
                //환율//
                if (iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, vIDX_EXCHANGE_RATE), 0) != 0)
                {
                    vExchange_Rate = iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, vIDX_EXCHANGE_RATE), 0);
                }
            }
            vExchange_Profit_Loss_Amount = vDR_GL_Amount - vCR_GL_Amount;
            if (vDR_Curr_GL_Amount + vCR_Curr_GL_Amount == 0 || vExchange_Profit_Loss_Amount == 0)
            {
                return true;
            }

            bool mExchange_Profit_Loss = false;
            object vAccount_DR_CR;
            object vAccount_DR_CR_Name;
            object vAccount_ID;
            object vAccount_Code;
            object vAccount_Desc;
            if (vExchange_Profit_Loss_Amount != Convert.ToDecimal(0))
            {
                idcEXCHANGE_PROFIT_LOSS.SetCommandParamValue("W_CONVERSION_AMOUNT", vExchange_Profit_Loss_Amount);
                idcEXCHANGE_PROFIT_LOSS.ExecuteNonQuery();
                vAccount_ID = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_ID");
                vAccount_Code = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_CODE");
                vAccount_Desc = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_DESC");
                vAccount_DR_CR = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_DR_CR");
                vAccount_DR_CR_Name = idcEXCHANGE_PROFIT_LOSS.GetCommandParamValue("O_ACCOUNT_DR_CR_NAME");

                // LINE 추가.
                idaSLIP_LINE.AddUnder();
                InsertSlipLine();

                //Set_Slip_Line_Seq();    //LINE SEQ 채번//               
                ACCOUNT_DR_CR.EditValue = vAccount_DR_CR;
                ACCOUNT_DR_CR_NAME.EditValue = vAccount_DR_CR_Name;
                ACCOUNT_CONTROL_ID.EditValue = vAccount_ID;
                ACCOUNT_CODE.EditValue = vAccount_Code;
                ACCOUNT_DESC.EditValue = vAccount_Desc;
                CURRENCY_CODE.EditValue = vCURRENCY_CODE;
                CURRENCY_DESC.EditValue = vCURRENCY_CODE;
                EXCHANGE_RATE.EditValue = vExchange_Rate;
                GL_CURRENCY_AMOUNT.EditValue = 0;
                GL_AMOUNT.EditValue = Math.Abs(iString.ISDecimaltoZero(vExchange_Profit_Loss_Amount));

                //참고항목 동기화.                
                Set_Control_Item_Prompt(idaSLIP_LINE.CurrentRow.RowState);
                Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);

                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                mExchange_Profit_Loss = true;
            }
            return mExchange_Profit_Loss;
        }

        private bool Init_Offset_Account(int pCurrent_Row_Index)
        { 
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
            foreach (System.Data.DataRow vRow in idaSLIP_LINE.CurrentRows)
            {
                if (vRow.RowState == DataRowState.Deleted)
                {

                }
                else
                {
                    if (iString.ISNull(vRow["ACCOUNT_DR_CR"]) == "1")
                    {
                        vDR_Curr_GL_Amount = vDR_Curr_GL_Amount + iString.ISDecimaltoZero(vRow["GL_CURRENCY_AMOUNT"]);
                        vDR_GL_Amount = vDR_GL_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                    }
                    else
                    {
                        vCR_Curr_GL_Amount = vCR_Curr_GL_Amount + iString.ISDecimaltoZero(vRow["GL_CURRENCY_AMOUNT"]);
                        vCR_GL_Amount = vCR_GL_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                    }

                    //통화//
                    if (iString.ISNull(vRow["CURRENCY_CODE"]) != mCurrency_Code)
                    {
                        vCURRENCY_CODE = iString.ISNull(vRow["CURRENCY_CODE"]);
                    }
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
            idaSLIP_LINE.AddUnder();
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
            GL_CURRENCY_AMOUNT.EditValue = vOffset_Curr_Amount;
            GL_AMOUNT.EditValue = vOffset_Amount;

            //참고항목 동기화.                
            Set_Control_Item_Prompt(idaSLIP_LINE.CurrentRow.RowState);
            Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);

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
                if (idaSLIP_LINE.CurrentRowPosition() != vIDX_ROW_CURR)
                {
                    return;
                }

                int vIDX_COL_GL_AMOUNT = igrSLIP_LINE.GetColumnToIndex("GL_AMOUNT");
                int vIDX_COL_DR = igrSLIP_LINE.GetColumnToIndex("DR_AMOUNT");
                int vIDX_COL_CR = igrSLIP_LINE.GetColumnToIndex("CR_AMOUNT");

                if (iString.ISNull(idaSLIP_LINE.CurrentRow["ACCOUNT_DR_CR"], "1") == "1".ToString())
                {
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, igrSLIP_LINE.GetCellValue(vIDX_ROW_CURR, vIDX_COL_GL_AMOUNT));
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, 0);
                }
                else if (iString.ISNull(idaSLIP_LINE.CurrentRow["ACCOUNT_DR_CR"], "1") == "2".ToString())
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
            decimal vCurr_DR_Amount = Convert.ToInt32(0);
            decimal vCurr_CR_Amount = Convert.ToInt32(0);

            foreach (DataRow vRow in idaSLIP_LINE.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Deleted)
                {
                    if (iString.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "1".ToString())
                    {
                        vDR_Amount = vDR_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                        vCurr_DR_Amount = vCurr_DR_Amount + iString.ISDecimaltoZero(vRow["GL_CURRENCY_AMOUNT"]);
                    }
                    else if (iString.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "2".ToString())
                    {
                        vCR_Amount = vCR_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                        vCurr_CR_Amount = vCurr_CR_Amount + iString.ISDecimaltoZero(vRow["GL_CURRENCY_AMOUNT"]);
                    }
                }
            }
            TOTAL_DR_AMOUNT.EditValue = iString.ISDecimaltoZero(vDR_Amount);
            TOTAL_CR_AMOUNT.EditValue = iString.ISDecimaltoZero(vCR_Amount);
            MARGIN_AMOUNT.EditValue = -(System.Math.Abs(iString.ISDecimaltoZero(vDR_Amount) - iString.ISDecimaltoZero(vCR_Amount)));

            CURR_TOTAL_DR_AMOUNT.EditValue = iString.ISDecimaltoZero(vCurr_DR_Amount);
            CURR_TOTAL_CR_AMOUNT.EditValue = iString.ISDecimaltoZero(vCurr_CR_Amount);
            CURR_MARGIN_AMOUNT.EditValue = -(System.Math.Abs(iString.ISDecimaltoZero(vCurr_DR_Amount) - iString.ISDecimaltoZero(vCurr_CR_Amount))); 
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

                GL_AMOUNT.ReadOnly = true;
                GL_AMOUNT.Insertable = false;
                GL_AMOUNT.Updatable = false;
                GL_AMOUNT.TabStop = false;
                GL_AMOUNT.Refresh(); 

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
                    MANAGEMENT1.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("MANAGEMENT1_DECIMAL_DIGITS"), 0);
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
                    MANAGEMENT2.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("MANAGEMENT2_DECIMAL_DIGITS"), 0);
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
                    REFER1.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER1_DECIMAL_DIGITS"), 0);
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
                    REFER2.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER2_DECIMAL_DIGITS"), 0);
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
                    REFER3.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER3_DECIMAL_DIGITS"), 0);
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
                    REFER4.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER4_DECIMAL_DIGITS"), 0);
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
                    REFER5.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER5_DECIMAL_DIGITS"), 0);
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
                    REFER6.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER6_DECIMAL_DIGITS"), 0);
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
                    REFER7.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER7_DECIMAL_DIGITS"), 0);
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
                    REFER8.NumberDecimalDigits = iString.ISNumtoZero(igrSLIP_LINE.GetCellValue("REFER8_DECIMAL_DIGITS"), 0);
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
            int mPreviousRowPosition = idaSLIP_LINE.CurrentRowPosition() - 1;
            object mPrevious_Code;
            object mPrevious_Name;
            string mData_Type;
            string mLookup_Type;

            if (mPreviousRowPosition > -1
                && iString.ISNull(REMARK.EditValue) == string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"]) != string.Empty)
            {//REMARK.
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"];
                REMARK.EditValue = mPrevious_Name;
            }

            //1
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT1_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT1.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT1.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_LOOKUP_TYPE"]))
            {//MANAGEMENT1_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_DESC"];

                MANAGEMENT1.EditValue = mPrevious_Code;
                MANAGEMENT1_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("MANAGEMENT1", mPrevious_Code, null, null, null);
            }
            //2
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT2_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT2.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT2.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_LOOKUP_TYPE"]))
            {//MANAGEMENT2_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_DESC"];

                MANAGEMENT2.EditValue = mPrevious_Code;
                MANAGEMENT2_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("MANAGEMENT2", mPrevious_Code, null, null, null);
            }
            //3
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER1_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER1_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER1.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER1.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_LOOKUP_TYPE"]))
            {//REFER1_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_DESC"];

                REFER1.EditValue = mPrevious_Code;
                REFER1_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER1", mPrevious_Code, null, null, null);
            }
            //4
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER2_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER2_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER2.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER2.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_LOOKUP_TYPE"]))
            {//REFER2_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_DESC"];

                REFER2.EditValue = mPrevious_Code;
                REFER2_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER2", mPrevious_Code, null, null, null);
            }
            //5
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER3_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER3_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER3.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER3.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_LOOKUP_TYPE"]))
            {//REFER3_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_DESC"];

                REFER3.EditValue = mPrevious_Code;
                REFER3_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER3", mPrevious_Code, null, null, null);
            }
            //6
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER4_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER4_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER4.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER4.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_LOOKUP_TYPE"]))
            {//REFER4_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_DESC"];

                REFER4.EditValue = mPrevious_Code;
                REFER4_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER4", mPrevious_Code, null, null, null);
            }
            //7
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER5_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER5_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER5.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER5.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_LOOKUP_TYPE"]))
            {//REFER5_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_DESC"];

                REFER5.EditValue = mPrevious_Code;
                REFER5_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER5", mPrevious_Code, null, null, null);
            }
            //8
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER6_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER6_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER6.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER6.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_LOOKUP_TYPE"]))
            {//REFER6_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_DESC"];

                REFER6.EditValue = mPrevious_Code;
                REFER6_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER6", mPrevious_Code, null, null, null);
            }
            //9
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER7_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER7_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER7.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER7.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_LOOKUP_TYPE"]))
            {//REFER7_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_DESC"];

                REFER7.EditValue = mPrevious_Code;
                REFER7_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER7", mPrevious_Code, null, null, null);
            }
            //10
            mData_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER8_DATA_TYPE"]);
            mLookup_Type = iString.ISNull(idaSLIP_LINE.CurrentRow["REFER8_LOOKUP_TYPE"]);
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER8.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER8.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue);;
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_LOOKUP_TYPE"]))
            {//REFER8_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_DESC"];

                REFER8.EditValue = mPrevious_Code;
                REFER8_DESC.EditValue = mPrevious_Name;

                //기본값 설정.
                Set_Ref_Management("REFER8", mPrevious_Code, null, null, null);
            }
        }
         
        private void Init_Currency_Code(string pInit_YN)
        {
            //if (iString.ISNull(idaSLIP_LINE.CurrentRow["CURRENCY_ENABLED_FLAG"], "N") == "Y")
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
        
        private void Init_Exchange_Rate(object pCurrency_Code)
        {
            if (iString.ISNull(pCurrency_Code) != string.Empty)
            {
                Init_Currency_Amount();
                if (iString.ISNull(pCurrency_Code) != mCurrency_Code)
                {
                    idcEXCHANGE_RATE.ExecuteNonQuery();
                    EXCHANGE_RATE.EditValue = idcEXCHANGE_RATE.GetCommandParamValue("O_EXCHANGE_RATE");

                    Init_GL_Amount();
                    EXCHANGE_RATE.Focus();
                }
            }
        }

        private void Init_Currency_Amount()
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) == string.Empty || iString.ISNull(CURRENCY_CODE.EditValue) == mCurrency_Code)
            {
                if (iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue) != Convert.ToDecimal(0))
                {
                    EXCHANGE_RATE.EditValue = null;
                }
                if (iString.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue) != Convert.ToDecimal(0))
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
                if (iString.ISNull(idaSLIP_LINE.CurrentRow["REF_SLIP_FLAG"]) != string.Empty)
                {
                    EXCHANGE_RATE.ReadOnly = true;
                    EXCHANGE_RATE.Insertable = false;
                    EXCHANGE_RATE.Updatable = false;
                    EXCHANGE_RATE.TabStop = false;

                    //원전표인 경우 금액수정 불가//
                    if (iString.ISNull(idaSLIP_LINE.CurrentRow["REF_SLIP_FLAG"]) == "S")
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
            object mVAT_ENABLED_FLAG = idaSLIP_LINE.CurrentRow["VAT_ENABLED_FLAG"];
            if (iString.ISNull(mVAT_ENABLED_FLAG, "N") != "Y")
            {
                return;
            }

            IDC_GET_ACCOUNT_DEFAULT_VALUE.SetCommandParamValue("W_ACCOUNT_TYPE", "DEFAULT_VAT_RATE");
            IDC_GET_ACCOUNT_DEFAULT_VALUE.ExecuteNonQuery();
            decimal vVAT_RATE = iString.ISDecimaltoZero(IDC_GET_ACCOUNT_DEFAULT_VALUE.GetCommandParamValue("O_VAT_RATE"));

            decimal mGL_AMOUNT = iString.ISDecimaltoZero(GL_AMOUNT.EditValue);
            IDC_GET_VAT_AMT_P.SetCommandParamValue("W_SUPPLY_AMT", mGL_AMOUNT);
            IDC_GET_VAT_AMT_P.SetCommandParamValue("W_VAT_RATE", vVAT_RATE);
            IDC_GET_VAT_AMT_P.ExecuteNonQuery();
            decimal mSUPPLY_AMOUNT = iString.ISDecimaltoZero(IDC_GET_VAT_AMT_P.GetCommandParamValue("O_VAT_AMT"));      //공급가액 설정. 
            Set_Management_Value("SUPPLY_AMOUNT", mSUPPLY_AMOUNT, null);
            Set_Management_Value("VAT_AMOUNT", mGL_AMOUNT, null);
        }

        //예산부서 동기화
        private void Init_Budget_Dept()
        {
            int mPreviousRowPosition = idaSLIP_LINE.CurrentRowPosition() - 1;
            object mPrevious_ID;
            object mPrevious_Code;
            object mPrevious_Name;

            if (mPreviousRowPosition > -1
                && iString.ISNull(BUDGET_DEPT_ID_L.EditValue) == string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_ID"]) != string.Empty)
            {//budget dept
                mPrevious_ID = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_ID"];
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_CODE"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["BUDGET_DEPT_NAME"];

                BUDGET_DEPT_NAME_L.EditValue = mPrevious_Name;
                BUDGET_DEPT_CODE_L.EditValue = mPrevious_Code;
                BUDGET_DEPT_ID_L.EditValue = mPrevious_ID;
            }
            else
            {
                BUDGET_DEPT_NAME_L.EditValue = BUDGET_DEPT_NAME.EditValue;
                BUDGET_DEPT_CODE_L.EditValue = BUDGET_DEPT_CODE.EditValue;
                BUDGET_DEPT_ID_L.EditValue = BUDGET_DEPT_ID.EditValue;
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
                mSUPPLY_AMOUNT = mSUPPLY_AMOUNT + iString.ISDecimaltoZero(igrDPR_ASSET.GetCellValue(r, mIDX_SUPPLY_AMOUNT));
                mVAT_AMOUNT = mVAT_AMOUNT + iString.ISDecimaltoZero(igrDPR_ASSET.GetCellValue(r, mIDX_VAT_AMOUNT));
                mCOUNT = mCOUNT + iString.ISDecimaltoZero(igrDPR_ASSET.GetCellValue(r, mIDX_COUNT));

                if ((mSUPPLY_AMOUNT + mVAT_AMOUNT) != 0 && iString.ISNull(igrDPR_ASSET.GetCellValue(r, mIDX_ITEM_CONTENTS)) == string.Empty)
                {//공급가액, 부가세 등록했는데 품목 등록 안함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10523"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (mSUPPLY_AMOUNT == 0 && iString.ISNull(igrDPR_ASSET.GetCellValue(r, mIDX_ITEM_CONTENTS)) != string.Empty)
                {//공급가액 등록 않했는데 품목 등록함 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (mVAT_AMOUNT == 0 && iString.ISNull(igrDPR_ASSET.GetCellValue(r, mIDX_ITEM_CONTENTS)) != string.Empty)
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

        private void SLIP_REQ_OK()
        {
            //결재승인단계 필수인경우 결재승인단계 설정 여부 검증//
            if (mAPPROVAL_YN == "Y")
            {
                bool vAPPR_STEP = false;
                foreach (DataRow vROW in IDA_APPROVAL_PERSON.CurrentRows)
                {
                    vAPPR_STEP = true;
                }
                if (vAPPR_STEP == false)
                {
                    if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10223"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }
                }
            }

            if (iString.ISNull(HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10118"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(string.Format("{0} :: {1}", BTN_REQ_OK.ButtonText, isMessageAdapter1.ReturnText("EAPP_10067")), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_APPROVAL_REQUEST_OK.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SET_APPROVAL_REQUEST_OK.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_APPROVAL_REQUEST_OK.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_APPROVAL_REQUEST_OK.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_APPROVAL_REQUEST_OK.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            Search_DB_DETAIL(HEADER_ID.EditValue);

            //인쇄여부 묻기//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10146"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                XLPrinting_Main("PRINT");
            }
        }

        private void SLIP_REQ_CANCEL()
        {
            if (iString.ISNull(HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10118"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(string.Format("{0} :: {1}", BTN_REQ_CANCEL.ButtonText, isMessageAdapter1.ReturnText("EAPP_10067")), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_APPROVAL_REQUEST_CANCEL.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SET_APPROVAL_REQUEST_CANCEL.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_APPROVAL_REQUEST_CANCEL.GetCommandParamValue("O_MESSAGE"));

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

            Search_DB_DETAIL(HEADER_ID.EditValue);
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
            saveFileDialog1.DefaultExt = "xlsx";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
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
                        vFileTransferAdv.UseBinary = true;
                        vFileTransferAdv.KeepAlive = false;
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive != "Y")
                        {
                            vFileTransferAdv.UsePassive = false;
                        }
                        else
                        {
                            vFileTransferAdv.UsePassive = true;
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


        private void AssmblyRun_Attachment(object pAssembly_ID, object pSLIP_Date, object pSLIP_Num)
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

#if DEBUG
            vASSEMBLY_FILE_NAME = "FCMF0228.dll";
#endif
             
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
                        vFileTransferAdv.UseBinary = true;
                        vFileTransferAdv.KeepAlive = false;
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive != "Y")
                        {
                            vFileTransferAdv.UsePassive = false;
                        }
                        else
                        {
                            vFileTransferAdv.UsePassive = true;
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
                            vParam[2] = "SLIP_BUDGET";     //카테고리
                            vParam[3] = pSLIP_Date;     //전표일자
                            vParam[4] = pSLIP_Num;      //전표번호
                            vParam[5] = "N";            //읽기 전용 여부

                            object vCreateInstance = Activator.CreateInstance(vType, vParam);
                            Office2007Form vForm = vCreateInstance as Office2007Form;
                            Point vPoint = new Point(30, 30);
                            vForm.Location = vPoint;
                            vForm.StartPosition = FormStartPosition.CenterParent;
                            vForm.Text = string.Format("{0}[{1}] - {2}", vASSEMBLY_NAME, vASSEMBLY_ID, vCurrAssemblyFileVersion);

                            vForm.Show();
                           
                        }
                        else
                        {
                            MessageBoxAdv.Show("Form Namespace Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch(Exception Ex)
                    {
                        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive != "Y")
                    {
                        vFileTransferAdv.UsePassive = false;
                    }
                    else
                    {
                        vFileTransferAdv.UsePassive = true;
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

        private void XLPrinting_Main(string pOutput_Type)
        {
            object vSlip_Header_id;
            object vSlip_Date;
            object vSlip_Num;

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            if (itbSLIP.SelectedTab.TabIndex == 2)
            {
                vSlip_Header_id = HEADER_ID.EditValue;
                vSlip_Date = SLIP_DATE.EditValue;
                vSlip_Num = SLIP_NUM.EditValue;
            }
            else
            {
                vSlip_Header_id = igrSLIP_BUDGET_LIST.GetCellValue("HEADER_ID");
                vSlip_Date = igrSLIP_BUDGET_LIST.GetCellValue("SLIP_DATE");
                vSlip_Num = igrSLIP_BUDGET_LIST.GetCellValue("SLIP_NUM");
            }
            AssmblyRun_Manual("FCMF0212", vSlip_Header_id, vSlip_Date, vSlip_Num);
            //IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", SLIP_DATE.EditValue);
            //IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0206");
            //IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            //string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
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

            SLIP_DATE.Focus();
        }

        private void XLPrinting(string pOutput_Type)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            int vCountRowGrid = igrSLIP_BUDGET_LIST.RowCount;
            if ((itbSLIP.SelectedIndex == 0 && vCountRowGrid > 0) ||
                (itbSLIP.SelectedIndex == 1 && iString.ISNull(HEADER_ID.EditValue) != string.Empty))
            {
                vMessageText = string.Format("Printing Starting", vPageTotal);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0206_001.xlsx";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        object vObject;
                        int vCountRow = 0;
                        int vRow = igrSLIP_BUDGET_LIST.RowIndex;

                        //인쇄일자 
                        IDC_GET_DATE.ExecuteNonQuery();
                        object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");


                        if (itbSLIP.SelectedTab.TabIndex == 2)
                        {
                            xlPrinting.HeaderWrite(idaSLIP_HEADER, vLOCAL_DATE);
                            vObject = HEADER_ID.EditValue;
                        }
                        else
                        {
                            xlPrinting.HeaderWrite(igrSLIP_BUDGET_LIST, vRow, vLOCAL_DATE);
                            vObject = igrSLIP_BUDGET_LIST.GetCellValue("HEADER_ID");
                        }
                        idaPRINT_SLIP_LINE.SetSelectParamValue("W_HEADER_ID", vObject);
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

                    System.Windows.Forms.Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }
        
        private void XLPrinting_BSK(string pOutput_Type)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            int vCountRowGrid = igrSLIP_BUDGET_LIST.RowCount;
            if ((itbSLIP.SelectedIndex == 0 && vCountRowGrid > 0) ||
                (itbSLIP.SelectedIndex == 1 && iString.ISNull(HEADER_ID.EditValue) != string.Empty))
            {
                vMessageText = string.Format("Printing Starting", vPageTotal);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0206_011.xlsx";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        object vObject;
                        int vCountRow = 0;
                        int vRow = igrSLIP_BUDGET_LIST.RowIndex;

                        //인쇄일자 
                        IDC_GET_DATE.ExecuteNonQuery();
                        object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                        //회계법인명.
                        IDC_GET_COMPANY_NAME_P.ExecuteNonQuery();
                        object vSOB_DESC = IDC_GET_COMPANY_NAME_P.GetCommandParamValue("O_SOB_DESC");

                        if (itbSLIP.SelectedTab.TabIndex == 2)
                        {
                            xlPrinting.HeaderWrite_BSK(idaSLIP_HEADER, vSOB_DESC, vLOCAL_DATE);
                            vObject = HEADER_ID.EditValue;
                        }
                        else
                        {
                            xlPrinting.HeaderWrite_BSK(igrSLIP_BUDGET_LIST, vRow, vSOB_DESC, vLOCAL_DATE);
                            vObject = igrSLIP_BUDGET_LIST.GetCellValue("HEADER_ID");
                        }
                        idaPRINT_SLIP_LINE.SetSelectParamValue("W_HEADER_ID", vObject);
                        idaPRINT_SLIP_LINE.Fill();

                        vCountRow = idaPRINT_SLIP_LINE.CurrentRows.Count;
                        if (vCountRow > 0)
                        {
                            vPageNumber = xlPrinting.LineWrite_BSK(idaPRINT_SLIP_LINE);
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

                    System.Windows.Forms.Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
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
                    if (idaSLIP_LINE.IsFocused)
                    {
                        idaSLIP_LINE.AddUnder();
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
                            idaSLIP_HEADER.SetSelectParamValue("W_HEADER_ID", 0);
                            idaSLIP_HEADER.Fill();

                            idaSLIP_HEADER.AddUnder();
                            idaSLIP_LINE.AddUnder();
                            InsertSlipHeader();
                            InsertSlipLine();

                            SLIP_DATE.Focus();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    gBTN_Req_YN = "N";
                    if (idaSLIP_BUDGET_LIST.IsFocused)
                    {
                        idaSLIP_BUDGET_LIST.Update();
                    }
                    else
                    {
                        Save_Data();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    SLIP_QUERY_STATUS.EditValue = "QUERY";
                    gBTN_Req_YN = "F";
                    if (idaSLIP_BUDGET_LIST.IsFocused)
                    {
                        idaSLIP_BUDGET_LIST.Cancel();
                    }
                    else if (idaSLIP_HEADER.IsFocused)
                    {
                        idaSLIP_LINE.Cancel();
                        idaSLIP_HEADER.Cancel();
                    }
                    else if (idaSLIP_LINE.IsFocused)
                    {
                        idaSLIP_LINE.Cancel();
                        Init_Total_GL_Amount();  //합계 금액 재 계산 //
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    gBTN_Req_YN = "F";
                    if (idaSLIP_BUDGET_LIST.IsFocused)
                    {
                        idaSLIP_BUDGET_LIST.Delete();
                    }
                    else if (idaSLIP_HEADER.IsFocused)
                    {
                        if (Check_Sub_Panel() == false)
                        {
                            return;
                        }

                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_DELETE_SLIP.SetCommandParamValue("W_HEADER_ID", HEADER_ID.EditValue);
                        IDC_DELETE_SLIP.ExecuteNonQuery();
                        Search_DB();
                        Search_DB_DETAIL(HEADER_ID.EditValue); 
                    }
                    else if (idaSLIP_LINE.IsFocused)
                    {
                        idaSLIP_LINE.Delete();
                        if (idaSLIP_LINE.CurrentRows.Count == 0)
                        {
                            idaSLIP_LINE.Cancel();
                        }
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

        private bool Save_Data()
        {
            SLIP_NUM.Focus();
            ACCOUNT_CODE.Focus();

            Init_DR_CR_Amount();    // 차대금액 생성 //
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //

            if (iString.ISDecimaltoZero(TOTAL_DR_AMOUNT.EditValue) != iString.ISDecimaltoZero(TOTAL_CR_AMOUNT.EditValue))
            {// 차대금액 일치 여부 체크.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10134"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (Check_Sub_Panel() == false)
            {
                return false;
            }

            //delete temp data
            Delete_Balance_Remain_TP();

            idaSLIP_HEADER.Update(); 
            return true;
        }

#endregion;

#region ----- Form Event ----- 
        
        private void FCMF0206_Load(object sender, EventArgs e)
        {
            irbCONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            CONFIRM_STATUS_0.EditValue = irbCONFIRM_ALL.RadioCheckedString;
            RB_DATE.CheckedState = ISUtil.Enum.CheckedState.Checked;
                        
            if (mAuto_Search_Flag == "N")
            {
                SLIP_DATE_FR_0.EditValue = iDate.ISDate_Add(iDate.ISGetDate(), -20);
                SLIP_DATE_TO_0.EditValue = iDate.ISGetDate();
            }

            // 회계장부 정보 설정.
            GetAccountBook();

            idaSLIP_BUDGET_LIST.FillSchema();
            idaSLIP_HEADER.FillSchema();
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            
            //서브판넬 
            Init_Sub_Panel(false, "ALL");

            //전표 복사 버튼 맨 앞으로 가져오기
            //btnGET_BALANCE_STATEMENT.BringToFront();
            BTN_COPY_SLIP.BringToFront();
            btnGET_BALANCE_STATEMENT.BringToFront();
            BUDGET_REMAIN_AMOUNT.BringToFront();
            REF_SLIP_FLAG.BringToFront(); 

            // 콤퍼넌트 동기화.
            //Init_Currency_Code();
            ibtSUB_FORM.Visible = false;

            idaSLIP_HEADER.FillSchema();
        }

        private void FCMF0206_Shown(object sender, EventArgs e)
        {
            
            Application.DoEvents();
            if (mAuto_Search_Flag == "L")
            {
                Search_DB(); 
            }
            else if (mAuto_Search_Flag == "Y")
            {
                Search_DB();
                if (igrSLIP_BUDGET_LIST.RowCount > 0)
                {
                    Search_DB_DETAIL(igrSLIP_BUDGET_LIST.GetCellValue("SLIP_HEADER_ID"));
                }
            }
            else
            {
                itbSLIP.SelectedIndex = 1;
                itbSLIP.SelectedTab.Focus();
            }
        }

        private void igrSLIP_LIST_CellDoubleClick(object pSender)
        {
            if (igrSLIP_BUDGET_LIST.RowCount > 0)
            {
                Search_DB_DETAIL(igrSLIP_BUDGET_LIST.GetCellValue("HEADER_ID"));
            }
        }

        private void H_REMARK_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(REMARK.EditValue) == string.Empty)
            {
                REMARK.EditValue = H_REMARK.EditValue;
            }
        }

        private void CURRENCY_DESC_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_Currency_Amount();
        }

        private void irbCONFIRM_Status_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;

            if (iStatus.Checked == true)
            {
                CONFIRM_STATUS_0.EditValue = iStatus.RadioCheckedString;
            }
        }

        private void RB_DATE_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;

            if (vRadio.Checked == true)
            {
                SORT_TYPE_0.EditValue = vRadio.RadioCheckedString;
            }
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
                GL_CURRENCY_AMOUNT.Focus();
                return;
            }
                         
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
                    mISSUE_DATE = SLIP_DATE.EditValue;
                }
                object mDUE_DATE = Get_Management_Value("DUE_DATE");
                if (iString.ISNull(mDUE_DATE) == string.Empty)
                {
                    mDUE_DATE = Get_Management_Value("TR_EXPIRATION_DATE");
                }
                object mDEPT_ID = DEPT_ID.EditValue;
                object mDEPT_NAME = DEPT_NAME.EditValue;

                FCMF0206_BILL vFCMF0206_BILL = new FCMF0206_BILL(isAppInterfaceAdv1.AppInterface, mDEPT_ID, mDEPT_NAME
                                                                    , mBILL_CLASS, mBILL_NUM, mBILL_AMOUNT
                                                                    , mVENDOR_CODE, mBANK_CODE
                                                                    , mVAT_ISSUE_DATE, mISSUE_DATE, mDUE_DATE
                                                                    , PERSON_ID.EditValue, PERSON_NAME.EditValue);
                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0206_BILL, isAppInterfaceAdv1.AppInterface);
                dlgResult = vFCMF0206_BILL.ShowDialog();
                if (dlgResult == DialogResult.OK)
                {
                    //어음금액
                    GL_AMOUNT.EditValue = vFCMF0206_BILL.Get_BILL_AMOUNT;
                    //거래처.
                    Set_Management_Value("CUSTOMER", vFCMF0206_BILL.Get_VENDOR_CODE, vFCMF0206_BILL.Get_VENDOR_NAME);
                    //은행
                    Set_Management_Value("BANK", vFCMF0206_BILL.Get_BANK_CODE, vFCMF0206_BILL.Get_BANK_NAME);
                    //세금계산서발행일
                    Set_Management_Value("VAT_ISSUE_DATE", vFCMF0206_BILL.Get_VAT_ISSUE_DATE, null);
                    //발행일자
                    Set_Management_Value("ISSUE_DATE", vFCMF0206_BILL.Get_ISSUE_DATE, null);
                    //만기일자
                    Set_Management_Value("DUE_DATE", vFCMF0206_BILL.Get_DUE_DATE, null);
                    //만기일자
                    Set_Management_Value("TR_EXPIRATION_DATE", vFCMF0206_BILL.Get_DUE_DATE, null);  
                    //어음번호.
                    Set_Management_Value("RECEIVABLE_BILL", vFCMF0206_BILL.Get_BILL_NUM, String.Format("{0:###,###,###,###,###,###}", vFCMF0206_BILL.Get_BILL_AMOUNT));
                     
                    Init_DR_CR_Amount();    // 차대금액 생성 //
                    Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                }
                vFCMF0206_BILL.Dispose();
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
                    mISSUE_DATE = SLIP_DATE.EditValue;
                }
                object mDUE_DATE = Get_Management_Value("DUE_DATE");
                if (iString.ISNull(mDUE_DATE) == string.Empty)
                {
                    mDUE_DATE = Get_Management_Value("TR_EXPIRATION_DATE");
                }
                object mDEPT_ID = DEPT_ID.EditValue;
                object mDEPT_NAME = DEPT_NAME.EditValue;

                FCMF0206_BILL vFCMF0206_BILL = new FCMF0206_BILL(isAppInterfaceAdv1.AppInterface, mDEPT_ID, mDEPT_NAME
                                                                    , mBILL_CLASS, mBILL_NUM, mBILL_AMOUNT
                                                                    , mVENDOR_CODE, mBANK_CODE
                                                                    , mVAT_ISSUE_DATE, mISSUE_DATE, mDUE_DATE
                                                                    , PERSON_ID.EditValue, PERSON_NAME.EditValue);

                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0206_BILL, isAppInterfaceAdv1.AppInterface);
                dlgResult = vFCMF0206_BILL.ShowDialog();
                if (dlgResult == DialogResult.OK)
                {
                    //어음금액
                    GL_AMOUNT.EditValue = vFCMF0206_BILL.Get_BILL_AMOUNT;
                    //거래처.
                    Set_Management_Value("CUSTOMER", vFCMF0206_BILL.Get_VENDOR_CODE, vFCMF0206_BILL.Get_VENDOR_NAME);
                    //은행
                    Set_Management_Value("BANK", vFCMF0206_BILL.Get_BANK_CODE, vFCMF0206_BILL.Get_BANK_NAME);
                    //세금계산서발행일
                    Set_Management_Value("VAT_ISSUE_DATE", vFCMF0206_BILL.Get_VAT_ISSUE_DATE, null);
                    //발행일자
                    Set_Management_Value("ISSUE_DATE", vFCMF0206_BILL.Get_ISSUE_DATE, null);
                    //만기일자
                    Set_Management_Value("DUE_DATE", vFCMF0206_BILL.Get_DUE_DATE, null);
                    //만기일자
                    Set_Management_Value("TR_EXPIRATION_DATE", vFCMF0206_BILL.Get_DUE_DATE, null);  
                    //어음번호.
                    Set_Management_Value("PAYABLE_BILL", vFCMF0206_BILL.Get_BILL_NUM, String.Format("{0:###,###,###,###,###,###}", vFCMF0206_BILL.Get_BILL_AMOUNT));
                    
                    Init_DR_CR_Amount();    // 차대금액 생성 //
                    Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                }
                vFCMF0206_BILL.Dispose();
            }
            else if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "AP_VAT".ToString())
            {
                S_SUPPLY_AMOUNT.EditValue = Get_Management_Value("SUPPLY_AMOUNT");   //공급가액 설정.
                S_VAT_AMOUNT.EditValue = Get_Management_Value("VAT_AMOUNT");      //세액 설정.

                //서브판넬 
                Init_Sub_Panel(true, "AP_VAT");
            }
            else if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "DEAL".ToString())
            {//외화획득용 구매(공급) 확인서
                object mISSUE_NUM = MANAGEMENT2.EditValue;
                object mVENDOR_CODE = MANAGEMENT1.EditValue;
                object mBANK_CODE = REFER4.EditValue;
                object mISSUE_DATE = REFER5.EditValue;
                object mCURRENCY_CODE = CURRENCY_CODE.EditValue;

                FCMF0206_ITEM_DEAL vFCMF0206_ITEM_DEAL = new FCMF0206_ITEM_DEAL(isAppInterfaceAdv1.AppInterface, mISSUE_NUM, mCURRENCY_CODE
                                                                                , mVENDOR_CODE, mBANK_CODE, mISSUE_DATE);

                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0206_ITEM_DEAL, isAppInterfaceAdv1.AppInterface);
                dlgResult = vFCMF0206_ITEM_DEAL.ShowDialog();
                if (dlgResult == DialogResult.OK)
                {
                    //거래처.
                    MANAGEMENT1.EditValue = vFCMF0206_ITEM_DEAL.Get_VENDOR_CODE;
                    MANAGEMENT1_DESC.EditValue = vFCMF0206_ITEM_DEAL.Get_VENDOR_NAME;
                    //구매(공급)확인번호
                    MANAGEMENT2.EditValue = vFCMF0206_ITEM_DEAL.Get_ISSUE_NUM;
                    MANAGEMENT2_DESC.EditValue = DBNull.Value;

                    REFER4.EditValue = vFCMF0206_ITEM_DEAL.Get_BANK_CODE;
                    REFER4_DESC.EditValue = vFCMF0206_ITEM_DEAL.Get_BANK_NAME;

                    REFER5.EditValue = vFCMF0206_ITEM_DEAL.Get_ISSUE_DATE;                    
                }
                vFCMF0206_ITEM_DEAL.Dispose();
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void REF_SLIP_FLAG_DoubleClick(object pSender)
        {
            Search_REF_LIST();
        }

        private void REF_SLIP_FLAG_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_REF_LIST();
            }
        }

        private void V_REF_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "REF_LIST");
        }

        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (idaSLIP_LINE.CurrentRow != null && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Deleted)
            {
                Init_GL_Amount();
            }
        }

        private void GL_CURRENCY_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (idaSLIP_LINE.CurrentRow != null && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Deleted)
            {
                Init_GL_Amount();
            }
        }

        private void GL_AMOUNT_EditValueChanged(object pSender)
        {
            if (idaSLIP_LINE.CurrentRow != null && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Deleted)
            {
                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_VAT_Amount();
            }
        }

        private void GL_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
        }

        private void BTN_REQ_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            gBTN_Req_YN = "Y";
            if (Save_Data() == false)
            {
                return;
            }

            SLIP_REQ_OK();
        }

        private void BTN_REQ_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SLIP_REQ_CANCEL();
        }

        private void btnGET_BALANCE_STATEMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Get_Balance_Statement();
        }

        private void BTN_DIFF_EXCHANGE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Init_Exchange_Profit_Loss(igrSLIP_LINE.RowIndex);
        }

        private void BTN_OFFSET_ACCOUNT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Init_Offset_Account(igrSLIP_LINE.RowIndex);
        }

        private void BTN_COPY_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //전표 작성중이면 저장후 작업해야 함
            if (iString.ISNull(SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10128"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Check_SlipHeader_Added() == true)
            {
                return;
            }
            
            //서브판넬 
            C_OLD_SLIP_DATE.EditValue = SLIP_DATE.EditValue;
            C_OLD_SLIP_NUM.EditValue = SLIP_NUM.EditValue;
            C_OLD_HEADER_ID.EditValue = HEADER_ID.EditValue;

            C_NEW_SLIP_DATE.EditValue = iDate.ISGetDate();
            C_NEW_SLIP_NUM.EditValue = string.Empty;
            C_NEW_HEADER_ID.EditValue = DBNull.Value;

            Init_Sub_Panel(true, "COPY_SLIP");
        }
       
        private void C_BTN_SET_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10303"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_COPY_SLIP_BUDGET.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SET_COPY_SLIP_BUDGET.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_COPY_SLIP_BUDGET.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_COPY_SLIP_BUDGET.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            C_NEW_HEADER_ID.EditValue = IDC_SET_COPY_SLIP_BUDGET.GetCommandParamValue("O_NEW_HEADER_ID");
            C_NEW_SLIP_NUM.EditValue = IDC_SET_COPY_SLIP_BUDGET.GetCommandParamValue("O_NEW_SLIP_NUM");            
        }

        private void C_BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "COPY_SLIP");

            if (CB_NEW_SLIP_SEARCH_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                Search_DB_DETAIL(C_NEW_HEADER_ID.EditValue);
            }
        }

        private void S_BTN_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "AP_VAT".ToString())
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

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
            //서브판넬 
            Init_Sub_Panel(false, "AP_VAT");

            MANAGEMENT1.Focus();    //focus 이동 
        }

        private void ACCOUNT_CODE_KeyUp(object pSender, KeyEventArgs e)
        {
            if (btnGET_BALANCE_STATEMENT.Visible == true && ACCOUNT_CODE.ReadOnly == false && e.Control == true && e.KeyCode == Keys.F9)
            {
                Get_Balance_Statement();
            }
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
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_USE_DEPT_ID", DBNull.Value);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaSLIP_NUM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
        }

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SLIP_TYPE_TEMP_DOCU.SetLookupParamValue("P_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaSLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SLIP_TYPE_TEMP_DOCU.SetLookupParamValue("P_ENABLED_FLAG", "Y");
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

        private void ILA_BUDGET_DEPT_H_SelectedRowData(object pSender)
        {
            Init_SLIP_APPR();
        }

        private void ilaBUDGET_DEPT_L_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBUDGET_DEPT_L.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT_L.SetLookupParamValue("W_EFFECTIVE_DATE_FR", SLIP_DATE.EditValue);
            ildBUDGET_DEPT_L.SetLookupParamValue("W_EFFECTIVE_DATE_TO", SLIP_DATE.EditValue);  
        }

        private void ilaBUDGET_DEPT_L_SelectedRowData(object pSender)
        {
            Init_Default_Management("DEPT");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaACCOUNT_DR_CR_SelectedRowData(object pSender)
        {
            //전호수주석 : 관리항목 변경.
            //Set_Control_Item_Prompt(idaSLIP_LINE.CurrentRow.RowState);
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
            Set_Control_Item_Prompt(idaSLIP_LINE.CurrentRow.RowState);
            Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow); 
            Init_Set_Item_Need(idaSLIP_LINE.CurrentRow); 
            if (idaSLIP_LINE.CurrentRow.RowState != DataRowState.Modified)
            {
                Init_Default_Value();
            }
            Init_Default_Management("DEPT");
            Init_Default_Management("TAX_CODE");
            GetSubForm();

            //로직 수정전//
            //Init_Currency_Code("Y");
            //Set_Control_Item_Prompt();
            //Init_Control_Management_Value();
            //Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);
            //Init_Set_Item_Need(idaSLIP_LINE.CurrentRow);
            //Init_Default_Value();
            //Init_Dept();

            //GetSubForm();
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_SelectedRowData(object pSender)
        {
            Init_Exchange_Rate(CURRENCY_CODE.EditValue);
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_DEPT_ID", BUDGET_DEPT_ID.EditValue);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_USE_DEPT_ID", BUDGET_DEPT_ID_L.EditValue);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }


        private void ILA_SLIP_REMARK_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_TEMP_SLIP_REMARK.SetLookupParamValue("P_ENABLED_FLAG", "Y");
            ILD_TEMP_SLIP_REMARK.SetLookupParamValue("P_ENABLED_DATE", GL_DATE.EditValue);
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
            Set_Ref_Management("MANAGEMENT1", MANAGEMENT1.EditValue, null, null, GL_AMOUNT.EditValue);

            ////거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "DUE_DATE", MANAGEMENT1.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("MANAGEMENT1", "CUSTOMER", "PAYMENT_METHOD", MANAGEMENT1.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("MANAGEMENT1", "CREDIT_CARD", "DUE_DATE", MANAGEMENT1.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("MANAGEMENT1", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", MANAGEMENT1.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaMANAGEMENT2_SelectedRowData(object pSender)
        {// 관리항목2 선택시 적용.
            Init_SELECT_LOOKUP("MANAGEMENT2");

            //관리항목 동기화// 
            Set_Ref_Management("MANAGEMENT2", MANAGEMENT2.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("MANAGEMENT2", "CUSTOMER", "DUE_DATE", MANAGEMENT2.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("MANAGEMENT2", "CUSTOMER", "PAYMENT_METHOD", MANAGEMENT2.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("MANAGEMENT2", "CREDIT_CARD", "DUE_DATE", MANAGEMENT2.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("MANAGEMENT2", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", MANAGEMENT2.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER1_SelectedRowData(object pSender)
        {// 관리항목3 선택시 적용.
            Init_SELECT_LOOKUP("REFER1");

            //관리항목 동기화// 
            Set_Ref_Management("REFER1", REFER1.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER1", "CUSTOMER", "DUE_DATE", REFER1.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER1", "CUSTOMER", "PAYMENT_METHOD", REFER1.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER1", "CREDIT_CARD", "DUE_DATE", REFER1.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER1", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER1.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER2_SelectedRowData(object pSender)
        {// 관리항목4 선택시 적용.
            Init_SELECT_LOOKUP("REFER2");

            //관리항목 동기화// 
            Set_Ref_Management("REFER2", REFER2.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER2", "CUSTOMER", "DUE_DATE", REFER2.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER2", "CUSTOMER", "PAYMENT_METHOD", REFER2.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER2", "CREDIT_CARD", "DUE_DATE", REFER2.EditValue);
            ////공급가액 동기화//
            Set_Ref_Management_Value("REFER2", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER2.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER3_SelectedRowData(object pSender)
        {// 관리항목5 선택시 적용.
            Init_SELECT_LOOKUP("REFER3");

            //관리항목 동기화// 
            Set_Ref_Management("REFER3", REFER3.EditValue, null, null, GL_AMOUNT.EditValue);
            
            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER3", "CUSTOMER", "DUE_DATE", REFER3.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER3", "CUSTOMER", "PAYMENT_METHOD", REFER3.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER3", "CREDIT_CARD", "DUE_DATE", REFER3.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER3", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER3.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER4_SelectedRowData(object pSender)
        {// 관리항목6 선택시 적용.
            Init_SELECT_LOOKUP("REFER4");

            //관리항목 동기화// 
            Set_Ref_Management("REFER4", REFER4.EditValue, null, null, GL_AMOUNT.EditValue);
            
            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER4", "CUSTOMER", "DUE_DATE", REFER4.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER4", "CUSTOMER", "PAYMENT_METHOD", REFER4.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER4", "CREDIT_CARD", "DUE_DATE", REFER4.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER4", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER4.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER5_SelectedRowData(object pSender)
        {// 관리항목7 선택시 적용.
            Init_SELECT_LOOKUP("REFER5");

            //관리항목 동기화// 
            Set_Ref_Management("REFER5", REFER5.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER5", "CUSTOMER", "DUE_DATE", REFER5.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER5", "CUSTOMER", "PAYMENT_METHOD", REFER5.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER5", "CREDIT_CARD", "DUE_DATE", REFER5.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER5", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER5.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER6_SelectedRowData(object pSender)
        {// 관리항목8 선택시 적용.
            Init_SELECT_LOOKUP("REFER6");

            //관리항목 동기화// 
            Set_Ref_Management("REFER6", REFER6.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER6", "CUSTOMER", "DUE_DATE", REFER6.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER6", "CUSTOMER", "PAYMENT_METHOD", REFER6.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER6", "CREDIT_CARD", "DUE_DATE", REFER6.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER6", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER6.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER7_SelectedRowData(object pSender)
        {// 관리항목9 선택시 적용.
            Init_SELECT_LOOKUP("REFER7");

            //관리항목 동기화// 
            Set_Ref_Management("REFER7", REFER7.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER7", "CUSTOMER", "DUE_DATE", REFER7.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER7", "CUSTOMER", "PAYMENT_METHOD", REFER7.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER7", "CREDIT_CARD", "DUE_DATE", REFER7.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER7", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER7.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaREFER8_SelectedRowData(object pSender)
        {// 관리항목10 선택시 적용.
            Init_SELECT_LOOKUP("REFER8");

            //관리항목 동기화// 
            Set_Ref_Management("REFER8", REFER8.EditValue, null, null, GL_AMOUNT.EditValue);

            ////1.거래처 선택시 만기일자 있으면 만기일자 설정//
            //Set_Ref_Management_Value("REFER8", "CUSTOMER", "DUE_DATE", REFER8.EditValue);
            ////거래처 선택시 지급방법 설정//
            //Set_Ref_Management_Value("REFER8", "CUSTOMER", "PAYMENT_METHOD", REFER8.EditValue);
            ////신용카드 결재일자//
            //Set_Ref_Management_Value("REFER8", "CREDIT_CARD", "DUE_DATE", REFER8.EditValue);
            ////공급가액 동기화//
            //Set_Ref_Management_Value("REFER8", "VAT_TAX_TYPE", "SUPPLY_AMOUNT", REFER8.EditValue, null, null, GL_AMOUNT.EditValue);
        }

        private void ilaVAT_ASSET_GB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_ASSET_GB", "Y");
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
            if (iString.ISNull(e.Row["SLIP_DATE"]) == string.Empty)
            {// 기표일자.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10117"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            // 전표번호 채번//
            if (iString.ISNull(SLIP_NUM.EditValue) == string.Empty || iString.ISNull(e.Row["SLIP_DATE"]).Substring(0, 7) != iString.ISNull(e.Row["OLD_SLIP_DATE"], e.Row["SLIP_DATE"]).Substring(0, 7))
            {
                GetSlipNum();
            }
            else if (iString.ISNull(SLIP_TYPE.EditValue) != iString.ISNull(OLD_SLIP_TYPE.EditValue))
            {
                GetSlipNum();
            }

            if (iString.ISNull(e.Row["SLIP_NUM"]) == string.Empty)
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
            //if (e.Row.RowState != DataRowState.Added)
            //{
            //    if (e.Row["CONFIRM_YN"].ToString() == "Y".ToString())
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10448"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        e.Cancel = true;
            //        return;
            //    }
            //}
        }

        private void idaSLIP_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["SLIP_LINE_SEQ"]) == string.Empty)
            {// 예산부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10415"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {// 예산부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            //예산관리 계정에 대해서 예산부서 검증.
            if (iString.ISNull(e.Row["BUDGET_ENABLED_FLAG"]) == "Y" && iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10458"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                        
            //관리항목 검증//
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_SLIP_DATE", SLIP_DATE.EditValue);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", e.Row["ACCOUNT_CONTROL_ID"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_ACCOUNT_CODE", e.Row["ACCOUNT_CODE"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_ACCOUNT_DR_CR", e.Row["ACCOUNT_DR_CR"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_MANAGEMENT1", e.Row["MANAGEMENT1"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_MANAGEMENT2", e.Row["MANAGEMENT2"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER1", e.Row["REFER1"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER2", e.Row["REFER2"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER3", e.Row["REFER3"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER4", e.Row["REFER4"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER5", e.Row["REFER5"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER6", e.Row["REFER6"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER7", e.Row["REFER7"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_REFER8", e.Row["REFER8"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.SetCommandParamValue("P_UNLIQUIDATE_SLIP_LINE_ID", e.Row["UNLIQUIDATE_SLIP_LINE_ID"]);
            IDC_ACCOUNT_VALIDATE_DATA_P.ExecuteNonQuery();
            string vStatus = iString.ISNull(IDC_ACCOUNT_VALIDATE_DATA_P.GetCommandParamValue("O_STATUS"));
            string vMessage = iString.ISNull(IDC_ACCOUNT_VALIDATE_DATA_P.GetCommandParamValue("O_MESSAGE"));
            if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
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
                idaSLIP_LINE.MoveFirst(this.Name);
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private void idaSLIP_HEADER_UpdateCompleted(object pSender)
        { 
            mSave_Appr_Status = true;
            foreach (DataRow vRow in IDA_APPROVAL_PERSON.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Unchanged)
                {
                    mSave_Appr_Status = false;
                }
            }
            IDA_APPROVAL_PERSON.Update();

            if (mSave_Appr_Status == false)
            {
                return;
            }
             
            if (idaSLIP_HEADER.UpdateChangedRowCount != 0)
            {
                IDC_SLIP_AUTO_REQUEST_FLAG_P.ExecuteNonQuery();
                string vAUTO_REQUEST_FLAG = iString.ISNull(IDC_SLIP_AUTO_REQUEST_FLAG_P.GetCommandParamValue("O_AUTO_REQUEST_FLAG"));
                if (vAUTO_REQUEST_FLAG == "Y" && gBTN_Req_YN == "N")
                {
                    SLIP_REQ_OK();
                }
            }

            string vSLIP_NUM = iString.ISNull(SLIP_NUM.EditValue); // igrSLIP_LIST.GetCellValue("GL_NUM"));
            int vIDX_SLIP_NUM = igrSLIP_BUDGET_LIST.GetColumnToIndex("SLIP_NUM");
            Search_DB();

            // 기존 위치 이동 : 없을 경우.
            for (int r = 0; r < igrSLIP_BUDGET_LIST.RowCount; r++)
            {
                if (vSLIP_NUM == iString.ISNull(igrSLIP_BUDGET_LIST.GetCellValue(r, vIDX_SLIP_NUM)))
                {
                    igrSLIP_BUDGET_LIST.CurrentCellMoveTo(r, vIDX_SLIP_NUM);
                    igrSLIP_BUDGET_LIST.CurrentCellActivate(r, vIDX_SLIP_NUM);
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
            if (iString.ISNull(e.Row["VAT_ASSET_GB"]) == string.Empty)
            {// 자산구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Asset Type(자산구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

#endregion

#region ---- Doc Att / Appr Step ----

        private void Init_SLIP_APPR()
        {
            if (mAPPROVAL_YN == "N")
            {
                return;
            }

            IDA_APPROVAL_PERSON.Cancel();
            if (IDA_APPROVAL_PERSON.CurrentRows.Count > 0)
            {
                return;
            }

            IDA_INIT_APPROVAL_LINE.SetSelectParamValue("W_SLIP_DATE", SLIP_DATE.EditValue);
            IDA_INIT_APPROVAL_LINE.SetSelectParamValue("W_DEPT_ID", BUDGET_DEPT_ID.EditValue);
            IDA_INIT_APPROVAL_LINE.Fill();
            foreach (DataRow row in IDA_INIT_APPROVAL_LINE.SelectRows)
            {
                IDA_APPROVAL_PERSON.AddUnder();

                IGR_APPROVAL_PERSON.SetCellValue("APPROVAL_PERSON_ID", row["APPROVAL_PERSON_ID"]);
                IGR_APPROVAL_PERSON.SetCellValue("SLIP_DATE", row["SLIP_DATE"]);
                IGR_APPROVAL_PERSON.SetCellValue("SLIP_NUM", row["SLIP_NUM"]);
                IGR_APPROVAL_PERSON.SetCellValue("DEPT_ID", row["DEPT_ID"]);
                IGR_APPROVAL_PERSON.SetCellValue("APPROVAL_STEP_SEQ", row["APPROVAL_STEP_SEQ"]);
                IGR_APPROVAL_PERSON.SetCellValue("APPROVAL_STEP_ID", row["APPROVAL_STEP_ID"]);
                IGR_APPROVAL_PERSON.SetCellValue("APPROVAL_STEP", row["APPROVAL_STEP"]);
                IGR_APPROVAL_PERSON.SetCellValue("APPROVAL_STEP_NAME", row["APPROVAL_STEP_NAME"]);
                IGR_APPROVAL_PERSON.SetCellValue("PERSON_ID", row["PERSON_ID"]);
                IGR_APPROVAL_PERSON.SetCellValue("PERSON_NUM", row["PERSON_NUM"]);
                IGR_APPROVAL_PERSON.SetCellValue("PERSON_NAME", row["PERSON_NAME"]);
                IGR_APPROVAL_PERSON.SetCellValue("EMAIL", row["EMAIL"]);
                IGR_APPROVAL_PERSON.SetCellValue("DESCRIPTION", row["DESCRIPTION"]);
                IGR_APPROVAL_PERSON.SetCellValue("REQUIRED_FLAG", row["REQUIRED_FLAG"]);
                IGR_APPROVAL_PERSON.SetCellValue("APPR_FLAG", row["APPR_FLAG"]);
            }
        }

        private void IDA_APPROVAL_PERSON_UpdateCompleted(object pSender)
        {
            if(IDA_APPROVAL_PERSON.UpdateModifiedRowCount != 0)
            {
                mSave_Appr_Status = true;
            }
        } 
        
        private void BTN_INSERT_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.AddUnder();
            IGR_APPROVAL_PERSON.CurrentCellMoveTo(IGR_APPROVAL_PERSON.GetColumnToIndex("APPROVAL_STEP_SEQ"));
            IGR_APPROVAL_PERSON.CurrentCellActivate(IGR_APPROVAL_PERSON.GetColumnToIndex("APPROVAL_STEP_SEQ"));
        }

        private void BTN_CANCEL_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.Cancel();
        }

        private void BTN_DELETE_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.Delete();
        }
         
        private void BTN_CLOSED_A_ButtonClick(object pSender, EventArgs pEventArgs)            
        {
            mSave_Appr_Status = true;
            if (iString.ISNull(SLIP_NUM.EditValue) != string.Empty)
            {
                foreach (DataRow vRow in IDA_APPROVAL_PERSON.CurrentRows)
                {
                    if(vRow.RowState != DataRowState.Unchanged)
                    {
                        mSave_Appr_Status = false;
                    }
                }

                if (mSave_Appr_Status == false)
                {
                    try
                    {
                        IDA_APPROVAL_PERSON.Update();
                    }
                    catch
                    {
                        return;
                    }
                    Init_Sub_Panel(false, "APPR_STEP");
                }
                else
                {                    
                    Init_Sub_Panel(false, "APPR_STEP");
                } 
            } 
            Init_Sub_Panel(false, "APPR_STEP");
        }

        private void DOC_ATT_FLAG()
        {
            IDC_GET_DOC_ATT_FLAG_P.ExecuteNonQuery();
            if (iString.ISNull(IDC_GET_DOC_ATT_FLAG_P.GetCommandParamValue("O_DOC_ATT_FLAG")) == "Y")
            {
                CB_DOC_ATT_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
            else
            {
                CB_DOC_ATT_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            }
        }

        private void BTN_FILE_ATTACH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iString.ISNull(SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            AssmblyRun_Attachment("FCMF0228", SLIP_DATE.EditValue, SLIP_NUM.EditValue); 
        }

        private void BTN_DOC_ATT_L_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            if (iString.ISNull(igrSLIP_BUDGET_LIST.GetCellValue("SLIP_NUM")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            AssmblyRun_Attachment("FCMF0228", igrSLIP_BUDGET_LIST.GetCellValue("SLIP_DATE"), igrSLIP_BUDGET_LIST.GetCellValue("SLIP_NUM"));             
        }
         
        private void BTN_APPR_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_GET_APPROVAL_PERSON_STATUS.SetCommandParamValue("W_APPROVAL_STEP_SEQ", DBNull.Value);
            IDC_GET_APPROVAL_PERSON_STATUS.ExecuteNonQuery();
            String vSTATUS = iString.ISNull(IDC_GET_APPROVAL_PERSON_STATUS.GetCommandParamValue("O_STATUS"));
            String vMESSAGE = iString.ISNull(IDC_GET_APPROVAL_PERSON_STATUS.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                if(vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            Init_SLIP_APPR();
            A_SLIP_DATE.EditValue = SLIP_DATE.EditValue;
            A_SLIP_NUM.EditValue = SLIP_NUM.EditValue;
            Init_Sub_Panel(true, "APPR_STEP");
        }
         
        private void GB_APPR_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;
               
                Point I = GB_APPR.Location;
                I.Offset(gx, gy);
                GB_APPR.Location = I;
            }
        }

        private void GB_APPR_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_APPR_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }
         
        private void isRadioButtonAdv1_CheckChanged(object sender, EventArgs e)
        {

        }

        private void itbSLIP_SelectedIndexChanged(object sender, EventArgs e)
        {
            Search_DB();
        }

        private void ilaSLIP_TYPE_SelectedRowData(object pSender)
        {
            //승인단계 설정.
            Init_SLIP_APPR();
        }
         
    #endregion

    }
}