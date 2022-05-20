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

namespace FCMF0210
{
    public partial class FCMF0210 : Office2007Form
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
        string mCurrency_Code;
        object mBudget_Control_YN;
        bool mSUB_SHOW_FLAG = false;

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

        #endregion;

        #region ----- Constructor -----

        public FCMF0210()
        {
            InitializeComponent();
        }

        public FCMF0210(Form pMainForm, ISAppInterface pAppInterface)
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
            if (iString.ISNull(idcACCOUNT_BOOK.GetCommandParamValue("O_SLIP_REMARK_FLAG")) == "Y")
            {
                REMARK.LookupAdapter = ILA_SLIP_REMARK;
            }
            else
            {
                REMARK.LookupAdapter = null;
            }

            //전표 승인단계 관리 여부.
            IDC_GET_SLIP_CONFIG_P.ExecuteNonQuery();
            mATT_FILE_YN = iString.ISNull(IDC_GET_SLIP_CONFIG_P.GetCommandParamValue("O_ATT_FILE_YN"));
            mAPPROVAL_YN = iString.ISNull(IDC_GET_SLIP_CONFIG_P.GetCommandParamValue("O_APPROVAL_YN"));
            if (mATT_FILE_YN == "N")
            {
                BTN_DOC_ATT_L.Visible = false;
                BTN_FILE_ATTACH.Visible = false;
                CB_DOC_ATT_FLAG.Visible = false;
                igrSLIP_LIST.GridAdvExColElement[igrSLIP_LIST.GetColumnToIndex("DOC_ATT_FLAG")].Visible = 0;
            }
            if (mAPPROVAL_YN == "N")
            {
                BTN_APPR_STEP.Visible = false;
            }
            igrSLIP_LIST.ResetDraw = true;
        }

        private void Search_DB()
        {
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            
            if (itbSLIP.SelectedTab.TabIndex == 2)
            {
                Search_DB_DETAIL(H_SLIP_HEADER_ID.EditValue);
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

                idaSLIP_HEADER_LIST.OraSelectData.AcceptChanges();
                idaSLIP_HEADER_LIST.Refillable = true;
                igrSLIP_LIST.LastConfirmChanges();

                string vGL_NUM = iString.ISNull(igrSLIP_LIST.GetCellValue("GL_NUM"));
                int vCOL_IDX = igrSLIP_LIST.GetColumnToIndex("GL_NUM");
                idaSLIP_HEADER_LIST.Fill();
                if (iString.ISNull(vGL_NUM) != string.Empty)
                {
                    for (int i = 0; i < igrSLIP_LIST.RowCount; i++)
                    {
                        if (vGL_NUM == iString.ISNull(igrSLIP_LIST.GetCellValue(i, vCOL_IDX)))
                        {
                            igrSLIP_LIST.CurrentCellMoveTo(i, vCOL_IDX);
                            igrSLIP_LIST.CurrentCellActivate(i, vCOL_IDX);
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
                if (mSUB_SHOW_FLAG == true)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                                
                SLIP_QUERY_STATUS.EditValue = "QUERY";
                itbSLIP.SelectedIndex = 1;
                itbSLIP.SelectedTab.Focus();
                idaSLIP_HEADER.SetSelectParamValue("W_SLIP_HEADER_ID", pSLIP_HEADER_ID);
                try
                {
                    idaSLIP_HEADER.Fill();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }

                //첨부파일 여부.//
                DOC_ATT_FLAG();

                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;
                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;
            }
        }

        private void Search_DB_DPR_ASSET()
        {
            idaDPR_ASSET.Fill();

            igrDPR_ASSET.CurrentCellMoveTo(1);
            igrDPR_ASSET.CurrentCellActivate(1);
            igrDPR_ASSET.Focus();
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
            igrSLIP_LINE.Invalidate();
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

        private void SetManagementParameter(string pManagement_Field, string pEnabled_YN, object pLookup_Type)
        {
            string mLookup_Type = iString.ISNull(pLookup_Type);
            
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
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", H_DEPT_CODE.EditValue);
            }
            else if (mLookup_Type == "COSTCENTER".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("DEPT"));
            }
            else if (mLookup_Type == "BANK_ACCOUNT".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("BANK_SITE"));
            }
            else if (mLookup_Type == "BANK_SITE".ToString())
            {
                ildMANAGEMENT.SetLookupParamValue("W_INQURIY_VALUE", GetLookup_Type("CUSTOMER"));
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
                if (iString.ISNull(GL_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = GL_DATE.DateTimeValue.ToShortDateString();
                }
                else if (iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    vGL_DATE = SLIP_DATE.DateTimeValue.ToShortDateString();
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

        private object GetLookup_Type(object pLookup_Type)
        {
            if (iString.ISNull(pLookup_Type) == string.Empty)
            {
                return null;
            }
            object mLookup_Value;
            if (iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT1.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT2_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = MANAGEMENT2.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER1_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER1_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER1.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER2_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER2_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER2.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER3_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER3_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER3.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER4_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER4_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER4.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER5_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER5_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER5.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER6_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER6_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER6.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER7_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER7_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER7.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["REFER8_LOOKUP_TYPE"]) != string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRow["REFER8_LOOKUP_TYPE"]) == iString.ISNull(pLookup_Type))
            {
                mLookup_Value = REFER8.EditValue;
            }
            else
            {
                mLookup_Value = null;
            }
            return mLookup_Value;
        }

        //private object GetLookup_Type(object pLookup_Type)
        //{
        //    if (iString.ISNull(pLookup_Type) == string.Empty)
        //    {
        //        return null;
        //    }
        //    object mLookup_Value;
        //    if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = MANAGEMENT1.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = MANAGEMENT2.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER1.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER2.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER3.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER4.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER5.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER6.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER7.EditValue;
        //    }
        //    else if (iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")) != string.Empty
        //        && iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE")) == iString.ISNull(pLookup_Type))
        //    {
        //        mLookup_Value = REFER8.EditValue;
        //    }
        //    else
        //    {
        //        mLookup_Value = null;
        //    }
        //    return mLookup_Value;
        //}

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
            ibtSUB_FORM.Left = 776;
            ibtSUB_FORM.Top = 73;
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

        private void Set_CheckBox()
        {
            int mIDX_Col = igrSLIP_LIST.GetColumnToIndex("SELECT_YN");

            object mCheck_YN = CB_SELECT_YN.CheckBoxValue;
            for (int r = 0; r < igrSLIP_LIST.RowCount; r++)
            {
                igrSLIP_LIST.SetCellValue(r, mIDX_Col, mCheck_YN);
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

                        GB_APPR.BringToFront();
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
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Left = 35;
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
                    else if (pSub_Panel == "DOC_ATT")
                    {
                        GB_DOC_ATT.Left = 340;
                        GB_DOC_ATT.Top = 100;

                        GB_DOC_ATT.Width = 415;
                        GB_DOC_ATT.Height = 230;

                        GB_DOC_ATT.Border3DStyle = Border3DStyle.Bump;
                        GB_DOC_ATT.BorderStyle = BorderStyle.Fixed3D;

                        GB_DOC_ATT.Controls[0].MouseDown += GB_DOC_ATT_MouseDown;
                        GB_DOC_ATT.Controls[0].MouseMove += GB_DOC_ATT_MouseMove;
                        GB_DOC_ATT.Controls[0].MouseUp += GB_DOC_ATT_MouseUp;
                        GB_DOC_ATT.Controls[1].MouseDown += GB_DOC_ATT_MouseDown;
                        GB_DOC_ATT.Controls[1].MouseMove += GB_DOC_ATT_MouseMove;
                        GB_DOC_ATT.Controls[1].MouseUp += GB_DOC_ATT_MouseUp;

                        GB_DOC_ATT.BringToFront();
                        GB_DOC_ATT.Visible = true;
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
            }
            else
            {
                try
                {
                    if (pSub_Panel == "ALL")
                    {
                        GB_AP_VAT.Visible = false;
                        GB_COPY_DOCUMENT.Visible = false;
                        GB_APPR.Visible = false;
                        GB_DOC_ATT.Visible = false;
                    }
                    else if (pSub_Panel == "AP_VAT")
                    {
                        GB_AP_VAT.Visible = false;
                    }
                    else if (pSub_Panel == "COPY_SLIP")
                    {
                        GB_COPY_DOCUMENT.Visible = false;
                    }
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Visible = false;
                    }
                    else if (pSub_Panel == "DOC_ATT")
                    {
                        GB_DOC_ATT.Visible = false;
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

        private bool Check_SlipHeader_Added()
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
            SLIP_DATE.EditValue = DateTime.Today;
            GL_DATE.EditValue = SLIP_DATE.EditValue;

            idcDV_SLIP_TYPE.SetCommandParamValue("W_GROUP_CODE", "SLIP_TYPE");
            idcDV_SLIP_TYPE.ExecuteNonQuery();
            SLIP_TYPE.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_CODE");
            SLIP_TYPE_NAME.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_CODE_NAME");
            SLIP_TYPE_CLASS.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_VALUE1");
            DOCUMENT_TYPE.EditValue = idcDV_SLIP_TYPE.GetCommandParamValue("O_VALUE2");

            idcUSER_INFO.ExecuteNonQuery();
            DEPT_NAME.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_NAME");
            H_DEPT_CODE.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_CODE");
            DEPT_ID.EditValue = idcUSER_INFO.GetCommandParamValue("O_DEPT_ID");
            PERSON_NAME.EditValue = idcUSER_INFO.GetCommandParamValue("O_PERSON_NAME");
            PERSON_ID.EditValue = isAppInterfaceAdv1.PERSON_ID;

            //헤더 예산부서
            H_BUDGET_DEPT_NAME.EditValue = DEPT_NAME.EditValue;
            H_BUDGET_DEPT_CODE.EditValue = H_DEPT_CODE.EditValue;
            H_BUDGET_DEPT_ID.EditValue = DEPT_ID.EditValue;             
        }

        private void InsertSlipLine()
        {
            //LINE SEQ 채번//
            try
            {
                int mPreviousRowPosition = idaSLIP_LINE.CurrentRowPosition() - 1;
                decimal mSLIP_LINE_SEQ;

                if (mPreviousRowPosition > -1)
                {
                    mSLIP_LINE_SEQ = iString.ISDecimaltoZero(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["SLIP_LINE_SEQ"], 1);
                    if (idaSLIP_LINE.CurrentRows[mPreviousRowPosition].RowState == DataRowState.Added)
                    {
                        if (mSLIP_LINE_SEQ - Math.Truncate(mSLIP_LINE_SEQ) != 0)
                        {
                            mSLIP_LINE_SEQ = iString.ISDecimaltoZero(Convert.ToDouble(mSLIP_LINE_SEQ) + Convert.ToDouble(0.001));
                        }
                        else
                        {
                            mSLIP_LINE_SEQ = mSLIP_LINE_SEQ + 1;
                        }
                    }
                    else
                    {
                        mSLIP_LINE_SEQ = iString.ISDecimaltoZero(Convert.ToDouble(mSLIP_LINE_SEQ) + Convert.ToDouble(0.001));
                    }
                }
                else
                {
                    mSLIP_LINE_SEQ = 1;
                }
                SLIP_LINE_SEQ.EditValue = mSLIP_LINE_SEQ;
            }
            catch
            {

            }

            CURRENCY_CODE.EditValue = mCurrency_Code;
            CURRENCY_DESC.EditValue = mCurrency_Code;
            Init_Currency_Amount();
            Init_Budget_Dept();

            BUDGET_DEPT_NAME_L.Focus();
        }

        private void Set_Insert_Slip_Line()
        {
            IDA_BALANCE_SLIP_LINE.Fill();
            if (IDA_BALANCE_SLIP_LINE.SelectRows.Count < 1)
            {
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Not found data, Check data");
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int Row_Count = igrSLIP_LINE.RowCount;
            igrSLIP_LINE.BeginUpdate();
            try
            {
                for (int i = 0; i < IDA_BALANCE_SLIP_LINE.SelectRows.Count; i++)
                {
                    idaSLIP_LINE.AddUnder();
                    for (int c = 0; c < igrSLIP_LINE.GridAdvExColElement.Count; c++)
                    {
                        if (igrSLIP_LINE.GridAdvExColElement[c].DataColumn.ToString() != "SLIP_HEADER_ID")
                        {
                            igrSLIP_LINE.SetCellValue(i + Row_Count, c, IDA_BALANCE_SLIP_LINE.OraDataSet().Rows[i][c]);
                        }
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

            //delete temp data
            Delete_Balance_Remain_TP();

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
            decimal mGL_AMOUNT = iString.ISDecimaltoZero(GL_CURRENCY_AMOUNT.EditValue) * iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue);
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
                idaSLIP_LINE.AddUnder();
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
                GL_CURRENCY_AMOUNT.EditValue = iString.ISDecimaltoZero(pCurrency_Amount);
                GL_AMOUNT.EditValue = Math.Abs(iString.ISDecimaltoZero(vExchange_Profit_Loss_Amount));

                //참고항목 동기화.                
                Set_Control_Item_Prompt();
                Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);

                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
                mExchange_Profit_Loss = true;
            }
            return mExchange_Profit_Loss;
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
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, idaSLIP_LINE.CurrentRow["GL_AMOUNT"]);
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, 0);
                }
                else if (iString.ISNull(idaSLIP_LINE.CurrentRow["ACCOUNT_DR_CR"], "1") == "2".ToString())
                {
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_DR, 0);
                    igrSLIP_LINE.SetCellValue(vIDX_ROW_CURR, vIDX_COL_CR, idaSLIP_LINE.CurrentRow["GL_AMOUNT"]);
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

            foreach (DataRow vRow in idaSLIP_LINE.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Deleted)
                {
                    if (iString.ISNull(vRow["ACCOUNT_DR_CR"], "1") == "1".ToString())
                    {
                        vDR_Amount = vDR_Amount + iString.ISDecimaltoZero(vRow["GL_AMOUNT"]);
                        vCurrency_DR_Amount = vCurrency_DR_Amount + iString.ISDecimaltoZero(vRow["GL_CURRENCY_AMOUNT"]);
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
            mDATA_VALUE = MANAGEMENT1.EditValue;
            MANAGEMENT1.Nullable = true;
            mDATA_TYPE = iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["MANAGEMENT1_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["MANAGEMENT2_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER1_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER1_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER2_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER2_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER3_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER3_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER4_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER4_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER5_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER5_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER6_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER6_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER7_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER7_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
            mDATA_TYPE = iString.ISNull(pDataRow["REFER8_DATA_TYPE"]);
            mDR_CR_YN = iString.ISNull(pDataRow["REFER8_YN"]); 
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = igrSLIP_LINE.GetCellValue("REFER8_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
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
                    MANAGEMENT1.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    MANAGEMENT2.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER1.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER2.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER3.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER4.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER5.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER6.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER7.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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
                    REFER8.EditValue = iDate.ISGetDate(SLIP_DATE.EditValue).ToShortDateString();
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

        private void Init_Currency_Amount()
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) == string.Empty || CURRENCY_CODE.EditValue.ToString() == mCurrency_Code.ToString())
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
                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                GL_CURRENCY_AMOUNT.ReadOnly = false;
                GL_CURRENCY_AMOUNT.Insertable = true;
                GL_CURRENCY_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                GL_CURRENCY_AMOUNT.TabStop = true;
            }
            EXCHANGE_RATE.Invalidate();
            GL_CURRENCY_AMOUNT.Invalidate();
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
            decimal mSUPPLY_AMOUNT = mGL_AMOUNT * vVAT_RATE; //공급가액 설정.

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
                BUDGET_DEPT_NAME_L.EditValue = H_BUDGET_DEPT_NAME.EditValue;
                BUDGET_DEPT_CODE_L.EditValue = H_BUDGET_DEPT_CODE.EditValue;
                BUDGET_DEPT_ID_L.EditValue = H_BUDGET_DEPT_ID.EditValue;
            }
        }

        //부서 
        private void Init_Dept()
        {
            if (iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == "DEPT" &&
                iString.ISNull(MANAGEMENT1.EditValue) == String.Empty)
            {
                MANAGEMENT1_DESC.EditValue = BUDGET_DEPT_NAME_L.EditValue;
                MANAGEMENT1.EditValue = BUDGET_DEPT_CODE_L.EditValue;
            }
            else if (iString.ISNull(idaSLIP_LINE.CurrentRow["MANAGEMENT1_LOOKUP_TYPE"]) == "DEPT")
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

        private void AssmblyRun_Manual(object pAssembly_ID, object pSlip_Header_ID, object pGL_Date, object pGL_Num, object pSLIP_Date, object pSLIP_Num)
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
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive != "N")
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

                            object[] vParam = new object[8];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = pSlip_Header_ID;     //전표 헤더 id
                            vParam[3] = pGL_Date;     //전표일자
                            vParam[4] = pGL_Num;                   //전표번호
                            vParam[5] = pSLIP_Date;     //전표일자
                            vParam[6] = pSLIP_Num;                   //전표번호
                            vParam[7] = "N";      //프린트 옵션 표시 여부

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
            object vSLIP_Date;
            object vSLIP_Num;

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            if (itbSLIP.SelectedTab.TabIndex == 2)
            {
                vSlip_Header_id = H_SLIP_HEADER_ID.EditValue;
                vGL_Date = GL_DATE.EditValue;
                vGL_Num = GL_NUM.EditValue;
                vSLIP_Date = SLIP_DATE.EditValue;
                vSLIP_Num = SLIP_NUM.EditValue;

                AssmblyRun_Manual("FCMF0211", vSlip_Header_id, vGL_Date, vGL_Num, vSLIP_Date, vSLIP_Num);
            }
            else
            {
                string vMessageText = string.Empty; 

                int vCountRowGrid = igrSLIP_LIST.RowCount;
                if (vCountRowGrid > 0)
                {
                    vMessageText = string.Format("Printing Start");
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText); 
                     
                    int vIDX_SELECT_YN = igrSLIP_LIST.GetColumnToIndex("SELECT_YN");
                    int vIDX_SLIP_HEADER_ID = igrSLIP_LIST.GetColumnToIndex("SLIP_HEADER_ID");
                    int vIDX_GL_DATE = igrSLIP_LIST.GetColumnToIndex("GL_DATE");
                    int vIDX_GL_NUM = igrSLIP_LIST.GetColumnToIndex("GL_NUM");
                    int vIDX_SLIP_DATE = igrSLIP_LIST.GetColumnToIndex("SLIP_DATE");
                    int vIDX_SLIP_NUM = igrSLIP_LIST.GetColumnToIndex("SLIP_NUM");

                    //-------------------------------------------------------------------------------------
                    for (int vRow = 0; vRow < vCountRowGrid; vRow++)
                    {
                        object vSELECT_YN = igrSLIP_LIST.GetCellValue(vRow, vIDX_SELECT_YN);
                        if (iString.ISNull(vSELECT_YN) == "Y")
                        {
                            igrSLIP_LIST.CurrentCellMoveTo(vRow, vIDX_SELECT_YN);
                            igrSLIP_LIST.CurrentCellActivate(vRow, vIDX_SELECT_YN);
                             
                            vSlip_Header_id = igrSLIP_LIST.GetCellValue(vRow, vIDX_SLIP_HEADER_ID);
                            vGL_Date = igrSLIP_LIST.GetCellValue(vRow, vIDX_GL_DATE);
                            vGL_Num = igrSLIP_LIST.GetCellValue(vRow, vIDX_GL_NUM);
                            vSLIP_Date = igrSLIP_LIST.GetCellValue(vRow, vIDX_SLIP_DATE);
                            vSLIP_Num = igrSLIP_LIST.GetCellValue(vRow, vIDX_SLIP_NUM);

                            AssmblyRun_Manual("FCMF0211", vSlip_Header_id, vGL_Date, vGL_Num, vSLIP_Date, vSLIP_Num);

                            igrSLIP_LIST.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                        }
                    }
                    //-------------------------------------------------------------------------------------
                }
            } 

            igrSLIP_LIST.LastConfirmChanges();
            idaSLIP_HEADER_LIST.OraSelectData.AcceptChanges();
            idaSLIP_HEADER_LIST.Refillable = true;

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
                    //if (idaSLIP_LINE.IsFocused)
                    //{
                    //    if (Check_Sub_Panel() == false)
                    //    {
                    //        return;
                    //    } 

                    //    idaSLIP_LINE.AddUnder();
                    //    InsertSlipLine();
                    //}
                    //else
                    //{
                    //    if (Check_SlipHeader_Added() == true)
                    //    {
                    //        return;
                    //    }
                    //    else
                    //    {
                    //        idaSLIP_HEADER.SetSelectParamValue("W_SLIP_HEADER_ID", 0);
                    //        idaSLIP_HEADER.Fill();

                    //        if (Check_Sub_Panel() == false)
                    //        {
                    //            return;
                    //        } 

                    //        idaSLIP_HEADER.AddUnder();
                    //        idaSLIP_LINE.AddUnder();
                    //        InsertSlipHeader();
                    //        InsertSlipLine();

                    //        SLIP_DATE.Focus();
                    //    }
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    //if (idaSLIP_HEADER_LIST.IsFocused)
                    //{
                    //    idaSLIP_HEADER_LIST.Update();
                    //}
                    //else
                    //{
                    //    ACCOUNT_CODE.Focus();

                    //    Init_DR_CR_Amount();    // 차대금액 생성 //
                    //    Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //

                    //    if (iString.ISDecimaltoZero(TOTAL_DR_AMOUNT.EditValue) != iString.ISDecimaltoZero(TOTAL_CR_AMOUNT.EditValue))
                    //    {// 차대금액 일치 여부 체크.
                    //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10134"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //        return;
                    //    }

                    //    if (Check_Sub_Panel() == false)
                    //    {
                    //        return;
                    //    } 

                    //    idaSLIP_HEADER.Update();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    SLIP_QUERY_STATUS.EditValue = "QUERY";
                    if (idaSLIP_HEADER_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_LIST.Cancel();
                    }
                    else if (idaSLIP_HEADER.IsFocused)
                    {
                        if (Check_Sub_Panel() == false)
                        {
                            return;
                        }
                        idaDPR_ASSET.Cancel();
                        idaSLIP_LINE.Cancel();
                        idaSLIP_HEADER.Cancel();
                    }
                    else if (idaSLIP_LINE.IsFocused)
                    {
                        if (Check_Sub_Panel() == false)
                        {
                            return;
                        }
                        idaDPR_ASSET.Cancel();                        
                        idaSLIP_LINE.Cancel();
                        Init_Total_GL_Amount();  //합계 금액 재 계산 //
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (idaSLIP_HEADER_LIST.IsFocused)
                    //{
                    //    idaSLIP_HEADER_LIST.Delete();
                    //}
                    //else if (idaSLIP_HEADER.IsFocused)
                    //{
                    //    if (Check_Sub_Panel() == false)
                    //    {
                    //        return;
                    //    }

                    //    for (int r = 0; r < igrDPR_ASSET.RowCount; r++)
                    //    {
                    //        idaDPR_ASSET.Delete();
                    //    }
                    //    for (int r = 0; r < igrSLIP_LINE.RowCount; r++)
                    //    {
                    //        idaSLIP_LINE.Delete();
                    //    }
                    //    idaSLIP_HEADER.Delete();
                    //}
                    //else if (idaSLIP_LINE.IsFocused)
                    //{
                    //    if (Check_Sub_Panel() == false)
                    //    {
                    //        return;
                    //    }
                    //    for (int r = 0; r < igrDPR_ASSET.RowCount; r++)
                    //    {
                    //        idaDPR_ASSET.Delete();
                    //    }
                    //    idaSLIP_LINE.Delete();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (Check_Sub_Panel() == false)
                    {
                        return;
                    } 

                    XLPrinting_Main();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (Check_Sub_Panel() == false)
                    {
                        return;
                    }

                    XLPrinting_Main();
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0210_Load(object sender, EventArgs e)
        {            
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            SLIP_DATE_FR_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            SLIP_DATE_TO_0.EditValue = iDate.ISGetDate();

            // 회계장부 정보 설정.
            GetAccountBook();

            // 콤퍼넌트 동기화.
            //Init_Currency_Code();
            
            //서브판넬 
            Init_Sub_Panel(false, "ALL");

            //전표 복사 버튼 맨 앞으로 가져오기
            btnGET_BALANCE_STATEMENT.BringToFront();
            BTN_COPY_SLIP.BringToFront();
            BTN_DOC_ATT_L.BringToFront();

            //서브 화면
            ibtSUB_FORM.Visible = false;            
        }

        private void FCMF0210_Shown(object sender, EventArgs e)
        {
            idaSLIP_HEADER_LIST.FillSchema();
            idaSLIP_HEADER.FillSchema();
        }

        private void igrSLIP_LIST_CellDoubleClick(object pSender)
        {
            if (igrSLIP_LIST.RowCount > 0)
            {
                Search_DB_DETAIL(igrSLIP_LIST.GetCellValue("SLIP_HEADER_ID"));
            }
        }

        private void H_REMARK_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(REMARK.EditValue) == string.Empty)
            {
                REMARK.EditValue = H_REMARK.EditValue;
            }
        }

        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (idaSLIP_LINE.CurrentRow != null && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_GL_Amount();
            }
        }

        private void GL_CURRENCY_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (idaSLIP_LINE.CurrentRow != null && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_GL_Amount();
            }
        }

        private void GL_AMOUNT_EditValueChanged(object pSender)
        {
            if (idaSLIP_LINE.CurrentRow != null && idaSLIP_LINE.CurrentRow.RowState != DataRowState.Unchanged)
            {
                Init_DR_CR_Amount();    // 차대금액 생성 //
                Init_VAT_Amount();
            }
        }

        private void GL_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
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
            //if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "RECEIVABLE_BILL".ToString())
            //{//받을어음
            //    object mBILL_TYPE = "2";  // 어음구분.
            //    object mBILL_NUM = Get_Management_Value("RECEIVABLE_BILL");
            //    object mBILL_AMOUNT = GL_AMOUNT.EditValue;
            //    object mVENDOR_CODE = Get_Management_Value("CUSTOMER");
            //    object mBANK_CODE = Get_Management_Value("BANK");
            //    object mVAT_ISSUE_DATE = Get_Management_Value("VAT_ISSUE_DATE");
            //    object mISSUE_DATE = Get_Management_Value("ISSUE_DATE");
            //    object mDUE_DATE = Get_Management_Value("DUE_DATE");
            //    object mDEPT_ID = DEPT_ID.EditValue;
            //    object mDEPT_NAME = DEPT_NAME.EditValue;

            //    FCMF0210_BILL vFCMF0210_BILL = new FCMF0210_BILL(isAppInterfaceAdv1.AppInterface, mDEPT_ID, mDEPT_NAME
            //                                                        , mBILL_TYPE, mBILL_NUM, mBILL_AMOUNT
            //                                                        , mVENDOR_CODE, mBANK_CODE
            //                                                        , mVAT_ISSUE_DATE, mISSUE_DATE, mDUE_DATE);
            //    dlgResult = vFCMF0210_BILL.ShowDialog();
            //    if (dlgResult == DialogResult.OK)
            //    {
            //        //어음금액
            //        GL_AMOUNT.EditValue = vFCMF0210_BILL.Get_BILL_AMOUNT;
            //        //거래처.
            //        Set_Management_Value("CUSTOMER", vFCMF0210_BILL.Get_VENDOR_CODE, vFCMF0210_BILL.Get_VENDOR_NAME);
            //        //은행
            //        Set_Management_Value("BANK", vFCMF0210_BILL.Get_BANK_CODE, vFCMF0210_BILL.Get_BANK_NAME);
            //        //세금계산서발행일
            //        Set_Management_Value("VAT_ISSUE_DATE", vFCMF0210_BILL.Get_VAT_ISSUE_DATE, null);
            //        //발행일자
            //        Set_Management_Value("ISSUE_DATE", vFCMF0210_BILL.Get_ISSUE_DATE, null);
            //        //만기일자
            //        Set_Management_Value("DUE_DATE", vFCMF0210_BILL.Get_DUE_DATE, null);
            //        //어음번호.
            //        Set_Management_Value("RECEIVABLE_BILL", vFCMF0210_BILL.Get_BILL_NUM, String.Format("{0:###,###,###,###,###,###}", vFCMF0210_BILL.Get_BILL_AMOUNT));
                     
            //        Init_DR_CR_Amount();    // 차대금액 생성 //
            //        Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
            //    }
            //    vFCMF0210_BILL.Dispose();
            //}
            //else if (iString.ISNull(ACCOUNT_CLASS_TYPE.EditValue) == "PAYABLE_BILL".ToString())
            //{//지급어음
            //    object mBILL_TYPE = "1";  // 어음구분.
            //    object mBILL_NUM = Get_Management_Value("PAYABLE_BILL");
            //    object mBILL_AMOUNT = GL_AMOUNT.EditValue;
            //    object mVENDOR_CODE = Get_Management_Value("CUSTOMER");
            //    object mBANK_CODE = Get_Management_Value("BANK");
            //    object mVAT_ISSUE_DATE = Get_Management_Value("VAT_ISSUE_DATE");
            //    object mISSUE_DATE = Get_Management_Value("ISSUE_DATE");
            //    object mDUE_DATE = Get_Management_Value("DUE_DATE");
            //    object mDEPT_ID = DEPT_ID.EditValue;
            //    object mDEPT_NAME = DEPT_NAME.EditValue;

            //    FCMF0210_BILL vFCMF0210_BILL = new FCMF0210_BILL(isAppInterfaceAdv1.AppInterface, mDEPT_ID, mDEPT_NAME
            //                                                        , mBILL_TYPE, mBILL_NUM, mBILL_AMOUNT
            //                                                        , mVENDOR_CODE, mBANK_CODE
            //                                                        , mVAT_ISSUE_DATE, mISSUE_DATE, mDUE_DATE);

            //    dlgResult = vFCMF0210_BILL.ShowDialog();
            //    if (dlgResult == DialogResult.OK)
            //    {
            //        //어음금액
            //        GL_AMOUNT.EditValue = vFCMF0210_BILL.Get_BILL_AMOUNT;
            //        //거래처.
            //        Set_Management_Value("CUSTOMER", vFCMF0210_BILL.Get_VENDOR_CODE, vFCMF0210_BILL.Get_VENDOR_NAME);
            //        //은행
            //        Set_Management_Value("BANK", vFCMF0210_BILL.Get_BANK_CODE, vFCMF0210_BILL.Get_BANK_NAME);
            //        //세금계산서발행일
            //        Set_Management_Value("VAT_ISSUE_DATE", vFCMF0210_BILL.Get_VAT_ISSUE_DATE, null);
            //        //발행일자
            //        Set_Management_Value("ISSUE_DATE", vFCMF0210_BILL.Get_ISSUE_DATE, null);
            //        //만기일자
            //        Set_Management_Value("DUE_DATE", vFCMF0210_BILL.Get_DUE_DATE, null);
            //        //어음번호.
            //        Set_Management_Value("PAYABLE_BILL", vFCMF0210_BILL.Get_BILL_NUM, String.Format("{0:###,###,###,###,###,###}", vFCMF0210_BILL.Get_BILL_AMOUNT));
                    
            //        Init_DR_CR_Amount();    // 차대금액 생성 //
            //        Init_Total_GL_Amount(); // 총합계 및 분개 차액 생성 //
            //    }
            //    vFCMF0210_BILL.Dispose();
            //}
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

            //    FCMF0210_ITEM_DEAL vFCMF0210_ITEM_DEAL = new FCMF0210_ITEM_DEAL(isAppInterfaceAdv1.AppInterface, mISSUE_NUM, mCURRENCY_CODE
            //                                                                    , mVENDOR_CODE, mBANK_CODE, mISSUE_DATE);

            //    dlgResult = vFCMF0210_ITEM_DEAL.ShowDialog();
            //    if (dlgResult == DialogResult.OK)
            //    {
            //        //거래처.4
            //        Set_Management_Value("CUSTOMER", vFCMF0210_ITEM_DEAL.Get_VENDOR_CODE,vFCMF0210_ITEM_DEAL.Get_VENDOR_NAME);
 
            //        //구매(공급)확인번호
            //        Set_Management_Value("PC_ISSUE_NO", vFCMF0210_ITEM_DEAL.Get_ISSUE_NUM, DBNull.Value);

            //        Set_Management_Value("BANK", vFCMF0210_ITEM_DEAL.Get_BANK_CODE, vFCMF0210_ITEM_DEAL.Get_BANK_NAME);

            //        Set_Management_Value("ISSUE_DATE", vFCMF0210_ITEM_DEAL.Get_ISSUE_DATE, DBNull.Value); 
            //    }
            //    vFCMF0210_ITEM_DEAL.Dispose();
            //}
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void SLIP_DATE_EditValueChanged(object pSender)
        {
            if (SLIP_DATE.DataAdapter.IsEditing == true)
            {
                GL_DATE.EditValue = SLIP_DATE.EditValue;
            }
        }

        private void CB_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Set_CheckBox();
        }

        private void igrSLIP_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if(e.ColIndex == igrSLIP_LIST.GetColumnToIndex("SELECT_YN"))
            {
                igrSLIP_LIST.LastConfirmChanges();
                idaSLIP_HEADER_LIST.OraSelectData.AcceptChanges();
                idaSLIP_HEADER_LIST.Refillable = true;
            }
        }

        private void btnGET_BALANCE_STATEMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (mSUB_SHOW_FLAG == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 승인된 전표도 처리 가능하도록 변경 //
            //if (CLOSED_YN.CheckedState == ISUtil.Enum.CheckedState.Checked)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            //DialogResult vRESULT;
            //FCMF0210_SET vFCMF0210_SET = new FCMF0210_SET(isAppInterfaceAdv1.AppInterface, "R");
            //vRESULT = vFCMF0210_SET.ShowDialog();
            //if (vRESULT == DialogResult.OK)
            //{
            //    if (iString.ISNull(ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
            //    {
            //        idaSLIP_LINE.Delete();
            //    }
            //    idaSLIP_LINE.MoveLast(igrSLIP_LINE.Name);
            //    Set_Insert_Slip_Line();
            //    Init_Currency_Code("Y");
            //    Init_Currency_Amount();
            //    Init_Total_GL_Amount();
            //}
            //vFCMF0210_SET.Dispose();

            ACCOUNT_CODE.Focus();
        }

        private void BTN_COPY_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //전표 작성중이면 저장후 작업해야 함
            if (iString.ISNull(GL_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10128"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Check_SlipHeader_Added() == true)
            {
                return;
            }

            //서브판넬 
            C_OLD_GL_DATE.EditValue = GL_DATE.EditValue;
            C_OLD_GL_NUM.EditValue = GL_NUM.EditValue;
            C_OLD_SLIP_HEADER_ID.EditValue = H_SLIP_HEADER_ID.EditValue;

            C_NEW_GL_DATE.EditValue = iDate.ISGetDate();

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

            IDC_SET_COPY_SLIP_HEADER.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SET_COPY_SLIP_HEADER.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_COPY_SLIP_HEADER.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_COPY_SLIP_HEADER.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
                return;
            }

            C_NEW_SLIP_HEADER_ID.EditValue = IDC_SET_COPY_SLIP_HEADER.GetCommandParamValue("O_NEW_SLIP_HEADER_ID");
            C_NEW_GL_NUM.EditValue = IDC_SET_COPY_SLIP_HEADER.GetCommandParamValue("O_NEW_GL_NUM");            
        }

        private void C_BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "COPY_SLIP");

            if (CB_NEW_SLIP_SEARCH_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                Search_DB_DETAIL(C_NEW_SLIP_HEADER_ID.EditValue);
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

        private void S_BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
            //서브판넬 
            Init_Sub_Panel(false, "AP_VAT");

            MANAGEMENT1.Focus();    //focus 이동 
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

        private void ilaGL_NUM_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildGL_NUM.SetLookupParamValue("W_GL_NUM", GL_NUM_0.EditValue);
        }

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", " VALUE1 <> 'BL'", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaSLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("SLIP_TYPE", " VALUE1 <> 'BL'", "Y");
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
            ildBUDGET_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildBUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "A");
        }

        private void ilaBUDGET_DEPT_L_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBUDGET_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildBUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", SLIP_DATE.EditValue);
            ildBUDGET_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "A");
        }

        private void ilaBUDGET_DEPT_L_SelectedRowData(object pSender)
        {
            Init_Dept();
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
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
            Set_Control_Item_Prompt();
            Init_Control_Management_Value();
            Init_Set_Item_Prompt(idaSLIP_LINE.CurrentRow);
            Init_Set_Item_Need(idaSLIP_LINE.CurrentRow);
            Init_Default_Value();
            Init_Dept();
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
                }
            }
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBUDGET_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
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
            if (iString.ISNull(e.Row["GL_DATE"]) == string.Empty)
            {// 전표일자.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            // 전표번호 채번//
            if (iString.ISNull(GL_NUM.EditValue) == string.Empty || iString.ISNull(e.Row["GL_DATE"]).Substring(0, 7) != iString.ISNull(e.Row["OLD_GL_DATE"], e.Row["GL_DATE"]).Substring(0, 7))
            {
                GetSlipNum();
            }
            else if(iString.ISNull(SLIP_TYPE.EditValue) != iString.ISNull(OLD_SLIP_TYPE.EditValue))
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
                idaSLIP_LINE.MoveFirst(this.Name);
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private void idaSLIP_HEADER_UpdateCompleted(object pSender)
        {
            string vGL_NUM = iString.ISNull(GL_NUM.EditValue); // igrSLIP_LIST.GetCellValue("GL_NUM"));
            int vIDX_GL_NUM = igrSLIP_LIST.GetColumnToIndex("GL_NUM");
            Search_DB();

            // 기존 위치 이동 : 없을 경우.
            for (int r = 0; r < igrSLIP_LIST.RowCount; r++)
            {
                if (vGL_NUM == iString.ISNull(igrSLIP_LIST.GetCellValue(r, vIDX_GL_NUM)))
                {
                    igrSLIP_LIST.CurrentCellMoveTo(r, vIDX_GL_NUM);
                    igrSLIP_LIST.CurrentCellActivate(r, vIDX_GL_NUM);
                }
            }
            SLIP_TYPE_NAME.Focus();
        }

        private void idaSLIP_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            //if (pBindingManager.DataRow == null)
            //{
            //    return;
            //}
            //Init_Currency_Code("N");
            //Init_Currency_Amount();
            //GetSubForm();
            //if (SLIP_QUERY_STATUS.EditValue.ToString() != "QUERY".ToString())
            //{
            //    Init_DR_CR_Amount();
            //}
            //Init_Total_GL_Amount();

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

        private void idaDPR_ASSET_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
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


        #region ----- Doc Att / Appr. Step -----

        private void BTN_CLOSED_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
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
            F_SLIP_DATE.EditValue = SLIP_DATE.EditValue;
            F_SLIP_NUM.EditValue = SLIP_NUM.EditValue;

            if (iString.ISNull(F_SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10221"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Init_Sub_Panel(true, "DOC_ATT");

            //FTP 정보//
            Set_FTP_Info();

            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();
        }

        private void BTN_DOC_ATT_L_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //전표 목록에서 파일 보기.
            F_SLIP_DATE.EditValue = igrSLIP_LIST.GetCellValue("SLIP_DATE");
            F_SLIP_NUM.EditValue = igrSLIP_LIST.GetCellValue("SLIP_NUM");

            if (iString.ISNull(F_SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Init_Sub_Panel(true, "DOC_ATT");

            //FTP 정보//
            Set_FTP_Info();

            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();
        }

        private void IGR_DOC_ATTACHMENT_CellDoubleClick(object pSender)
        {
            //if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            //{
            //    return;
            //}

            //string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            //string vUSER_FILE_NAME = string.Format("{0}{1}", mDownload_Folder, IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            //if (DownLoadFile(vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            //{
            //    return;
            //} 
        }

        private void BTN_ATT_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_GET_DOC_ATT_STATUS.SetCommandParamValue("P_SOURCE_CATEGORY", "SLIP");
            IDC_GET_DOC_ATT_STATUS.ExecuteNonQuery();
            String vSTATUS = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_STATUS"));
            String vMESSAGE = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            UpLoadFile(F_SLIP_DATE.EditValue, F_SLIP_NUM.EditValue);
            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_ATT_DOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID");
            string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            string vUSER_FILE_NAME = string.Format("{0}{1}", mDownload_Folder, IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            if (DownLoadFile(vDOC_ATTACHMENT_ID, vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }
        }

        private void BTN_ATT_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10220"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            if (iString.ISNull(F_SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_GET_DOC_ATT_STATUS.SetCommandParamValue("P_SOURCE_CATEGORY", "SLIP");
            IDC_GET_DOC_ATT_STATUS.ExecuteNonQuery();
            String vSTATUS = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_STATUS"));
            String vMESSAGE = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID");
            string vFTP_FileName = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            DeleteFile(vDOC_ATTACHMENT_ID, vFTP_FileName);
            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();
        }

        private void BTN_ATT_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DOC_ATT_FLAG();
            Init_Sub_Panel(false, "DOC_ATT");
        }

        private void BTN_APPR_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            A_SLIP_DATE.EditValue = SLIP_DATE.EditValue;
            A_SLIP_NUM.EditValue = SLIP_NUM.EditValue;
            IDA_APPROVAL_PERSON.Fill();
            Init_Sub_Panel(true, "APPR_STEP");
        }

        #region ----- FTP Infomation -----
        //ftp 접속정보 및 환경 정보 설정 
        private void Set_FTP_Info()
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            mFTP_Connect_Status = false;
            mHost = string.Empty;
            mPort = string.Empty;
            mPassive = "N";
            mUserID = string.Empty;
            mPassword = string.Empty;
            mFTP_Folder = string.Empty;
            mClient_Folder = String.Empty;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "SLIP_DOC");
                IDC_FTP_INFO.ExecuteNonQuery();
                if (IDC_FTP_INFO.ExcuteError)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }
                mHost = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mPort = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mUserID = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_NO"));
                mPassword = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mPassive = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));
                mFTP_Folder = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mClient_Folder = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            if (mHost == string.Empty)
            {
                //ftp접속정보 오류          
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            try
            {
                //FileTransfer Initialze
                mFileTransfer = new ISFileTransferAdv();
                mFileTransfer.Host = mHost;
                mFileTransfer.Port = mPort;
                if (mPassive == "Y")
                {
                    mFileTransfer.UsePassive = true;
                }
                else
                {
                    mFileTransfer.UsePassive = false;
                }
                mFileTransfer.UserId = mUserID;
                mFileTransfer.Password = mPassword;

                mDownload_Folder = string.Format("{0}{1}", mClient_Base_Path, mClient_Folder.Replace("/", "\\"));
            }
            catch (System.Exception Ex)
            {
                //ftp접속정보 오류 
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            //Client Download Folder 없으면 생성 
            System.IO.DirectoryInfo vDownload_Folder = new System.IO.DirectoryInfo(mDownload_Folder);
            if (vDownload_Folder.Exists == false) //있으면 True, 없으면 False
            {
                vDownload_Folder.Create();
            }
            else
            {
                //기존 파일 삭제//
                try
                {
                    string vDate = DateTime.Today.ToString("yyyyMMdd");
                    System.IO.FileInfo[] vFiles = vDownload_Folder.GetFiles();
                    foreach (System.IO.FileInfo rFile in vFiles)
                    {
                        if (vDate.CompareTo(rFile.LastWriteTime.ToString("yyyyMMdd")) > 0)
                        {
                            try
                            {
                                System.IO.File.Delete(vDownload_Folder + "\\" + rFile);
                            }
                            catch
                            {
                                //
                            }
                        }
                    }
                }
                catch
                {
                    //
                }
            }

            mFTP_Connect_Status = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- File Upload Methods -----
        //ftp에 file upload 처리 
        private bool UpLoadFile(object pSLIP_DATE, object pSLIP_NUM)
        {
            bool isUpload = false;
            OpenFileDialog vOpenFileDialog1 = new OpenFileDialog();
            vOpenFileDialog1.RestoreDirectory = true;

            if (mFTP_Connect_Status == false)
            {
                isAppInterfaceAdv1.OnAppMessage("FTP Server Connect Fail. Check FTP Server");
                return isUpload;
            }

            if (iString.ISNull(pSLIP_NUM) != string.Empty)
            {
                string vSTATUS = "F";
                string vMESSAGE = string.Empty;

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                vOpenFileDialog1.Title = "Select Open File";
                vOpenFileDialog1.Filter = "All File(*.*)|*.*|pdf File(*.pdf)|*.pdf|jpg file(*.jpg)|*.jpg|bmp file(*.bmp)|*.bmp";
                vOpenFileDialog1.DefaultExt = "*.pdf";
                vOpenFileDialog1.FileName = "";
                vOpenFileDialog1.Multiselect = true;


                if (vOpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSelectFullPath = string.Empty;
                    string vSelectDirectoryPath = string.Empty;

                    string vFileName = string.Empty;
                    string vFileExtension = string.Empty;

                    //1. 사용자 선택 파일 
                    for (int i = 0; i < vOpenFileDialog1.FileNames.Length; i++)
                    {
                        vSelectFullPath = vOpenFileDialog1.FileNames[i];
                        vSelectDirectoryPath = System.IO.Path.GetDirectoryName(vSelectFullPath);

                        vFileName = System.IO.Path.GetFileName(vSelectFullPath);
                        vFileExtension = System.IO.Path.GetExtension(vSelectFullPath).ToUpper();

                        //2. 첨부파일 DB 저장 
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_CATEGORY", "SLIP_DOC"); //구분 
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_DATE", pSLIP_DATE);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_NUM", pSLIP_NUM);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_USER_FILE_NAME", vFileName);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_FTP_FILE_NAME", vFileName);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_EXTENSION_NAME", vFileExtension);
                        IDC_INSERT_DOC_ATTACHMENT.ExecuteNonQuery();

                        vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));
                        object vDOC_ATTACHMENT_ID = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_DOC_ATTACHMENT_ID");
                        object vFTP_FILE_NAME = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_FTP_FILE_NAME");

                        //O_DOC_ATTACHMENT_ID.EditValue = vDOC_ATTACHMENT_ID;
                        //O_FTP_FILE_NAME.EditValue = vFTP_FILE_NAME;

                        if (IDC_INSERT_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();

                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return isUpload;
                        }

                        //3. 첨부파일 로그 저장 
                        IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
                        IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                        IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
                        vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return isUpload;
                        }

                        //4. 파일 업로드
                        try
                        {
                            mFileTransfer.ShowProgress = true;      //진행바 보이기 

                            //업로드 환경 설정 
                            mFileTransfer.SourceDirectory = vSelectDirectoryPath;
                            mFileTransfer.SourceFileName = vFileName;
                            mFileTransfer.TargetDirectory = mFTP_Folder;
                            mFileTransfer.TargetFileName = iString.ISNull(vFTP_FILE_NAME);

                            bool isUpLoad = mFileTransfer.Upload();

                            if (isUpLoad == true)
                            {
                                isUpload = true;
                            }
                            else
                            {
                                isUpload = false;
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                            //5. 적용 
                        }
                        catch (Exception Ex)
                        {
                            isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                            return isUpload;
                        }
                    }
                }
            }
            return isUpload;
        }

        #endregion;


        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(object pDOC_ATTACHMENT_ID, string pFTP_FileName, string pClient_FileName)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();            
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", pDOC_ATTACHMENT_ID);
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return IsDownload;
            }

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFTP_FileName);
            try
            {
                System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                if (vDownFileInfo.Exists == true)
                {
                    try
                    {
                        System.IO.File.Delete(vTempFileName);
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

            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------
            mFileTransfer.SourceDirectory = mFTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FileName;
            mFileTransfer.TargetDirectory = mDownload_Folder;
            mFileTransfer.TargetFileName = vTempFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mDownload_Folder, vTempFileName);      //임시

                    System.IO.File.Delete(pClient_FileName);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, pClient_FileName);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
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
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
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
            if (IsDownload == true)
            {
                System.Diagnostics.Process.Start(pClient_FileName);
            }
            else
            {
                string vMessage = string.Format("{0} {1}", isMessageAdapter1.ReturnText("EAPP_10212"), isMessageAdapter1.ReturnText("QM_10102"));
                MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;

        #region ----- file Delete Methods -----
        //ftp file delete 처리 
        private bool DeleteFile(object pDOC_ATTACHMENT_ID, string pFTP_FileName)
        {
            bool IsDelete = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            if (iString.ISNull(pDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }
            if (pFTP_FileName == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();


            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리  
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", pDOC_ATTACHMENT_ID);
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "DELETE");
            IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return IsDelete;
            }

            //2. 파일 삭제 
            IDC_DELETE_DOC_ATTACHMENT.SetCommandParamValue("W_DOC_ATTACHMENT_ID", pDOC_ATTACHMENT_ID);
            IDC_DELETE_DOC_ATTACHMENT.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
            {
                IsDelete = false;
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();


                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return IsDelete;
            }

            //3. 실제 삭제 
            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FileName;
            mFileTransfer.TargetDirectory = mFTP_Folder;
            mFileTransfer.TargetFileName = pFTP_FileName;

            IsDelete = mFileTransfer.Delete();
            if (IsDelete == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                return IsDelete;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return IsDelete;
        }

        #endregion; 

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

        private void GB_DOC_ATT_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_DOC_ATT_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_DOC_ATT.Location;
                I.Offset(gx, gy);
                GB_DOC_ATT.Location = I;
            }
        }

        private void GB_DOC_ATT_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_AP_VAT_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_AP_VAT_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_AP_VAT.Location;
                I.Offset(gx, gy);
                GB_AP_VAT.Location = I;
            }
        }

        private void GB_AP_VAT_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_COPY_DOCUMENT_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_COPY_DOCUMENT_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_COPY_DOCUMENT.Location;
                I.Offset(gx, gy);
                GB_COPY_DOCUMENT.Location = I;
            }
        }

        private void GB_COPY_DOCUMENT_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void BTN_INQUIRY_APPR_PERSON_LIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.Fill();
        }

        #endregion

    }
}