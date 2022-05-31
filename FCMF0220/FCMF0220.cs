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

namespace FCMF0220
{
    public partial class FCMF0220 : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

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
        string mAuto_Search_Flag = "N"; 
         
        #endregion;

        #region ----- Constructor -----

        public FCMF0220()
        {
            InitializeComponent();
        }

        public FCMF0220(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            mAuto_Search_Flag = "N";
        }

        public FCMF0220(Form pMainForm, ISAppInterface pAppInterface, object pGL_Date_FR, object pGL_Date_TO, object pSlip_Num)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iString.ISNull(pSlip_Num) != string.Empty)
            {
                SLIP_DATE_FR_0.EditValue = pGL_Date_FR;
                SLIP_DATE_TO_0.EditValue = pGL_Date_TO;
                SLIP_NUM_0.EditValue = pSlip_Num;
                irbCONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;

                mAuto_Search_Flag = "Y";
            }
        }

        public FCMF0220(Form pMainForm, ISAppInterface pAppInterface, object pGL_Date_FR, object pGL_Date_TO,
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
                    irbCONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
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
                Search_DB_DETAIL(H_HEADER_INTERFACE_ID.EditValue);
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
                
                igrSLIP_LIST.LastConfirmChanges();
                idaSLIP_HEADER_IF_LIST.OraSelectData.AcceptChanges();
                idaSLIP_HEADER_IF_LIST.Refillable = true;

                string vGL_NUM = iString.ISNull(igrSLIP_LIST.GetCellValue("SLIP_NUM"));
                int vCOL_IDX = igrSLIP_LIST.GetColumnToIndex("SLIP_NUM");
                idaSLIP_HEADER_IF_LIST.Fill();
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
                SLIP_QUERY_STATUS.EditValue = "QUERY"; 
                itbSLIP.SelectedIndex = 1;
                itbSLIP.SelectedTab.Focus();
                 
                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;

                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;

                Application.DoEvents();
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

                //데이터 그리드 초기화.
                idaSLIP_HEADER.SetSelectParamValue("P_SOB_ID", -1);
                idaSLIP_HEADER.Fill();
                INIT_MANAGEMENT_COLUMN();
                SET_GRID_COL_VISIBLE(pSLIP_HEADER_ID);  // 그리드 보이기/감추기 설정.

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                idaSLIP_HEADER.SetSelectParamValue("W_HEADER_INTERFACE_ID", pSLIP_HEADER_ID);
                idaSLIP_HEADER.SetSelectParamValue("P_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
                idaSLIP_HEADER.Fill();
                  
                //첨부파일 여부.//
                DOC_ATT_FLAG();
                 
                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;

                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;

                GL_DATE_L.Focus();
            }
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            idaMANAGEMENT_PROMPT.Fill();
            if (idaMANAGEMENT_PROMPT.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            // Adapter Column.
            int mIDX_Column;            // 시작 COLUMN.
            int mMax_Column = idaMANAGEMENT_PROMPT.SelectColumns.Count - 10; // 종료 COLUMN.
            object mCOLUMN_DESC;        // 헤더 프롬프트.
            
            //Grid Column.
            int mGrid_Column = 11;     // 그리드 시작 Column.
            for (mIDX_Column = 1; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = idaMANAGEMENT_PROMPT.CurrentRow[mIDX_Column];
                if (iString.ISNull(mCOLUMN_DESC, ":=") == ":=".ToString())
                {
                    igrSLIP_LINE.GridAdvExColElement[mGrid_Column].Visible = 0;
                }
                else
                {
                    igrSLIP_LINE.GridAdvExColElement[mGrid_Column].Visible = 1;
                    igrSLIP_LINE.GridAdvExColElement[mGrid_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    igrSLIP_LINE.GridAdvExColElement[mGrid_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
                mGrid_Column = mGrid_Column + 1;
            }
            igrSLIP_LINE.ResetDraw = true;
        }


        private void SET_GRID_COL_VISIBLE(object pSLIP_HEADER_ID)
        { 
            idaSLIP_MANAGEMENT_YN.SetSelectParamValue("W_HEADER_INTERFACE_ID", pSLIP_HEADER_ID);
            idaSLIP_MANAGEMENT_YN.Fill();

            // Adapter Column.
            int mIDX_Column;            // 시작 COLUMN.
            int mMax_Column = 0;        // 종료 COLUMN.
            int mGrid_Column = 11;      // 그리드 시작 Column.
            object mVISIBLE_YN = ":=";   // 보이기 여부.

            if (idaSLIP_MANAGEMENT_YN.OraSelectData.Rows.Count == 0)
            {
                // Adapter Column.
                mMax_Column = idaMANAGEMENT_PROMPT.SelectColumns.Count - 2; // 종료 COLUMN.
                for (mIDX_Column = 1; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mVISIBLE_YN = idaMANAGEMENT_PROMPT.CurrentRow[mIDX_Column];
                    if (iString.ISNull(mVISIBLE_YN, ":=") == ":=".ToString())
                    {
                        igrSLIP_LINE.GridAdvExColElement[mGrid_Column].Visible = 0;
                    }
                    else
                    {
                        igrSLIP_LINE.GridAdvExColElement[mGrid_Column].Visible = 1;
                    }
                    mGrid_Column = mGrid_Column + 1;
                }
            }
            else
            {
                // Adapter Column.
                mMax_Column = idaSLIP_MANAGEMENT_YN.SelectColumns.Count - 2; // 종료 COLUMN.
                for (mIDX_Column = 1; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mVISIBLE_YN = idaSLIP_MANAGEMENT_YN.CurrentRow[mIDX_Column];
                    if (iString.ISNull(mVISIBLE_YN, ":=") == ":=".ToString())
                    {
                        igrSLIP_LINE.GridAdvExColElement[mGrid_Column].Visible = 0;
                    }
                    else
                    {
                        igrSLIP_LINE.GridAdvExColElement[mGrid_Column].Visible = 1;
                    }
                    mGrid_Column = mGrid_Column + 1;
                }
            }
            igrSLIP_LINE.ResetDraw = true;
        }


        private void Set_Grid_Select(ISGridAdvEx pGrid, object pGROUPING_STATUS)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("SELECT_CHECK_YN");
            pGrid.GridAdvExColElement[vIDX_CHECK].Insertable = pGROUPING_STATUS;
            pGrid.GridAdvExColElement[vIDX_CHECK].Updatable = pGROUPING_STATUS;
        }

        private void Set_Grid_Status(ISGridAdvEx pGrid, object pGROUPING_STATUS)
        {
            //int vIDX_CHECK = pGrid.GetColumnToIndex("GL_DATE");
            //pGrid.GridAdvExColElement[vIDX_CHECK].Insertable = pGROUPING_STATUS;
            //pGrid.GridAdvExColElement[vIDX_CHECK].Updatable = pGROUPING_STATUS;
        }

        private void Set_Slip_List_Position(object pSlip_Num)
        {
            int vIDX_SELECT_CHECK_YN = igrSLIP_LIST.GetColumnToIndex("SELECT_CHECK_YN");
            int vIDX_SLIP_NUM = igrSLIP_LIST.GetColumnToIndex("SLIP_NUM");
            
            // 기존 위치 이동 : 없을 경우.
            for (int r = 0; r < igrSLIP_LIST.RowCount; r++)
            {
                if (iString.ISNull(pSlip_Num) == iString.ISNull(igrSLIP_LIST.GetCellValue(r, vIDX_SLIP_NUM)))
                {
                    igrSLIP_LIST.CurrentCellMoveTo(r, vIDX_SELECT_CHECK_YN);
                    igrSLIP_LIST.CurrentCellActivate(r, vIDX_SELECT_CHECK_YN);
                }
            }
            igrSLIP_LIST.Focus();
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
           
        private void GetSubForm()
        {
            ibtSUB_FORM.Visible = false;
            ACCOUNT_CLASS_YN.EditValue = null;
            ACCOUNT_CLASS_TYPE.EditValue = null;
            string vBTN_CAPTION = null;
            
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("ACCOUNT_CONTROL_ID")) == string.Empty || iString.ISNull(igrSLIP_LINE.GetCellValue("ACCOUNT_DR_CR")) == string.Empty)
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
            ibtSUB_FORM.Top = 102;
            ibtSUB_FORM.ButtonTextElement[0].Default = vBTN_CAPTION;
            ibtSUB_FORM.BringToFront();
            ibtSUB_FORM.Visible = true;
            ibtSUB_FORM.TabStop = true;
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
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Left = 65;
                        GB_APPR.Top = 115;

                        GB_APPR.Width = 900;
                        GB_APPR.Height = 240;

                        GB_APPR.Border3DStyle = Border3DStyle.Bump;
                        GB_APPR.BorderStyle = BorderStyle.Fixed3D;

                        GB_APPR.Controls[0].MouseDown += GB_APPR_MouseDown;
                        GB_APPR.Controls[0].MouseMove += GB_APPR_MouseMove;
                        GB_APPR.Controls[0].MouseUp += GB_APPR_MouseUp;
                        GB_APPR.Controls[1].MouseDown += GB_APPR_MouseDown;
                        GB_APPR.Controls[1].MouseMove += GB_APPR_MouseMove;
                        GB_APPR.Controls[1].MouseUp += GB_APPR_MouseUp;

                        GB_APPR.BringToFront();
                        GB_APPR.Visible = true;
                    } 
                    else if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Left = 278;
                        GB_RETURN.Top = 89;

                        GB_RETURN.Width = 600;
                        GB_RETURN.Height = 200;

                        GB_RETURN.Border3DStyle = Border3DStyle.Bump;
                        GB_RETURN.BorderStyle = BorderStyle.Fixed3D;

                        //GroupBox 이동// 
                        GB_RETURN.Controls[0].MouseDown += GB_RETURN_MouseDown;
                        GB_RETURN.Controls[0].MouseMove += GB_RETURN_MouseMove;
                        GB_RETURN.Controls[0].MouseUp += GB_RETURN_MouseUp;
                        GB_RETURN.Controls[1].MouseDown += GB_RETURN_MouseDown;
                        GB_RETURN.Controls[1].MouseMove += GB_RETURN_MouseMove;
                        GB_RETURN.Controls[1].MouseUp += GB_RETURN_MouseUp;

                        //값 초기화.
                        V_RETURN_REMARK.EditValue = string.Empty;
                        GB_RETURN.BringToFront();
                        GB_RETURN.Visible = true;
                    }
                    else if (pSub_Panel == "APPROVAL")
                    {
                        GB_APPROVAL.Left = 278;
                        GB_APPROVAL.Top = 89;

                        GB_APPROVAL.Width = 600;
                        GB_APPROVAL.Height = 200;

                        GB_APPROVAL.Border3DStyle = Border3DStyle.Bump;
                        GB_APPROVAL.BorderStyle = BorderStyle.Fixed3D;

                        //GroupBox 이동// 
                        GB_APPROVAL.Controls[0].MouseDown += GB_APPROVAL_MouseDown;
                        GB_APPROVAL.Controls[0].MouseMove += GB_APPROVAL_MouseMove;
                        GB_APPROVAL.Controls[0].MouseUp += GB_APPROVAL_MouseUp;
                        GB_APPROVAL.Controls[1].MouseDown += GB_APPROVAL_MouseDown;
                        GB_APPROVAL.Controls[1].MouseMove += GB_APPROVAL_MouseMove;
                        GB_APPROVAL.Controls[1].MouseUp += GB_APPROVAL_MouseUp;

                        //값 초기화.
                        V_APPROVAL_DESCRIPTION.EditValue = string.Empty;
                        GB_APPROVAL.BringToFront();
                        GB_APPROVAL.Visible = true;
                    }
                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                itpSLIP_LIST.Enabled = false;
                igbSLIP_HEADER.Enabled = false;
                igrSLIP_LINE.Enabled = false;
                igbCONFIRM_INFOMATION.Enabled = false;
                igbACCOUNT_LINE.Enabled = false; 
                GB_APPR_STEP.Enabled = false;
            }
            else
            {
                try
                {
                    if (pSub_Panel == "ALL")
                    {
                        GB_AP_VAT.Visible = false;  
                        GB_APPR_STEP.Enabled = false;
                        GB_APPR.Visible = false;
                        GB_RETURN.Visible = false;
                        GB_APPROVAL.Visible = false;
                    }
                    else if (pSub_Panel == "AP_VAT")
                    {
                        GB_AP_VAT.Visible = false;
                    }
                    else if (pSub_Panel == "APPR_STEP")
                    {
                        GB_APPR.Visible = false;
                    } 
                    else if (pSub_Panel == "RETURN")
                    {
                        GB_RETURN.Visible = false;
                    }
                    else if (pSub_Panel == "APPROVAL")
                    {
                        GB_APPROVAL.Visible = false;
                    }
                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                itpSLIP_LIST.Enabled = true;
                igbSLIP_HEADER.Enabled = true;
                igrSLIP_LINE.Enabled = true;
                igbCONFIRM_INFOMATION.Enabled = true;
                igbACCOUNT_LINE.Enabled = true; 
                GB_APPR_STEP.Enabled = true;
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


        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID)
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

                            object[] vParam = new object[5];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = SLIP_DATE.EditValue;     //기표일자 시작
                            vParam[3] = SLIP_DATE.EditValue;     //기표일자 종료
                            vParam[4] = SLIP_NUM.EditValue;      //기표번호 

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
                            vParam[5] = "Y";            //읽기 전용 여부

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
                    catch (Exception Ex)
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
         
        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                //else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                //{
                //    if (idaSLIP_LINE.IsFocused)
                //    {
                //        idaSLIP_LINE.AddOver();
                //        InsertSlipLine();
                //    }                    
                //}
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaSLIP_HEADER_IF_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_IF_LIST.Cancel();
                    }
                    else if (idaSLIP_HEADER.IsFocused)
                    {
                        idaSLIP_LINE.Cancel();
                        idaSLIP_HEADER.Cancel();
                    }
                    else if (idaSLIP_LINE.IsFocused)
                    {
                        idaSLIP_LINE.Cancel(); 
                    }
                    else if (idaSLIP_HEADER_IF_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_IF_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (idaSLIP_HEADER_IF_LIST.IsFocused)
                    //{
                    //    if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    //    {
                    //        return;
                    //    }
                    //    idaSLIP_HEADER_IF_LIST.Cancel();
                    //    idaSLIP_HEADER_IF_LIST.Delete();
                    //    idaSLIP_HEADER_IF_LIST.Update();
                    //}
                    //else if (idaSLIP_HEADER.IsFocused)
                    //{                        
                    //    if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    //    {
                    //        return;
                    //    }
                         
                    //    IDC_DELETE_SLIP.SetCommandParamValue("W_HEADER_INTERFACE_ID", H_HEADER_INTERFACE_ID.EditValue);
                    //    IDC_DELETE_SLIP.SetCommandParamValue("W_WITH_TEMP_SLIP", "N");
                    //    IDC_DELETE_SLIP.ExecuteNonQuery();

                    //    Search_DB();
                    //    Search_DB_DETAIL(DBNull.Value); 
                    //} 
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
        
        private void FCMF0220_Load(object sender, EventArgs e)
        {
            SLIP_QUERY_STATUS.EditValue = "NON-QUERY";
            if (mAuto_Search_Flag == "N")
            {
                SLIP_DATE_FR_0.EditValue = iDate.ISMonth_1st(iDate.ISDate_Month_Add(iDate.ISGetDate(), -2));
                SLIP_DATE_TO_0.EditValue = iDate.ISGetDate();
            }

            // 회계장부 정보 설정.
            GetAccountBook();

            idaSLIP_HEADER_IF_LIST.FillSchema();
            idaSLIP_HEADER.FillSchema(); 
        }

        private void FCMF0220_Shown(object sender, EventArgs e)
        {     
            //서브판넬 
            Init_Sub_Panel(false, "ALL"); 

            //서브 화면
            ibtSUB_FORM.Visible = false;
            BTN_VIEW_TEMP_SLIP.BringToFront();  
            REF_SLIP_FLAG.BringToFront();
            V_CB_ALL.BringToFront();
            BTN_INQUIRY_APPR_PERSON_LIST.BringToFront();

            irbCONFIRM_NO.CheckedState = ISUtil.Enum.CheckedState.Checked; 
            ibtSUB_FORM.Visible = false;

            Application.DoEvents();
            if (mAuto_Search_Flag == "L")
            {
                Search_DB();
            }
            else if (mAuto_Search_Flag == "Y")
            {
                Search_DB();
                if (igrSLIP_LIST.RowCount > 0)
                {
                    Search_DB_DETAIL(igrSLIP_LIST.GetCellValue("SLIP_HEADER_ID"));
                }
            } 
        }

        private void igrSLIP_LIST_CellDoubleClick(object pSender)
        {
            if (igrSLIP_LIST.RowCount > 0)
            {
                Search_DB_DETAIL(igrSLIP_LIST.GetCellValue("HEADER_INTERFACE_ID"));
            }
        }
          
        private void igrSLIP_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == igrSLIP_LIST.GetColumnToIndex("SELECT_CHECK_YN"))
            {
                object vCONFIRM_YN = igrSLIP_LIST.GetCellValue("CONFIRM_YN");
                int vIDX_GL_DATE = igrSLIP_LIST.GetColumnToIndex("GL_DATE");

                 
            }

            igrSLIP_LIST.LastConfirmChanges();
            idaSLIP_HEADER_IF_LIST.OraSelectData.AcceptChanges();
            idaSLIP_HEADER_IF_LIST.Refillable = true;
        }

        private void C_GL_DATE_EditValueChanged(object pSender)
        {
            idaCONFIRM_STATUS.OraSelectData.AcceptChanges();
            idaCONFIRM_STATUS.Refillable = true;
        }
         
        private void btnCONFIRM_YES_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            }

            //서브판넬 
            Init_Sub_Panel(true, "APPROVAL");
        }

        private void C_BTN_EXEC_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_DATE.Focus();
                return;
            }
            if (iString.ISNull(SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            } 

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_APPROVAL_PERSON.SetCommandParamValue("P_APPROVAL_DESCRIPTION", V_APPROVAL_DESCRIPTION.EditValue);
            IDC_EXEC_APPROVAL_PERSON.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_APPROVAL_PERSON.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_APPROVAL_PERSON.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_EXEC_APPROVAL_PERSON.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            Init_Sub_Panel(false, "APPROVAL");
            Search_DB_DETAIL(H_HEADER_INTERFACE_ID.EditValue); 
        }

        private void C_BTN_EXEC_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "APPROVAL");
        }

        private void btnCONFIRM_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idcSLIP_CONFIRM_YN.ExecuteNonQuery();
            object vCONFIRM_YN = idcSLIP_CONFIRM_YN.GetCommandParamValue("O_CONFIRM_YN");
            if (iString.ISNull(vCONFIRM_YN) == "Y")
            {//마감 처리됨.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10115"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (iString.ISNull(vCONFIRM_YN) == "R")
            {//반려 처리됨.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10135"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //서브판넬 
            Init_Sub_Panel(true, "RETURN"); 
        }

        private void C_BTN_EXEC_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_RETURN_REMARK.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_RETURN_REMARK))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_RETURN_APPROVAL_PERSON.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_RETURN_APPROVAL_PERSON.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_RETURN_APPROVAL_PERSON.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_RETURN_APPROVAL_PERSON.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //서브판넬 
            Init_Sub_Panel(false, "RETURN");
            Search_DB_DETAIL(H_HEADER_INTERFACE_ID.EditValue);
        }

        private void C_BTN_RETURN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //서브판넬 
            Init_Sub_Panel(false, "RETURN");
        }

        private void btnCONFIRM_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idcSLIP_CONFIRM_YN.ExecuteNonQuery();
            object vCONFIRM_YN = idcSLIP_CONFIRM_YN.GetCommandParamValue("O_CONFIRM_YN");
            //if (iString.ISNull(vCONFIRM_YN) == "Y")
            //{//미승인상태 -> 취소 데이터 없음
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10137"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            
            if (iString.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_DATE.Focus();
                return;
            }
            if (iString.ISNull(SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            }
            if (iString.ISNull(H_HEADER_INTERFACE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_APPROVAL_PERSON.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_APPROVAL_PERSON.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_APPROVAL_PERSON.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_APPROVAL_PERSON.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Search_DB_DETAIL(H_HEADER_INTERFACE_ID.EditValue);
        }
          
        private void irbCONFIRM_Status_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            
            CONFIRM_STATUS_0.EditValue = iStatus.RadioCheckedString;
            Search_DB();

            //Set_Grid_Select(igrSLIP_LIST, "1"); 
            //if (iString.ISNull(CONFIRM_STATUS_0.EditValue) == "N")
            //{ 
            //    Set_Grid_Status(igrSLIP_LIST, "1");
            //}
            //else if (iString.ISNull(CONFIRM_STATUS_0.EditValue) == "Y")
            //{ 
            //    Set_Grid_Status(igrSLIP_LIST, "0");
            //}
            //else if (iString.ISNull(CONFIRM_STATUS_0.EditValue) == "R")
            //{ 
            //    Set_Grid_Status(igrSLIP_LIST, "0");
            //}
            //else if (iString.ISNull(CONFIRM_STATUS_0.EditValue) == String.Empty)
            //{ 
            //    Set_Grid_Select(igrSLIP_LIST, "0");
            //    Set_Grid_Status(igrSLIP_LIST, "0");
            //}
        }
          
        private void BTN_VIEW_TEMP_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(SLIP_NUM.EditValue) != string.Empty)
            {            
                AssmblyRun_Manual("FCMF0206");
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
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
            //서브판넬 
            Init_Sub_Panel(false, "AP_VAT"); 
        } 
         
        #endregion

        #region ----- Lookup Event ----- 


        private void ilaSLIP_NUM_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildSLIP_NUM.SetLookupParamValue("W_SLIP_NUM", SLIP_NUM_0.EditValue);
        }

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SLIP_TYPE_TEMP_DOCU.SetLookupParamValue("P_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_ACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }
          
        private void ilaVAT_ASSET_GB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_ASSET_GB", "Y");
        }

        #endregion       

        #region ----- Adapter Event -----

        private void idaSLIP_HEADER_IF_LIST_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            string vCONFIRM_YN = iString.ISNull(CONFIRM_STATUS_0.EditValue);
            if (vCONFIRM_YN == string.Empty)
            {
                Set_Grid_Select(igrSLIP_LIST, "0");
                Set_Grid_Status(igrSLIP_LIST, "0");
            }
            else if (iString.ISNull(pBindingManager.DataRow["CONFIRM_YN"]) == "Y")
            {
                Set_Grid_Status(igrSLIP_LIST, "0");
            }
            else if (iString.ISNull(pBindingManager.DataRow["CONFIRM_YN"]) == "R")
            {
                Set_Grid_Status(igrSLIP_LIST, "0");
            }
            else
            {
                Set_Grid_Status(igrSLIP_LIST, "1");
            }
        }

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
                if (e.Row["CONFIRM_YN"].ToString() == "Y".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }
         
        private void idaSLIP_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                if (e.Row["CONFIRM_YN"].ToString() == "Y".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10052"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }
         
        private void idaSLIP_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            } 
            GetSubForm(); 
        }
         
        private void idaSLIP_HEADER_FillCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            } 
        }

        private void idaSLIP_HEADER_UpdateCompleted_1(object pSender)
        {
            string vGL_NUM = iString.ISNull(SLIP_NUM.EditValue); // igrSLIP_LIST.GetCellValue("GL_NUM"));
            int vIDX_GL_NUM = igrSLIP_LIST.GetColumnToIndex("SLIP_NUM");
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

        private void IDA_APPROVAL_PERSON_UpdateCompleted(object pSender)
        {
            if (IDA_APPROVAL_PERSON.UpdateModifiedRowCount != 0)
            {
                mSave_Appr_Status = true;
            }
        }

        private void BTN_INSERT_A_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.AddUnder();
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
                    if (vRow.RowState != DataRowState.Unchanged)
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
            if (iString.ISNull(SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            AssmblyRun_Attachment("FCMF0228", SLIP_DATE.EditValue, SLIP_NUM.EditValue); 
        }

        private void BTN_DOC_ATT_L_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(igrSLIP_LIST.GetCellValue("SLIP_NUM")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            AssmblyRun_Attachment("FCMF0228", igrSLIP_LIST.GetCellValue("SLIP_DATE"), igrSLIP_LIST.GetCellValue("SLIP_NUM")); 
        }
          
        private void BTN_APPR_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        {             
            A_SLIP_DATE.EditValue = SLIP_DATE.EditValue;
            A_SLIP_NUM.EditValue = SLIP_NUM.EditValue;
            A_DEPT_ID.EditValue = H_BUDGET_DEPT_ID.EditValue;
            IDA_APPROVAL_PERSON.Fill();
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
         
        private void GB_APPROVAL_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_APPROVAL_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_APPROVAL.Location;
                I.Offset(gx, gy);
                GB_APPROVAL.Location = I;
            }
        }

        private void GB_APPROVAL_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void GB_RETURN_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_RETURN_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_RETURN.Location;
                I.Offset(gx, gy);
                GB_RETURN.Location = I;
            }
        }

        private void GB_RETURN_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }

        private void BTN_INQUIRY_APPR_PERSON_LIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON_LIST.Fill();
        }

        #endregion

    }
}