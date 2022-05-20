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

namespace FCMF0522
{
    public partial class FCMF0522 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0522()
        {
            InitializeComponent();
        }

        public FCMF0522(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(W_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BALANCE_DATE.Focus();
                return;
            }
            //Set_Tab_Focus();

            if (TB_BALANCE_STATEMENT.SelectedTab.TabIndex == TP_BALANCE_STATEMENT.TabIndex)
            {
                if (iString.ISNull(W_ACCOUNT_CODE_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_ACCOUNT_CODE_FR.Focus();
                    return;
                }
                if (iString.ISNull(W_ACCOUNT_CODE_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_ACCOUNT_CODE_TO.Focus();
                    return;
                }

                //adater 초기화//
                IGR_BALANCE_STATEMENT.LastConfirmChanges();
                IDA_BALANCE_STATEMENT.OraSelectData.AcceptChanges();
                IDA_BALANCE_STATEMENT.Refillable = true;

                int vIDX_Col = IGR_BALANCE_ACCOUNT.GetColumnToIndex("ACCOUNT_CONTROL_ID");
                decimal vACCOUNT_CONTROL_ID = iString.ISDecimaltoZero(IGR_BALANCE_ACCOUNT.GetCellValue("ACCOUNT_CONTROL_ID"), 0);
                
                IDA_BALANCE_ACCOUNT.Fill();
                
                //focus 이동.
                if (vACCOUNT_CONTROL_ID <= 0)
                {
                    //
                }
                for (int nRow = 0; nRow < IGR_BALANCE_ACCOUNT.RowCount; nRow++)
                {
                    if (vACCOUNT_CONTROL_ID == Convert.ToInt32(iString.ISDecimaltoZero(IGR_BALANCE_ACCOUNT.GetCellValue(nRow, vIDX_Col), 0)))
                    {
                        IGR_BALANCE_ACCOUNT.CurrentCellMoveTo(nRow, 1);
                        IGR_BALANCE_ACCOUNT.CurrentCellActivate(nRow, 1); 
                        return;
                    }
                }
                IGR_BALANCE_ACCOUNT.Focus();

                Init_DUE_DATE(); 
            }
            else if (TB_BALANCE_STATEMENT.SelectedTab.TabIndex == TP_STATEMENT_SLIP.TabIndex)
            {
                IDA_BALANCE_STATEMENT_DTL.Fill();
                IGR_BALANCE_STATEMENT_DTL.Focus();
            }
        }

        private void Set_Tab_Focus(object pACCOUNT_CONTROL_ID)
        {
            if(TB_BALANCE_STATEMENT.SelectedTab.TabIndex == TP_BALANCE_STATEMENT.TabIndex)
            {
                int vIDX_Col = IGR_BALANCE_STATEMENT.GetColumnToIndex("ITEM_GROUP_ID"); 

                if (iString.ISNull(pACCOUNT_CONTROL_ID) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                IDA_BALANCE_STATEMENT.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", -1);
                IDA_BALANCE_STATEMENT.Fill();

                INIT_MANAGEMENT_COLUMN(pACCOUNT_CONTROL_ID);
                Application.DoEvents();

                IDA_BALANCE_STATEMENT.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
                IDA_BALANCE_STATEMENT.Fill();
            }
            else if (TB_BALANCE_STATEMENT.SelectedTab.TabIndex == TP_STATEMENT_SLIP.TabIndex)
            {
                IDA_BALANCE_STATEMENT_DTL.Fill();
                IGR_BALANCE_STATEMENT_DTL.Focus();
            }
        }

        private void INIT_EDIT_TYPE()
        {
            W_MANAGEMENT_DESC.EditValue = null;
            W_MANAGEMENT_DESC.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            W_MANAGEMENT_DESC.NumberDecimalDigits = 0;
            if (iString.ISNull(W_DATA_TYPE.EditValue) == "NUMBER".ToString())
            {
                W_MANAGEMENT_DESC.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(W_DATA_TYPE.EditValue) == "RATE".ToString())
            {
                W_MANAGEMENT_DESC.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                W_MANAGEMENT_DESC.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(W_DATA_TYPE.EditValue) == "DATE".ToString())
            {
                W_MANAGEMENT_DESC.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }

            if (iString.ISNull(W_LOOKUP_YN.EditValue, "N") == "N")
            {
                W_MANAGEMENT_DESC.LookupAdapter = null;
            }
            else
            {
                W_MANAGEMENT_DESC.LookupAdapter = ILA_MANAGEMENT_ITEM;
            }
            W_MANAGEMENT_DESC.Refresh();
        }

        private void INIT_MANAGEMENT_COLUMN(object pACCOUNT_CONTROL_ID)
        {
            IDA_ITEM_PROMPT.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_ITEM_PROMPT.Fill();

            int mStart_Column = 4;
            int mIDX_Column;          // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            if (IDA_ITEM_PROMPT.OraSelectData.Rows.Count == 0)
            {
                for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mENABLED_COLUMN = mMax_Column + mIDX_Column;
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0; 
                }
                IGR_BALANCE_STATEMENT.ResetDraw = true;
                return;
            }
            
            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = IDA_ITEM_PROMPT.CurrentRow[mENABLED_COLUMN];

                if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                }
            }

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = IDA_ITEM_PROMPT.CurrentRow[mIDX_Column];
                if (iString.ISNull(mCOLUMN_DESC) != string.Empty)
                {
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = mCOLUMN_DESC.ToString();
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = mCOLUMN_DESC.ToString();
                }
            }

            // 전표일자 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }

            // 전표번호 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_NUM");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }

            // 적요.
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("REMARK");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 1;
            }


            // 외화금액 - 통화관리 하는 경우 적용.
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("CURR_GL_AMOUNT");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["CURR_CONTROL_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {                
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Insertable = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Updatable = 1;
            }

            // 환산환율 적용 - 환산환율 관리 하는 경우 적용.
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("NEW_EXCHANGE_RATE");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["CURR_ESTIMATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {   
                // 환산환율.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                //환산원화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 0;
                //환산손익.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 2].Visible = 0;
            }
            else
            {
                // 환산환율.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                //환산원화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 1;
                //환산손익.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 2].Visible = 1;
            }
            IGR_BALANCE_STATEMENT.ResetDraw = true;
        }
        
        private void SET_GRID_COL_STATUS(DataRow pDATA_ROW)
        {
            if (pDATA_ROW == null)
            {
                return;
            }
            int mIDX_CURR_GL_AMOUNT = IGR_BALANCE_STATEMENT.GetColumnToIndex("CURR_GL_AMOUNT");
            int mIDX_GL_AMOUNT = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_AMOUNT");
            int mIDX_DESCRIPTION = IGR_BALANCE_STATEMENT.GetColumnToIndex("DESCRIPTION");

            if (iString.ISNull(pDATA_ROW["SUMMARY_FLAG"], "T") == "N")
            {
                //외화금액.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_CURR_GL_AMOUNT].Insertable = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_CURR_GL_AMOUNT].Updatable = 1;
                //원화금액.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_GL_AMOUNT].Insertable = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_GL_AMOUNT].Updatable = 1;
                //비고.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_DESCRIPTION].Insertable = 1;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 1;
            }
            else
            {
                //외화금액.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_CURR_GL_AMOUNT].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_CURR_GL_AMOUNT].Updatable = 0;
                //원화금액.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_GL_AMOUNT].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_GL_AMOUNT].Updatable = 0;
                //비고.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_DESCRIPTION].Insertable = 0;
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 0;
            }
        }
         
        private void Show_Slip_Detail()
        {
            int mSLIP_HEADER_ID = iString.ISNumtoZero(IGR_BALANCE_STATEMENT_DTL.GetCellValue("SLIP_HEADER_ID"));
            if (mSLIP_HEADER_ID != Convert.ToInt32(0))
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
                vFCMF0204.Show();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.UseWaitCursor = false;
            }
        }

        private void Init_DUE_DATE()
        {
            object mENABLED_YN = "N";
            IDC_MANAGEMENT_ENABLED_YN.SetCommandParamValue("P_LOOKUP_TYPE", "DUE_DATE");
            IDC_MANAGEMENT_ENABLED_YN.ExecuteNonQuery();
            mENABLED_YN = IDC_MANAGEMENT_ENABLED_YN.GetCommandParamValue("O_ENABLED_YN");
            if (iString.ISNull(mENABLED_YN) == "Y")
            {
                W_GROUPING_DUE_DATE.Visible = true;
            }
            else
            {
                W_GROUPING_DUE_DATE.Visible = false;
            }
            W_GROUPING_DUE_DATE.CheckBoxValue = "N";
            W_GROUPING_DUE_DATE.Invalidate();
        }

        private void Insert_Balance_Statement()
        {
            object vBALANCE_DATE = BALANCE_DATE.EditValue;
            if (iString.ISNull(vBALANCE_DATE) == String.Empty)
            {
                vBALANCE_DATE = W_BALANCE_DATE.EditValue;
            }
            if (iString.ISNull(vBALANCE_DATE) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BALANCE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object vACCOUNT_CONTROL_ID = IGR_BALANCE_ACCOUNT.GetCellValue("ACCOUNT_CONTROL_ID");
            object vACCOUNT_CODE = IGR_BALANCE_ACCOUNT.GetCellValue("ACCOUNT_CODE");
            object vACCOUNT_DESC = IGR_BALANCE_ACCOUNT.GetCellValue("ACCOUNT_DESC");
            //if (iString.ISNull(W_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_ACCOUNT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}            

            //FCMF0522_ADD vFCMF0522_ADD = new FCMF0522_ADD(isAppInterfaceAdv1.AppInterface, vBALANCE_DATE, 
            //                                                W_ACCOUNT_CONTROL_ID.EditValue, W_ACCOUNT_CODE_FR.EditValue, W_ACCOUNT_DESC_FR.EditValue);
            FCMF0522_ADD vFCMF0522_ADD = new FCMF0522_ADD(isAppInterfaceAdv1.AppInterface, vBALANCE_DATE,
                                                            vACCOUNT_CONTROL_ID, vACCOUNT_CODE, vACCOUNT_DESC);
            DialogResult = vFCMF0522_ADD.ShowDialog();
            if (DialogResult == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0522_ADD.Dispose();
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

        #region ----- XL Print 1 (계정잔액명세서) Method ----

        private void XLPrinting_1(string pOutChoice)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            object vBALANCE_DATE = iDate.ISGetDate(BALANCE_DATE.EditValue).ToShortDateString();
            object vACCOUNT_CODE = W_ACCOUNT_CODE_FR.EditValue;
            object vACCOUNT_DESC = W_ACCOUNT_DESC_FR.EditValue;
            object vTerritory = string.Empty;
            object vGROUPING_OPTION = null;
            
            if (iString.ISNull(vBALANCE_DATE) == String.Empty)
            {//기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (iString.ISNull(vACCOUNT_CODE) == String.Empty)
            {//계정과목코드
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int vCountRow = IGR_BALANCE_STATEMENT.RowCount;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = string.Format("Balance_{0}_{1}", vBALANCE_DATE, vACCOUNT_DESC);

            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.Filter = "Excel file(*.xlsx)|*.xlsx";
            saveFileDialog1.DefaultExt = "xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = saveFileDialog1.FileName;
                System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                try
                {
                    if (vFileName.Exists)
                    {
                        vFileName.Delete();
                    }
                }
                catch (Exception EX)
                {
                    MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            vTerritory = GetTerritory();
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0522_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //조회시 그룹핑 옵션.
                    if (iString.ISNull(W_GROUPING_DUE_DATE.CheckBoxValue) == "Y")
                    {
                        switch (iString.ISNull(vTerritory))
                        {
                            case "TL1_KR":
                                vGROUPING_OPTION = string.Format("내역서({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL1_KR);
                                break;
                            case "TL2_CN":
                                vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL2_CN);
                                break;
                            case "TL3_VN":
                                vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL3_VN);
                                break;
                            case "TL4_JP":
                                vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL4_JP);
                                break;
                            case "TL5_XAA":
                                vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].TL5_XAA);
                                break;
                            default:                                
                                vGROUPING_OPTION = string.Format("Detailed Statement({0})", W_GROUPING_DUE_DATE.PromptTextElement[0].Default);
                                break;
                        }
                    }
                    else
                    {
                        switch (iString.ISNull(vTerritory))
                        {
                            case "TL1_KR":
                                vGROUPING_OPTION = "내역서";
                                break;
                            case "TL2_CN":
                                vGROUPING_OPTION = "Detailed Statement";
                                break;
                            case "TL3_VN":
                                vGROUPING_OPTION = "Detailed Statement";
                                break;
                            case "TL4_JP":
                                vGROUPING_OPTION = "Detailed Statement";
                                break;
                            case "TL5_XAA":
                                vGROUPING_OPTION = "Detailed Statement";
                                break;
                            default:                                
                                vGROUPING_OPTION = "Detailed Statement";
                                break;
                        }
                    }
                    //날짜형식 변경.
                    IDC_DATE_YYYYMMDD.SetCommandParamValue("P_DATE", vBALANCE_DATE);
                    IDC_DATE_YYYYMMDD.ExecuteNonQuery();
                    vBALANCE_DATE = IDC_DATE_YYYYMMDD.GetCommandParamValue("O_DATE");
                    xlPrinting.HeaderWrite(vACCOUNT_CODE, vACCOUNT_DESC, vBALANCE_DATE, vGROUPING_OPTION, iString.ISNull(vTerritory), IGR_BALANCE_STATEMENT);

                    // 실제 인쇄
                    //vPageNumber = xlPrinting.LineWrite(vBALANCE_DATE, iString.ISNull(vTerritory), pGRID);
                    vPageNumber = xlPrinting.LineWrite(IGR_BALANCE_STATEMENT);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {

                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = "Printing End";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

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
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_BALANCE_STATEMENT.IsFocused)
                    {
                        Insert_Balance_Statement();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BALANCE_STATEMENT.IsFocused)
                    {
                        Insert_Balance_Statement();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    int vREC_COUNT = IDA_BALANCE_STATEMENT.DeletedRowCount ;
                    if (vREC_COUNT > 0)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    } 
                    IDA_BALANCE_STATEMENT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_BALANCE_STATEMENT.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BALANCE_STATEMENT.IsFocused)
                    {
                        IDA_BALANCE_STATEMENT.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0522_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;

            int vIDX_BS_CONFIRM_FLAG = IGR_BALANCE_STATEMENT.GetColumnToIndex("CONFIRM_FLAG");
            int vIDX_BSD_CONFIRM_FLAG = IGR_BALANCE_STATEMENT_DTL.GetColumnToIndex("CONFIRM_FLAG");

            W_BALANCE_DATE.EditValue = DateTime.Today;
            W_BALANCE_DATE_FR.EditValue = iDate.ISMonth_1st(W_BALANCE_DATE.EditValue);
            W_BALANCE_DATE_TO.EditValue = W_BALANCE_DATE.EditValue;

            if (mCONFIRM_CHECK == "Y")
            {
                GB_CONFIRM_STATUS.Visible = true;

                IGR_BALANCE_STATEMENT.GridAdvExColElement[vIDX_BS_CONFIRM_FLAG].Visible = 1;
                IGR_BALANCE_STATEMENT_DTL.GridAdvExColElement[vIDX_BSD_CONFIRM_FLAG].Visible = 1; 
            }
            else
            {
                GB_CONFIRM_STATUS.Visible = false;

                IGR_BALANCE_STATEMENT.GridAdvExColElement[vIDX_BS_CONFIRM_FLAG].Visible = 0;
                IGR_BALANCE_STATEMENT_DTL.GridAdvExColElement[vIDX_BSD_CONFIRM_FLAG].Visible = 0; 
            }

            IGR_BALANCE_STATEMENT.ResetDraw = true;
            IGR_BALANCE_STATEMENT_DTL.ResetDraw = true;
        }

        private void FCMF0522_Shown(object sender, EventArgs e)
        {           
            IDA_BALANCE_STATEMENT.FillSchema();  
            Init_DUE_DATE();
        }

        private void V_RB_CONFIRM_ALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;

            V_CONFIRM_STATUS.EditValue = iStatus.RadioCheckedString;
        }

        private void igrSTATEMENT_SLIP_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        private void IGR_BALANCE_STATEMENT_CellDoubleClick(object pSender)
        {
            if (IGR_BALANCE_STATEMENT.RowCount > 0)
            {                
                TB_BALANCE_STATEMENT.SelectedIndex = 1;
                TB_BALANCE_STATEMENT.Focus();

                Set_Tab_Focus(IGR_BALANCE_STATEMENT.GetCellValue("ACCOUNT_CONTROL_ID"));
            }
        }

        private void btnEXE_STATEMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BALANCE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0522_SET vFCMF0522_SET = new FCMF0522_SET(isAppInterfaceAdv1.AppInterface, W_BALANCE_DATE.EditValue);
            vRESULT = vFCMF0522_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0522_SET.Dispose();
        }

        private void btnCLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BALANCE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0522_CLOSED vFCMF0522_CLOSED = new FCMF0522_CLOSED(isAppInterfaceAdv1.AppInterface, BALANCE_DATE.EditValue);
            vRESULT = vFCMF0522_CLOSED.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0522_CLOSED.Dispose();
        }
                
        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BALANCE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0522_CANCEL vFCMF0522_CANCEL = new FCMF0522_CANCEL(isAppInterfaceAdv1.AppInterface, BALANCE_DATE.EditValue);
            vRESULT = vFCMF0522_CANCEL.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0522_CANCEL.Dispose();
        }

        private void btnEXE_FORWARD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BALANCE_DATE.Focus();
                return;
            }

            String mMESSAGE;
            idcSET_CARRY_FORWARD.ExecuteNonQuery();
            mMESSAGE = iString.ISNull(idcSET_CARRY_FORWARD.GetCommandParamValue("O_MESSAGE"));
            if (mMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
         
        private void BTN_BSD_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SearchDB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", null);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_TO_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", W_ACCOUNT_CODE_FR.EditValue);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_ACCOUNT_CODE_FR_SelectedRowData(object pSender)
        {
            W_ACCOUNT_CODE_TO.EditValue = W_ACCOUNT_CODE_FR.EditValue;
            W_ACCOUNT_DESC_TO.EditValue = W_ACCOUNT_DESC_FR.EditValue;
        }

        private void ilaMANAGEMENT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_MANAGEMENT_TYPE.SetLookupParamValue("W_ACCOUNT_CONTROL_ID", IGR_BALANCE_ACCOUNT.GetCellValue("ACCOUNT_CONTROL_ID"));
            ILD_MANAGEMENT_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT_TYPE_SelectedRowData(object pSender)
        {
            INIT_EDIT_TYPE();
        }

        private void ilaMANAGEMENT_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_MANAGEMENT_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----
        
        private void idaBALANCE_STATEMENT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BALANCE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["ITEM_GROUP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Item Group ID(관리항목 그룹 ID)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaBALANCE_STATEMENT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["BALANCE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["ITEM_GROUP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Item Group ID(관리항목 그룹 ID)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaBALANCE_STATEMENT_UpdateCompleted(object pSender)
        {
            SearchDB();
        }

        private void idaBALANCE_STATEMENT_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            SET_GRID_COL_STATUS(pBindingManager.DataRow);
        }

        private void idaSTATEMENT_EXCHANGE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BALANCE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaSTATEMENT_EXCHANGE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["BALANCE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_BALANCE_ACCOUNT_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Set_Tab_Focus(-1);
                return;
            }

            Set_Tab_Focus(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
        }

        #endregion


    }
}