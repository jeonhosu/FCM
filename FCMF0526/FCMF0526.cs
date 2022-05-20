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

namespace FCMF0526
{
    public partial class FCMF0526 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0526()
        {
            InitializeComponent();
        }

        public FCMF0526(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch 
            {
                vDateTime = DateTime.Today;
            }
            return vDateTime;
        } 

        private void SEARCH_DB()
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_FR.Focus();
                return;
            }

            string mAccount_Code = iString.ISNull(IGR_BATCH_BILL_ACCOUNT.GetCellValue("ACCOUNT_CODE"));
            int mIDX_Account_Code = IGR_BATCH_BILL_ACCOUNT.GetColumnToIndex("ACCOUNT_CODE");
            IDA_BATCH_BILL_ACCOUNT.Fill();
            IDA_BATCH_BILL_SELECTED.Fill();
            if (mAccount_Code != string.Empty)
            {
                for (int r = 0; r < IGR_BATCH_BILL_ACCOUNT.RowCount; r++)
                {
                    if (mAccount_Code == iString.ISNull(IGR_BATCH_BILL_ACCOUNT.GetCellValue(r, mIDX_Account_Code)))
                    {
                        IGR_BATCH_BILL_ACCOUNT.CurrentCellMoveTo(r, mIDX_Account_Code);
                        IGR_BATCH_BILL_ACCOUNT.CurrentCellActivate(r, mIDX_Account_Code);
                    }
                }
            }
        }

        private void SEARCH_DETAIL(DataRow pDataRow)
        {
            // 대량지급 내역 조회.
            IDA_BATCH_BILL_SELECTED.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pDataRow["ACCOUNT_CONTROL_ID"]);
            IDA_BATCH_BILL_SELECTED.SetSelectParamValue("W_GL_DATE", pDataRow["GL_DATE"]);
            IDA_BATCH_BILL_SELECTED.Fill();
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            IDA_ITEM_PROMPT.Fill();
            if (IDA_ITEM_PROMPT.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            int mStart_Column = 4;
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = IDA_ITEM_PROMPT.CurrentRow[mENABLED_COLUMN];
                mCOLUMN_DESC = IDA_ITEM_PROMPT.CurrentRow[mIDX_Column];
                if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                    IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }

            //// 전표일자 표시
            //mIDX_Column = 0;
            //mIDX_Column = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("GL_DATE");
            //mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            //if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            //{
            //    IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            //}
            //else
            //{
            //    IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            //}

            // 적요.
            mIDX_Column = 0;
            mIDX_Column = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("SLIP_REMARK");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 외화금액 - 통화관리 하는 경우 적용.
            mIDX_Column = 0;
            mIDX_Column = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("GL_CURR_AMOUNT");
            mENABLED_FLAG = iString.ISNull(IDA_ITEM_PROMPT.CurrentRow["CURR_CONTROL_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Insertable = 1;
                IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_Column].Updatable = 1;
            }
            IGR_BATCH_BILL_SELECTED.ResetDraw = true;
        }

        private void Set_Selected_Total_Amount()
        {
            decimal mTotal_Curr_Amount = 0;
            decimal mTotal_Amount = 0;
            int mIDX_GL_CURR_AMOUNT = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("GL_CURR_AMOUNT");
            int mIDX_GL_AMOUNT = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("GL_AMOUNT");

            for (int i = 0; i < IGR_BATCH_BILL_SELECTED.RowCount; i++)
            {
                mTotal_Curr_Amount = iString.ISDecimaltoZero(mTotal_Curr_Amount, 0) +
                                        iString.ISDecimaltoZero(IGR_BATCH_BILL_SELECTED.GetCellValue(i, mIDX_GL_CURR_AMOUNT), 0);

                mTotal_Amount = iString.ISDecimaltoZero(mTotal_Amount, 0) +
                                    iString.ISDecimaltoZero(IGR_BATCH_BILL_SELECTED.GetCellValue(i, mIDX_GL_AMOUNT), 0);
                
            }
            TOTAL_CURR_AMOUNT.EditValue = mTotal_Curr_Amount;
            TOTAL_AMOUNT.EditValue = mTotal_Amount;
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

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;
            object vGL_DATE = iDate.ISGetDate(V_DUE_DATE_FR.EditValue).ToShortDateString();

            if (iString.ISNull(vGL_DATE) == String.Empty)
            {//기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 데이터 조회.
            IDA_PRINT_BATCH_BILL.Fill();
            vCountRow = IDA_PRINT_BATCH_BILL.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("Bill_{0}", vGL_DATE);

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
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;
                    
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
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {   
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0526_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite1(vGL_DATE, IDA_PRINT_BATCH_BILL);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
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
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_BATCH_BILL_SELECTED.IsFocused)
                    {
                       // IDA_BATCH_BILL_ACCOUNT.Update();
                        IDA_BATCH_BILL_SELECTED.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BATCH_BILL_SELECTED.IsFocused)
                    {
                        IDA_BATCH_BILL_SELECTED.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BATCH_BILL_SELECTED.IsFocused)
                    {
                        if (iString.ISNull(IGR_BATCH_BILL_SELECTED.GetCellValue("SUMMARY_FLAG")) != "N")
                        {
                            return;
                        }
                        IDA_BATCH_BILL_SELECTED.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting1("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void FCMF0526_Load(object sender, EventArgs e)
        {
            IDA_BATCH_BILL_ACCOUNT.FillSchema();
        }

        private void FCMF0526_Shown(object sender, EventArgs e)
        {
            V_DUE_DATE_FR.EditValue = GetDate();
            V_DUE_DATE_TO.EditValue = V_DUE_DATE_FR.EditValue;

            V_GL_DATE.EditValue = V_DUE_DATE_TO.EditValue;
        }

        private void BTN_CREATE_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(V_DUE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_TO.Focus();
                return;
            }
            if (iString.ISNull(V_GL_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_GL_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0526_SET vFCMF0526_SET = new FCMF0526_SET(isAppInterfaceAdv1.AppInterface, V_DUE_DATE_FR.EditValue, V_DUE_DATE_TO.EditValue, V_GL_DATE.EditValue);
            vRESULT = vFCMF0526_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SEARCH_DB();
            }
            vFCMF0526_SET.Dispose();
        }

        private void BTN_CONFIRM_Y_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_FR.Focus();
                return;
            }

            if (IDA_BATCH_BILL_SELECTED.ModifiedRowCount != 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
                        
            string mSTATUS = "F";
            string mMESSAGE = null;

            IDC_BILL_CONFIRM_Y.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_BILL_CONFIRM_Y.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_BILL_CONFIRM_Y.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_BILL_CONFIRM_Y.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            if (mMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            SEARCH_DB();
        }

        private void BTN_CONFIRM_N_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_FR.Focus();
                return;
            }

            if (IDA_BATCH_BILL_SELECTED.ModifiedRowCount != 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDA_BATCH_BILL_SELECTED.Cancel();
             
            string mSTATUS = "F";
            string mMESSAGE = null;

            IDC_BILL_CONFIRM_N.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_BILL_CONFIRM_N.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_BILL_CONFIRM_N.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_BILL_CONFIRM_N.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 

            if (mMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            SEARCH_DB();
        }

        #endregion

        #region ----- Lookup Event -----


        #endregion

        #region ----- Adapter event -----
        
        private void IDA_PAYMENT_SELECTED_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_BATCH_PAYMENT_ACCOUNT_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            SEARCH_DETAIL(pBindingManager.DataRow);
        }

        private void IDA_PAYMENT_SELECTED_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            int mEdit_Flag = 0;
            int mIDX_GL_AMOUNT = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("GL_AMOUNT");
            int mIDX_GL_CURR_AMOUNT = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("GL_CURR_AMOUNT");
            int mIDX_EXCHANGE_RATE = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("EXCHANGE_RATE");
            int mIDX_REMARK = IGR_BATCH_BILL_SELECTED.GetColumnToIndex("REMARK"); 

            if (iString.ISNull(pBindingManager.DataRow["CONFIRM_YN"]) == "Y")
            {
                mEdit_Flag = 0;
            }
            else if (iString.ISNull(pBindingManager.DataRow["SUMMARY_FLAG"]) == "N")
            {
                mEdit_Flag = 1;
            }
            else
            {
                mEdit_Flag = 0;
            }
            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_GL_AMOUNT].Insertable = mEdit_Flag;
            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_GL_AMOUNT].Updatable = mEdit_Flag;

            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_GL_CURR_AMOUNT].Insertable = mEdit_Flag;
            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_GL_CURR_AMOUNT].Updatable = mEdit_Flag;

            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_EXCHANGE_RATE].Insertable = mEdit_Flag;
            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_EXCHANGE_RATE].Updatable = mEdit_Flag;

            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_REMARK].Insertable = mEdit_Flag;
            IGR_BATCH_BILL_SELECTED.GridAdvExColElement[mIDX_REMARK].Updatable = mEdit_Flag;

            IGR_BATCH_BILL_SELECTED.Refresh();
        }

        #endregion

    }
}