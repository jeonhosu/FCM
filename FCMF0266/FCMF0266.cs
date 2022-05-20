using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0266
{
    public partial class FCMF0266 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mCurrency_Code;
        object mBudget_Control_YN;

        object mSlip_Num;
        object mSlip_Header_ID;
        object mSlip_Date;
        
        #endregion;

        #region ----- Constructor -----

        public FCMF0266()
        {
            InitializeComponent();
        }

        public FCMF0266(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void GetAccountBook()
        {
            IDC_ACCOUNT_BOOK.ExecuteNonQuery();
            mSession_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_SESSION_ID");
            mAccount_Book_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = iConv.ISNull(IDC_ACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private void Sync_BTN_Status(string pSLIP_FLAG)
        {
            if(pSLIP_FLAG == "N")
            {
                BTN_CANCEL_SLIP.Enabled = false;
                BTN_SET_SLIP.Enabled = true;
            }
            else if(pSLIP_FLAG == "Y")
            {
                BTN_SET_SLIP.Enabled = false;
                BTN_CANCEL_SLIP.Enabled = true;
            }
            else
            {
                BTN_SET_SLIP.Enabled = false;
                BTN_CANCEL_SLIP.Enabled = false;
            }
        }

        private void Search_DB()
        {
            if (iConv.ISNull(W_USE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_USE_DATE_FR.Focus();
                return;
            }

            if (iConv.ISNull(W_USE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_USE_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(W_USE_DATE_FR.EditValue) > Convert.ToDateTime(W_USE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_USE_DATE_FR.Focus();
                return;
            }

            IGR_CREDIT_APPR.LastConfirmChanges();
            IDA_CARD_APPROVAL.OraSelectData.AcceptChanges();
            IDA_CARD_APPROVAL.Refillable = true;
            IGR_CREDIT_APPR.ResetDraw = true;

            string vAPPR_NUM = iConv.ISNull(IGR_CREDIT_APPR.GetCellValue("APPROVAL_NUM"));
            int vCOL_IDX = IGR_CREDIT_APPR.GetColumnToIndex("APPROVAL_NUM");
            IDA_CARD_APPROVAL.Fill();
            Sum_Amount();

            if (iConv.ISNull(vAPPR_NUM) != string.Empty)
            {
                for (int i = 0; i < IGR_CREDIT_APPR.RowCount; i++)
                {
                    if (vAPPR_NUM == iConv.ISNull(IGR_CREDIT_APPR.GetCellValue(i, vCOL_IDX)))
                    {
                        IGR_CREDIT_APPR.CurrentCellMoveTo(i, vCOL_IDX);
                        IGR_CREDIT_APPR.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
        }

        private void Sum_Amount()
        {
            decimal vUSE_AMOUNT = 0;
            decimal vVAT_AMOUNT = 0;
            decimal vTOTAL_AMOUNT = 0;

            foreach (DataRow vROW in IDA_CARD_APPROVAL.CurrentRows)
            {
                //IDA_CARD_ACQUIRE
                vUSE_AMOUNT = vUSE_AMOUNT + iConv.ISDecimaltoZero(vROW["BASE_APPR_AMOUNT"]);
                vVAT_AMOUNT = vVAT_AMOUNT + iConv.ISDecimaltoZero(vROW["BASE_VAT_AMOUNT"]);
                vTOTAL_AMOUNT = vTOTAL_AMOUNT + iConv.ISDecimaltoZero(vROW["BASE_TOTAL_AMOUNT"]);
            }
            V_BASE_APPR_AMOUNT.EditValue = vUSE_AMOUNT;
            V_BASE_VAT_AMOUNT.EditValue = vVAT_AMOUNT;
            V_BASE_TOTAL_AMOUNT.EditValue = vTOTAL_AMOUNT;
        }

        private void Init_Select_YN(string pStatus)
        {
            int vIDX_SELECT_YN = IGR_CREDIT_APPR.GetColumnToIndex("SELECT_YN");
            int vIDX_SLIP_DATE = IGR_CREDIT_APPR.GetColumnToIndex("SLIP_DATE");
            int vIDX_USE_PERSON_NAME = IGR_CREDIT_APPR.GetColumnToIndex("USE_PERSON_NAME");
            int vIDX_BUDGET_DEPT_NAME = IGR_CREDIT_APPR.GetColumnToIndex("BUDGET_DEPT_NAME");
            //int vIDX_EXP_ACCOUNT_CODE = IGR_CREDIT_APPR.GetColumnToIndex("EXP_ACCOUNT_CODE");
            //int vIDX_VAT_ACCOUNT_CODE = IGR_CREDIT_APPR.GetColumnToIndex("VAT_ACCOUNT_CODE");
            //int vIDX_REMARK = IGR_CREDIT_APPR.GetColumnToIndex("REMARK");

            //if (pStatus == "N")
            //{
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SLIP_DATE].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SLIP_DATE].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_USE_PERSON_NAME].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_USE_PERSON_NAME].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_BUDGET_DEPT_NAME].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_BUDGET_DEPT_NAME].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXP_ACCOUNT_CODE].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXP_ACCOUNT_CODE].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_VAT_ACCOUNT_CODE].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_VAT_ACCOUNT_CODE].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_REMARK].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_REMARK].Updatable = 1;
            //}
            //else if (pStatus == "Y")
            //{
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 1;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 1;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SLIP_DATE].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SLIP_DATE].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_USE_PERSON_NAME].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_USE_PERSON_NAME].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_BUDGET_DEPT_NAME].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_BUDGET_DEPT_NAME].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXP_ACCOUNT_CODE].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXP_ACCOUNT_CODE].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_VAT_ACCOUNT_CODE].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_VAT_ACCOUNT_CODE].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_REMARK].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_REMARK].Updatable = 0;
            //}
            //else
            //{
            //    IDA_CARD_APPROVAL.Cancel();

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SLIP_DATE].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_SLIP_DATE].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_USE_PERSON_NAME].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_USE_PERSON_NAME].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_BUDGET_DEPT_NAME].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_BUDGET_DEPT_NAME].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXP_ACCOUNT_CODE].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXP_ACCOUNT_CODE].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_VAT_ACCOUNT_CODE].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_VAT_ACCOUNT_CODE].Updatable = 0;

            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_REMARK].Insertable = 0;
            //    IGR_CREDIT_APPR.GridAdvExColElement[vIDX_REMARK].Updatable = 0;
            //}
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
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            if (iConv.ISNull(mSlip_Header_ID) == string.Empty)
            {
                return;
            }

            AssmblyRun_Manual("FCMF0212", mSlip_Header_ID, mSlip_Date, mSlip_Num); 

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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_CARD_APPROVAL.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CARD_APPROVAL.IsFocused)
                    {
                        IDA_CARD_APPROVAL.Cancel();
                        IDA_CARD_SLIP_LIST.Cancel();
                    }
                    else
                    { 
                        IDA_CARD_SLIP_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    mSlip_Header_ID = IGR_CREDIT_APPR.GetCellValue("SLIP_HEADER_ID");
                    mSlip_Date = IGR_CREDIT_APPR.GetCellValue("SLIP_DATE");
                    mSlip_Num = IGR_CREDIT_APPR.GetCellValue("SLIP_NUM");
                    XLPrinting_Main("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    mSlip_Header_ID = IGR_CREDIT_APPR.GetCellValue("SLIP_HEADER_ID");
                    mSlip_Date = IGR_CREDIT_APPR.GetCellValue("SLIP_DATE");
                    mSlip_Num = IGR_CREDIT_APPR.GetCellValue("SLIP_NUM");
                    XLPrinting_Main("EXCEL");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0266_Load(object sender, EventArgs e)
        {
            W_ABROAD_FLAG.BringToFront();
            W_CANCEL_FLAG.BringToFront();
            BTN_SLIP_DELETE.BringToFront();

            V_RB_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SLIP_IF_FLAG.EditValue = V_RB_NO.RadioCheckedString;

            W_USE_DATE_FR.EditValue = iDate.ISMonth_1st(iDate.ISGetDate());
            W_USE_DATE_TO.EditValue = iDate.ISGetDate(); 
        }

        private void FCMF0266_Shown(object sender, EventArgs e)
        {
            IDA_CARD_APPROVAL.FillSchema();

            //회계 장부//
            GetAccountBook();
        }

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            if (V_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_ALL.RadioCheckedString; 
                Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
                Sync_BTN_Status("ALL");
            }
        }

        private void V_RB_YES_Click(object sender, EventArgs e)
        {
            if (V_RB_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_YES.RadioCheckedString;
                Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
                Sync_BTN_Status("Y");
            }
        }

        private void V_RB_NO_Click(object sender, EventArgs e)
        {
            if (V_RB_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_NO.RadioCheckedString;
                Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue));
                Sync_BTN_Status("N");
            }
        }
          
        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {                          
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            int vREC_CNT = 0;

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vSYS_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE"));

            int vIDX_SELECT_YN = IGR_CREDIT_APPR.GetColumnToIndex("SELECT_YN");
            int vIDX_CARD_APPROVAL_ID = IGR_CREDIT_APPR.GetColumnToIndex("CARD_APPROVAL_ID");
            int vIDX_CARD_NUM = IGR_CREDIT_APPR.GetColumnToIndex("CARD_NUM");
            int vIDX_APPROVAL_NUM = IGR_CREDIT_APPR.GetColumnToIndex("APPROVAL_NUM");

            int vIDX_BASE_APPR_AMOUNT = IGR_CREDIT_APPR.GetColumnToIndex("BASE_APPR_AMOUNT");
            int vIDX_BASE_VAT_AMOUNT = IGR_CREDIT_APPR.GetColumnToIndex("BASE_VAT_AMOUNT");
            int vIDX_BASE_TOTAL_AMOUNT = IGR_CREDIT_APPR.GetColumnToIndex("BASE_TOTAL_AMOUNT");
            int vIDX_PART_CANCEL_FLAG = IGR_CREDIT_APPR.GetColumnToIndex("PART_CANCEL_FLAG");

            for (int r = 0; r < IGR_CREDIT_APPR.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_CREDIT_APPR.GetCellValue(r, vIDX_SELECT_YN)))
                {
                    vREC_CNT++;
                    IGR_CREDIT_APPR.CurrentCellMoveTo(r, vIDX_SELECT_YN);
                    IGR_CREDIT_APPR.CurrentCellActivate(r, vIDX_SELECT_YN);

                    //if (iConv.ISNull(IGR_CREDIT_APPR.GetCellValue("USE_PERSON_ID")) == string.Empty)
                    //{
                    //    Application.UseWaitCursor = false;
                    //    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    //    Application.DoEvents();

                    //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10037"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //    return;
                    //} 

                    //전표 일괄생성 패키지에서 처리하므로 폼에서는 제어할수 없음
                    //해당 체크박스는 숨김 처리
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_CARD_APPROVAL_ID", IGR_CREDIT_APPR.GetCellValue(r, vIDX_CARD_APPROVAL_ID));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_SYS_DATE", vSYS_DATE);
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_CARD_NUM", IGR_CREDIT_APPR.GetCellValue(r, vIDX_CARD_NUM));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_APPROVAL_NUM", IGR_CREDIT_APPR.GetCellValue(r, vIDX_APPROVAL_NUM));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_USE_AMT", IGR_CREDIT_APPR.GetCellValue(r, vIDX_BASE_APPR_AMOUNT));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_VAT_AMT", IGR_CREDIT_APPR.GetCellValue(r, vIDX_BASE_VAT_AMOUNT));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_TOT_AMT", IGR_CREDIT_APPR.GetCellValue(r, vIDX_BASE_TOTAL_AMOUNT));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_PART_CANCEL_FLAG", IGR_CREDIT_APPR.GetCellValue(r, vIDX_PART_CANCEL_FLAG));
                    IDC_SAVE_CARD_SLIP_GROUP.SetCommandParamValue("P_SESSION_ID", mSession_ID);
                    IDC_SAVE_CARD_SLIP_GROUP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SAVE_CARD_SLIP_GROUP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SAVE_CARD_SLIP_GROUP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SAVE_CARD_SLIP_GROUP.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_SAVE_CARD_SLIP_GROUP.ExcuteErrorMsg, "Slip Create-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Slip Create-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    } 
                    //IGR_CREDIT_APPR.SetCellValue(r, vIDX_SELECT_YN, "N");

                    IGR_CREDIT_APPR.LastConfirmChanges();
                    IDA_CARD_APPROVAL.OraSelectData.AcceptChanges();
                    IDA_CARD_APPROVAL.Refillable = true;
                }
            }
            if (vREC_CNT == 0)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                return;
            }

            //전표 생성 화면//
            FCMF0266_SLIP vFCMF0266_SLIP = new FCMF0266_SLIP(this.MdiParent, isAppInterfaceAdv1.AppInterface, iConv.ISNull(W_SLIP_IF_FLAG.EditValue), mSession_ID, vSYS_DATE);
            DialogResult dlgResult = vFCMF0266_SLIP.ShowDialog(); 

            mSlip_Header_ID = vFCMF0266_SLIP.Get_Slip_Header_ID;
            mSlip_Num = vFCMF0266_SLIP.Get_Slip_Num;
            mSlip_Date = vFCMF0266_SLIP.Get_Slip_Date;

            vFCMF0266_SLIP.Close();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();
            if (dlgResult != DialogResult.OK)
            { 
                return;
            }

            //전표인쇄//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10146"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            XLPrinting_Main("PRINT");
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vSYS_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE"));

            int vIDX_SELECT_YN = IGR_CREDIT_APPR.GetColumnToIndex("SELECT_YN");
            int vIDX_CARD_APPROVAL_ID = IGR_CREDIT_APPR.GetColumnToIndex("CARD_APPROVAL_ID");
            int vIDX_CARD_NUM = IGR_CREDIT_APPR.GetColumnToIndex("CARD_NUM");
            int vIDX_APPROVAL_NUM = IGR_CREDIT_APPR.GetColumnToIndex("APPROVAL_NUM");

            for (int r = 0; r < IGR_CREDIT_APPR.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_CREDIT_APPR.GetCellValue(r, vIDX_SELECT_YN)))
                {
                    IGR_CREDIT_APPR.CurrentCellMoveTo(r, vIDX_SELECT_YN);
                    IGR_CREDIT_APPR.CurrentCellActivate(r, vIDX_SELECT_YN);

                    if (iConv.ISNull(IGR_CREDIT_APPR.GetCellValue("SLIP_HEADER_ID")) == string.Empty)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10037"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
   
                    //전표 일괄생성 패키지에서 처리하므로 폼에서는 제어할수 없음
                    //해당 체크박스는 숨김 처리
                    IDC_SAVE_CANCEL_CARD_SLIP.SetCommandParamValue("P_CARD_APPROVAL_ID", IGR_CREDIT_APPR.GetCellValue(r, vIDX_CARD_APPROVAL_ID));
                    IDC_SAVE_CANCEL_CARD_SLIP.SetCommandParamValue("P_SYS_DATE", vSYS_DATE);
                    IDC_SAVE_CANCEL_CARD_SLIP.SetCommandParamValue("P_CARD_NUM", IGR_CREDIT_APPR.GetCellValue(r, vIDX_CARD_NUM));
                    IDC_SAVE_CANCEL_CARD_SLIP.SetCommandParamValue("P_APPROVAL_NUM", IGR_CREDIT_APPR.GetCellValue(r, vIDX_APPROVAL_NUM));
                    IDC_SAVE_CANCEL_CARD_SLIP.SetCommandParamValue("P_SESSION_ID", mSession_ID);
                    IDC_SAVE_CANCEL_CARD_SLIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SAVE_CANCEL_CARD_SLIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SAVE_CANCEL_CARD_SLIP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SAVE_CANCEL_CARD_SLIP.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_SAVE_CANCEL_CARD_SLIP.ExcuteErrorMsg, "Slip Create-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Slip Create-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    //IGR_CREDIT_APPR.SetCellValue(r, vIDX_SELECT_YN, "N");

                    IGR_CREDIT_APPR.LastConfirmChanges();
                    IDA_CARD_APPROVAL.OraSelectData.AcceptChanges();
                    IDA_CARD_APPROVAL.Refillable = true;
                }
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            //전표 삭제 대상 메시지//
            IDC_GET_CANCEL_CARD_SLIP_P.SetCommandParamValue("P_SYS_DATE", vSYS_DATE);
            IDC_GET_CANCEL_CARD_SLIP_P.SetCommandParamValue("P_SESSION_ID", mSession_ID);
            IDC_GET_CANCEL_CARD_SLIP_P.ExecuteNonQuery();
            vMESSAGE = iConv.ISNull(IDC_GET_CANCEL_CARD_SLIP_P.GetCommandParamValue("O_MESSAGE"));
            if(MessageBoxAdv.Show(vMESSAGE, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            //전표 삭제//
            IDC_CANCEL_SLIP_TRANSFER.SetCommandParamValue("W_SYS_DATE", vSYS_DATE);
            IDC_CANCEL_SLIP_TRANSFER.SetCommandParamValue("W_SESSION_ID", mSession_ID);
            IDC_CANCEL_SLIP_TRANSFER.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CANCEL_SLIP_TRANSFER.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CANCEL_SLIP_TRANSFER.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CANCEL_SLIP_TRANSFER.ExcuteError)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(IDC_CANCEL_SLIP_TRANSFER.ExcuteErrorMsg, "Slip Cancel-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(vMESSAGE, "Slip Cancel-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Search_DB();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents(); 
        }


        private void BTN_SLIP_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_CREDIT_APPR.RowCount < 0)
            {
                return;
            }

            //전표인쇄//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            object vCARD_APPROVAL_ID = IGR_CREDIT_APPR.GetCellValue("CARD_APPROVAL_ID");
            IDC_DELETE_CARD_APPROVAL_SLIP.SetCommandParamValue("W_CARD_APPROVAL_ID", vCARD_APPROVAL_ID);
            IDC_DELETE_CARD_APPROVAL_SLIP.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_DELETE_CARD_APPROVAL_SLIP.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_DELETE_CARD_APPROVAL_SLIP.GetCommandParamValue("O_MESSAGE"));
            if (IDC_DELETE_CARD_APPROVAL_SLIP.ExcuteError)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(IDC_DELETE_CARD_APPROVAL_SLIP.ExcuteErrorMsg, "Slip Delete-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(vMESSAGE, "Slip Delete-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            IDA_CARD_SLIP_LIST.Fill();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_CREDIT_CARD_CORP_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_CODE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_PERSON_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
             
        }

        private void ILA_USE_PERSON_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
             
        }

        private void ILA_CREDIT_CARD_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CREDIT_CARD.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BUDGET_USE_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_BUDGET_USE_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_BUDGET_USE_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", IGR_CREDIT_APPR.GetCellValue("SLIP_DATE"));
            ILD_BUDGET_USE_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", IGR_CREDIT_APPR.GetCellValue("SLIP_DATE")); 
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_CREDIT_APPR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["CARD_APPROVAL_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
        }

        private void IDA_CREDIT_APPR_UpdateCompleted(object pSender)
        {
            IDA_CARD_SLIP_LIST.Fill();
        }

        private void IDA_CREDIT_APPR_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            int vIDX_EXCHANGE_RATE = IGR_CREDIT_APPR.GetColumnToIndex("EXCHANGE_RATE");
            object vUpdatable = 0;
            if (pBindingManager.DataRow == null)
            {
                vUpdatable = 0;
            }
            else
            {
                if(iConv.ISNull(pBindingManager.DataRow["CURRENCY_CODE"]) == iConv.ISNull(pBindingManager.DataRow["BASE_CURRENCY_CODE"]))
                {
                    vUpdatable = 0;
                }
                else
                {
                    vUpdatable = 1;
                }
            }
            IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXCHANGE_RATE].Insertable = vUpdatable;
            IGR_CREDIT_APPR.GridAdvExColElement[vIDX_EXCHANGE_RATE].Updatable = vUpdatable;

            Init_Select_YN(iConv.ISNull(W_SLIP_IF_FLAG.EditValue)); 
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
        }

        private void IDA_CMS_SLIP_DETAIL_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            } 
        }

        private void IDA_CMS_SLIP_DETAIL_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            } 
        }


        #endregion

    }
}