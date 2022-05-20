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
 

namespace FCMF0270
{
    public partial class FCMF0270 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        object mSession_ID;
        object mAccount_Book_ID;
        object mAccount_Set_ID;
        object mFiscal_Calendar_ID;
        object mDept_Level;
        object mAccount_Book_Name;
        string mCurrency_Code;
        object mBudget_Control_YN; 

        #endregion;


        #region ----- Constructor -----

        public FCMF0270(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void GetAccountBook()
        {
            IDC_ACCOUNT_BOOK.ExecuteNonQuery();
            mSession_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_SESSION_ID");
            mAccount_Book_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_ID");
            mAccount_Book_Name = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_BOOK_NAME");
            mAccount_Set_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_ACCOUNT_SET_ID");
            mFiscal_Calendar_ID = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_FISCAL_CALENDAR_ID");
            mDept_Level = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
            mCurrency_Code = iString.ISNull(IDC_ACCOUNT_BOOK.GetCommandParamValue("O_CURRENCY_CODE"));
            mBudget_Control_YN = IDC_ACCOUNT_BOOK.GetCommandParamValue("O_BUDGET_CONTROL_YN");
        }

        private void Search()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_PREPAID_EXP_LIST.TabIndex)
            {
                IDA_PREPAID_EXPENSE_LIST.Fill();
                IGR_PREPAID_EXPENSE_LIST.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_PREPAID_EXP_DETAIL.TabIndex)
            {
                IDA_PREPAID_EXPENSE.Fill();
                PREPAID_EXPENSE_CODE.Focus();

                Init_IF_VALUE(IF_SLIP_FLAG.CheckBoxString);
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        //조회된 자료에서 더블클릭하면 전표팝업 띄워준다.
        private void Show_Slip_Detail(int pSLIP_HEADER_ID)
        {
            if (pSLIP_HEADER_ID != 0)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, pSLIP_HEADER_ID);
                vFCMF0204.Show();

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                Application.UseWaitCursor = false;
            }
        }

        private void Init_Insert()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_PREPAID_EXP_LIST.TabIndex)
            {                
                TB_MAIN.SelectedIndex = TP_PREPAID_EXP_DETAIL.TabIndex - 1;
                TB_MAIN.SelectedTab.Focus();  
            }

            //기본값 설정
            //통화, 상태
            CURRENCY_CODE.EditValue = mCurrency_Code;

            //비용계산방법 
            IDC_GET_DEFAULT_VALUE_GROUP.SetCommandParamValue("W_GROUP_CODE", "EXP_SPREAD_METHOD");
            IDC_GET_DEFAULT_VALUE_GROUP.ExecuteNonQuery();
            EXP_SPREAD_METHOD_CODE.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE");
            EXP_SPREAD_METHOD_DESC.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE_NAME");

            //상태
            IDC_GET_DEFAULT_VALUE_GROUP.SetCommandParamValue("W_GROUP_CODE", "PREPAID_EXP_STATUS");
            IDC_GET_DEFAULT_VALUE_GROUP.ExecuteNonQuery();
            PREPAID_EXPENSE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE");
            PREPAID_EXPENSE_STATUS_DESC.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE_NAME");

            OCCUR_DATE.EditValue = DateTime.Today;
            CONTRACT_DATE_FR.EditValue = OCCUR_DATE.EditValue;
            CONTRACT_DATE_TO.EditValue = CONTRACT_DATE_FR.EditValue;
            REPLACE_DATE_FR.EditValue = OCCUR_DATE.EditValue;
            REPLACE_DATE_TO.EditValue = REPLACE_DATE_FR.EditValue;

            //환율, 외화금액 설정 
            Init_Currency();

            PREPAID_EXPENSE_DESC.Focus();
        }

        //외화금액 입력에 따른 BASE 금액 계산 
        private void Init_Amount()
        {
            if (iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue) == 0)
            {
                return;
            }
            else if (iString.ISDecimaltoZero(CURR_AMOUNT.EditValue) == 0)
            {
                return;
            }

            decimal mAMOUNT = iString.ISDecimaltoZero(CURR_AMOUNT.EditValue) * iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue);
            try
            {
                IDC_CONVERSION_BASE_AMOUNT.SetCommandParamValue("W_BASE_CURRENCY_CODE", mCurrency_Code);
                IDC_CONVERSION_BASE_AMOUNT.SetCommandParamValue("W_CONVERSION_AMOUNT", mAMOUNT);
                IDC_CONVERSION_BASE_AMOUNT.ExecuteNonQuery();
                AMOUNT.EditValue = Convert.ToDecimal(IDC_CONVERSION_BASE_AMOUNT.GetCommandParamValue("O_BASE_AMOUNT"));
            }
            catch
            {
                AMOUNT.EditValue = Convert.ToDecimal(Math.Round(mAMOUNT, 0));
            } 
        }

        // 부가세 관련 설정 제어 - 세액/공급가액(세액 * 10)
        private void Init_VAT_Amount()
        {
            IDC_VAT_AMT_P.SetCommandParamValue("W_VAT_TAX_TYPE", VAT_TAX_TYPE.EditValue);
            IDC_VAT_AMT_P.SetCommandParamValue("W_SUPPLY_AMT", RETURN_AMOUNT.EditValue);
            IDC_VAT_AMT_P.ExecuteNonQuery();
            VAT_AMOUNT.EditValue = IDC_VAT_AMT_P.GetCommandParamValue("O_VAT_AMT");
        }

        private void Init_Currency()
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) == string.Empty || iString.ISNull(CURRENCY_CODE.EditValue) == mCurrency_Code)
            {
                if (iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue) != Convert.ToDecimal(0))
                {
                    EXCHANGE_RATE.EditValue = null;
                }
                if (iString.ISDecimaltoZero(CURR_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    CURR_AMOUNT.EditValue = null;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                CURR_AMOUNT.ReadOnly = true;
                CURR_AMOUNT.Insertable = false;
                CURR_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                CURR_AMOUNT.TabStop = false;
            }
            else
            {
                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                CURR_AMOUNT.ReadOnly = false;
                CURR_AMOUNT.Insertable = true;
                CURR_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                CURR_AMOUNT.TabStop = true;
            }
            EXCHANGE_RATE.Invalidate();
            CURR_AMOUNT.Invalidate();
        }

        private void Init_IF_VALUE(object pIF_SLIP_FLAG)
        {
            if (iString.ISNull(pIF_SLIP_FLAG)  == "Y")
            {
                //전표전송됨 -> 수정 불가 
                OCCUR_DATE.Updatable = false;
                OCCUR_DATE.TabStop = false;

                EXP_SPREAD_METHOD_CODE.Updatable = false;
                EXP_SPREAD_METHOD_CODE.TabStop = false;

                //CONTRACT_DATE_FR.Updatable = false;
                //CONTRACT_DATE_FR.TabStop = false;

                //CONTRACT_DATE_TO.Updatable = false;
                //CONTRACT_DATE_TO.TabStop = false;

                REPLACE_DATE_FR.Updatable = false;
                REPLACE_DATE_FR.TabStop = false;

                REPLACE_DATE_TO.Updatable = false;
                REPLACE_DATE_TO.TabStop = false;

                CURRENCY_CODE.Updatable = false;
                CURRENCY_CODE.TabStop = false;
 
                EXCHANGE_RATE.Updatable = false;
                EXCHANGE_RATE.TabStop = false;

                CURR_AMOUNT.Updatable = false;
                CURR_AMOUNT.TabStop = false;

                AMOUNT.Updatable = false;
                AMOUNT.TabStop = false;
            }
            else
            {
                OCCUR_DATE.Updatable = true;
                OCCUR_DATE.TabStop = true;

                EXP_SPREAD_METHOD_CODE.Updatable = true;
                EXP_SPREAD_METHOD_CODE.TabStop = true;

                //CONTRACT_DATE_FR.Updatable = true;
                //CONTRACT_DATE_FR.TabStop = true;

                //CONTRACT_DATE_TO.Updatable = true;
                //CONTRACT_DATE_TO.TabStop = true;

                REPLACE_DATE_FR.Updatable = true;
                REPLACE_DATE_FR.TabStop = true;

                REPLACE_DATE_TO.Updatable = true;
                REPLACE_DATE_TO.TabStop = true;

                CURRENCY_CODE.Updatable = true;
                CURRENCY_CODE.TabStop = true;

                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                CURR_AMOUNT.ReadOnly = false;
                CURR_AMOUNT.Insertable = true;
                CURR_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                CURR_AMOUNT.TabStop = true;

                AMOUNT.Updatable = true;
                AMOUNT.TabStop = true;                 
            }
            OCCUR_DATE.Invalidate();
            EXP_SPREAD_METHOD_CODE.Invalidate();
            //CONTRACT_DATE_FR.Invalidate();
            //CONTRACT_DATE_TO.Invalidate();
            REPLACE_DATE_FR.Invalidate();
            REPLACE_DATE_TO.Invalidate();
            CURRENCY_CODE.Invalidate();
            EXCHANGE_RATE.Invalidate();
            CURR_AMOUNT.Invalidate();
            AMOUNT.Invalidate();
        }

        private bool Save_Slip_Date()
        {
            if (iString.ISNull(CANCEL_SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return false;
            }

            IDC_SAVE_PREPAID_EXP.SetCommandParamValue("W_PREPAID_EXPENSE_ID", PREPAID_EXPENSE_ID.EditValue);
            IDC_SAVE_PREPAID_EXP.SetCommandParamValue("P_CANCEL_SLIP_DATE", CANCEL_SLIP_DATE.EditValue);
            IDC_SAVE_PREPAID_EXP.SetCommandParamValue("P_CANCEL_EXCHANGE_RATE", CANCEL_EXCHANGE_RATE.EditValue);
            IDC_SAVE_PREPAID_EXP.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SAVE_PREPAID_EXP.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SAVE_PREPAID_EXP.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SAVE_PREPAID_EXP.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SAVE_PREPAID_EXP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CANCEL_SLIP_DATE.Focus();
                return false;
            }
            else if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CANCEL_SLIP_DATE.Focus();
                return false;
            }
            return true;
        }
         
        #endregion;


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
                            vParam[2] = CANCEL_SLIP_DATE.EditValue;     //전표일자 시작
                            vParam[3] = CANCEL_SLIP_DATE.EditValue;     //전표일자 종료
                            vParam[4] = DBNull.Value;                   //기표번호
                            vParam[5] = CANCEL_SLIP_NUM.EditValue;      //전표번호

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

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)        //검색
                {
                    Search();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)  //위에 새레코드 추가
                {
                    IDA_PREPAID_EXPENSE.AddOver();
                    Init_Insert();
                                         
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder) //아래에 새레코드 추가
                {
                    IDA_PREPAID_EXPENSE.AddUnder();
                    Init_Insert();
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)   //저장
                {
                    if (IDA_PE_CANCEL_SLIP_ACC.IsFocused)
                    {
                        IDA_PE_CANCEL_SLIP_ACC.Update();
                    }
                    else
                    {
                        IDA_PREPAID_EXPENSE.Update();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)   //취소
                {
                    if (IDA_PREPAID_EXPENSE.IsFocused)
                    {
                        IDA_PREPAID_EXPENSE.Cancel(); 
                    }
                    else if (IDA_PE_CANCEL_SLIP_ACC.IsFocused)
                    {
                        IDA_PE_CANCEL_SLIP_ACC.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)   //삭제
                {
                    if (IDA_PREPAID_EXPENSE.IsFocused)
                    {
                        IDA_PREPAID_EXPENSE.Delete();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)    //인쇄
                {
                    //XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                {
                    //XLPrinting("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0270_Load(object sender, EventArgs e)
        {            
        }

        private void FCMF0270_Shown(object sender, EventArgs e)
        {
            IDC_GET_DEFAULT_VALUE_GROUP.SetCommandParamValue("W_GROUP_CODE", "PREPAID_EXP_STATUS");
            IDC_GET_DEFAULT_VALUE_GROUP.ExecuteNonQuery();
            W_PREPAID_EXPENSE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE");
            W_PREPAID_EXPENSE_STATUS_DESC.EditValue = IDC_GET_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE_NAME");

            GetAccountBook();
            Init_Currency();

            IDA_PREPAID_EXPENSE_LIST.FillSchema();    //로드시에 FillSchema를 해주면 바로 신규 자료 INSERT가 가능하다.
            IDA_PREPAID_EXPENSE.FillSchema();
            IDA_PE_CANCEL_SLIP_ACC.FillSchema();
        }

        private void IGR_PREPAID_EXPENSE_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_PREPAID_EXPENSE_LIST.RowIndex > -1)
            {
                V_PREPAID_EXPENSE_ID.EditValue = IGR_PREPAID_EXPENSE_LIST.GetCellValue("PREPAID_EXPENSE_ID");
                TB_MAIN.SelectedIndex = TP_PREPAID_EXP_DETAIL.TabIndex - 1;
                TB_MAIN.SelectedTab.Focus();

                Search();
            }
        }

        private void IGR_PREPAID_EXP_HISTORY_CellDoubleClick(object pSender)
        {
            if (IGR_PREPAID_EXP_HISTORY.RowIndex > -1)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_PREPAID_EXP_HISTORY.GetCellValue("SLIP_HEADER_ID"));
                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        private void RETURN_AMOUNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_VAT_Amount();
        }

        private void BTN_AR_ACC_CODE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PREPAID_EXPENSE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PREPAID_EXPENSE_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PREPAID_EXPENSE_CODE.Focus();
                return;
            }
            if (iString.ISNull(CANCEL_SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return;
            }
            if (mCurrency_Code != iString.ISNull(CURRENCY_CODE.EditValue) && iString.ISDecimaltoZero(CANCEL_EXCHANGE_RATE.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_EXCHANGE_RATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return;
            }

            if (Save_Slip_Date() == false)
            {
                return;
            }

            //외화금액//
            object vCURRENCY_AMOUNT = 0;
            if (mCurrency_Code != iString.ISNull(CURRENCY_CODE.EditValue))
            {
                IDC_GET_CURR_AMOUNT.SetCommandParamValue("P_GL_AMOUNT", V_GAP_AMOUNT.EditValue);
                IDC_GET_CURR_AMOUNT.SetCommandParamValue("P_EXCHANGE_AMOUNT", CANCEL_EXCHANGE_RATE.EditValue);
                IDC_GET_CURR_AMOUNT.ExecuteNonQuery();
                vCURRENCY_AMOUNT = IDC_GET_CURR_AMOUNT.GetCommandParamValue("O_CURR_AMOUNT");
            }
            DialogResult vResult = DialogResult.None;
            FCMF0270_ACCOUNT vFCMF0270_ACCOUNT = new FCMF0270_ACCOUNT(MdiParent, isAppInterfaceAdv1.AppInterface, PREPAID_EXPENSE_ID.EditValue
                                                                    , "AR", CANCEL_SLIP_DATE.EditValue, USE_DEPT_ID.EditValue, USE_DEPT_NAME.EditValue
                                                                    , USE_DEPT_ID.EditValue, USE_DEPT_ID.EditValue, CURRENCY_CODE.EditValue
                                                                    , CANCEL_EXCHANGE_RATE.EditValue, vCURRENCY_AMOUNT, V_RETURN_TOTAL_AMOUNT.EditValue
                                                                    , VENDOR_CODE.EditValue, VENDOR_DESC.EditValue
                                                                    , VAT_TAX_TYPE.EditValue, VAT_TAX_TYPE_DESC.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0270_ACCOUNT, isAppInterfaceAdv1.AppInterface);
            vResult = vFCMF0270_ACCOUNT.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                IDA_PE_CANCEL_SLIP_ACC.Fill();
            }
            vFCMF0270_ACCOUNT.Dispose();
        }

        private void BTN_VAT_ACC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PREPAID_EXPENSE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PREPAID_EXPENSE_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PREPAID_EXPENSE_CODE.Focus();
                return;
            }
            if (iString.ISNull(CANCEL_SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return;
            }
            if (mCurrency_Code != iString.ISNull(CURRENCY_CODE.EditValue) && iString.ISDecimaltoZero(CANCEL_EXCHANGE_RATE.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_EXCHANGE_RATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return;
            }

            if (Save_Slip_Date() == false)
            {
                return;
            }

            //외화금액//
            object vCURRENCY_AMOUNT = 0;
            if (mCurrency_Code != iString.ISNull(CURRENCY_CODE.EditValue))
            {
                IDC_GET_CURR_AMOUNT.SetCommandParamValue("P_GL_AMOUNT", V_GAP_AMOUNT.EditValue);
                IDC_GET_CURR_AMOUNT.SetCommandParamValue("P_EXCHANGE_AMOUNT", CANCEL_EXCHANGE_RATE.EditValue);
                IDC_GET_CURR_AMOUNT.ExecuteNonQuery();
                vCURRENCY_AMOUNT = IDC_GET_CURR_AMOUNT.GetCommandParamValue("O_CURR_AMOUNT");
            }
            DialogResult vResult = DialogResult.None;
            FCMF0270_ACCOUNT vFCMF0270_ACCOUNT = new FCMF0270_ACCOUNT(MdiParent, isAppInterfaceAdv1.AppInterface, PREPAID_EXPENSE_ID.EditValue
                                                                    , "VAT", CANCEL_SLIP_DATE.EditValue, USE_DEPT_ID.EditValue, USE_DEPT_NAME.EditValue
                                                                    , USE_DEPT_ID.EditValue, USE_DEPT_ID.EditValue, CURRENCY_CODE.EditValue
                                                                    , CANCEL_EXCHANGE_RATE.EditValue, vCURRENCY_AMOUNT, VAT_AMOUNT.EditValue
                                                                    , VENDOR_CODE.EditValue, VENDOR_DESC.EditValue
                                                                    , VAT_TAX_TYPE.EditValue, VAT_TAX_TYPE_DESC.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0270_ACCOUNT, isAppInterfaceAdv1.AppInterface);
            vResult = vFCMF0270_ACCOUNT.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                IDA_PE_CANCEL_SLIP_ACC.Fill();
            }
            vFCMF0270_ACCOUNT.Dispose();
        }

        private void BTN_PL_ACC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PREPAID_EXPENSE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PREPAID_EXPENSE_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PREPAID_EXPENSE_CODE.Focus();
                return;
            }
            if (iString.ISNull(CANCEL_SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return;
            }
            if (mCurrency_Code != iString.ISNull(CURRENCY_CODE.EditValue) && iString.ISDecimaltoZero(CANCEL_EXCHANGE_RATE.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_EXCHANGE_RATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                return;
            }

            if (Save_Slip_Date() == false)
            {
                return;
            }

            //외화금액//
            object vCURRENCY_AMOUNT = 0;
            if (mCurrency_Code != iString.ISNull(CURRENCY_CODE.EditValue))
            {
                IDC_GET_CURR_AMOUNT.SetCommandParamValue("P_GL_AMOUNT", V_GAP_AMOUNT.EditValue);
                IDC_GET_CURR_AMOUNT.SetCommandParamValue("P_EXCHANGE_AMOUNT", CANCEL_EXCHANGE_RATE.EditValue);
                IDC_GET_CURR_AMOUNT.ExecuteNonQuery();
                vCURRENCY_AMOUNT = IDC_GET_CURR_AMOUNT.GetCommandParamValue("O_CURR_AMOUNT");
            }
            DialogResult vResult = DialogResult.None;
            FCMF0270_ACCOUNT vFCMF0270_ACCOUNT = new FCMF0270_ACCOUNT(MdiParent, isAppInterfaceAdv1.AppInterface, PREPAID_EXPENSE_ID.EditValue
                                                                    , "PL", CANCEL_SLIP_DATE.EditValue, USE_DEPT_ID.EditValue, USE_DEPT_NAME.EditValue
                                                                    , USE_DEPT_ID.EditValue, USE_DEPT_ID.EditValue, CURRENCY_CODE.EditValue
                                                                    , Math.Abs(iString.ISDecimaltoZero(CANCEL_EXCHANGE_RATE.EditValue))
                                                                    , Math.Abs(iString.ISDecimaltoZero(vCURRENCY_AMOUNT))
                                                                    , Math.Abs(iString.ISDecimaltoZero(V_GAP_AMOUNT.EditValue))
                                                                    , VENDOR_CODE.EditValue, VENDOR_DESC.EditValue
                                                                    , VAT_TAX_TYPE.EditValue, VAT_TAX_TYPE_DESC.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0270_ACCOUNT, isAppInterfaceAdv1.AppInterface);
            vResult = vFCMF0270_ACCOUNT.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                IDA_PE_CANCEL_SLIP_ACC.Fill();
            }
            vFCMF0270_ACCOUNT.Dispose();
        }

        private void BTN_SLIP_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PREPAID_EXPENSE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PREPAID_EXPENSE_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_PE_CANCEL_SLIP.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_SET_PE_CANCEL_SLIP.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_SET_PE_CANCEL_SLIP.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_PE_CANCEL_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_PE_CANCEL_SLIP.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            IDA_PE_CANCEL_SLIP_ACC.Fill();
        }

        private void BTN_SLIP_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(PREPAID_EXPENSE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PREPAID_EXPENSE_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_PE_CANCEL_SLIP.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_PE_CANCEL_SLIP.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_PE_CANCEL_SLIP.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_PE_CANCEL_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CANCEL_PE_CANCEL_SLIP.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }
            
            IDA_PE_CANCEL_SLIP_ACC.Fill();
        }

        private void BTN_EXEC_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(CANCEL_SLIP_NUM.EditValue) != string.Empty)
            {
                AssmblyRun_Manual(CANCEL_ASSEMBLY_ID.EditValue);
            }
        }

        #endregion


        #region ----- Lookup Event -----

        private void ILA_COSTCENTER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_PREPAID_EXP_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_STATUS", "Y");
        }

        private void ILA_ENTERED_METHOD_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("ENTERED_METHOD", "Y");
        }

        private void ILA_PREPAID_EXP_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_TYPE", "Y");
        }

        private void ILA_PREPAID_EXP_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_TYPE", "Y");
        }

        private void ILA_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_ENTRY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_CLASS.SetLookupParamValue("W_ACCOUNT_CLASS_TYPE", "PREPAID");
            ILD_ACCOUNT_CONTROL_CLASS.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_REPLACE_ACCOUNT_CONTROL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_EXP_SPREAD_METHOD_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("EXP_SPREAD_METHOD", "Y");
        }

        private void ILA_PREPAID_EXP_STATUS_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("PREPAID_EXP_STATUS", "Y");
        }

        private void ILA_CURRENCY_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_SelectedRowData(object pSender)
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) != string.Empty)
            {
                if (iString.ISNull(CURRENCY_CODE.EditValue) != mCurrency_Code)
                {
                    IDC_EXCHANGE_RATE.ExecuteNonQuery();
                    EXCHANGE_RATE.EditValue = IDC_EXCHANGE_RATE.GetCommandParamValue("X_EXCHANGE_RATE");

                    Init_Amount();
                }
            }
            Init_Currency();
        }

        private void ILA_VAT_TAX_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VAT_TAX_TYPE.SetLookupParamValue("W_AP_AR_TYPE", "AR");
            ILD_VAT_TAX_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_VAT_TAX_TYPE_SelectedRowData(object pSender)
        {
            Init_VAT_Amount();
        }

        #endregion

        #region ----- Adapter Lookup Event -----
                   
        private void IDA_PREPAID_EXPENSE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //대체계정은 필수 항목입니다.
            if (iString.ISNull(e.Row["REPLACE_ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REPLACE_ACCOUNT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                 
                REPLACE_ACCOUNT_CODE.Focus();
                e.Cancel = true; 
                return;
            }

            //원가코드는 필수 항목입니다.
            if (iString.ISNull(e.Row["COST_CENTER_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10524"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                COST_CENTER_CODE.Focus();
                e.Cancel = true; 
                return;
            }

            //거래처정보는 필수항목입니다.
            if (iString.ISNull(e.Row["VENDOR_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                VENDOR_CODE.Focus();
                e.Cancel = true; 
                return;
            }

            //부서코드는 필수 항목입니다
            //if (iString.ISNull(e.Row["USE_DEPT_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10019"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    USE_DEPT_CODE.Focus(); 
            //    return;
            //}

            //자료상태는 필수 항목입니다.
            if (iString.ISNull(e.Row["PREPAID_EXPENSE_STATUS"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10534"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PREPAID_EXPENSE_STATUS.Focus();
                e.Cancel = true; 
                return;
            }


            //적요는 필수항목입니다.
            if (iString.ISNull(e.Row["PREPAID_EXPENSE_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PREPAID_EXPENSE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PREPAID_EXPENSE_DESC.Focus();
                e.Cancel = true; 
                return;
            }

            //계산방법은 필수 항목입니다.
            if (iString.ISNull(e.Row["EXP_SPREAD_METHOD_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10536"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                EXP_SPREAD_METHOD_CODE.Focus();
                e.Cancel = true; 
                return;
            }
            
            //통화 필수 항목입니다.
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CURRENCY_CODE.Focus();
                e.Cancel = true;
                return;
            }

            //금액은 필수 항목입니다.
            if (iString.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10537"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AMOUNT.Focus();
                e.Cancel = true; 
                return;
            }

            //계약기간은 필수 항목입니다.(시작일)
            if (iString.ISNull(e.Row["CONTRACT_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CONTRACT_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CONTRACT_DATE_FR.Focus();
                e.Cancel = true; 
                return;
            }


            //계약기간은 필수 항목입니다.(종료일)
            if (iString.ISNull(e.Row["CONTRACT_DATE_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CONTRACT_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CONTRACT_DATE_TO.Focus();
                e.Cancel = true; 
                return;
            }

            if (iString.ISNull(e.Row["EXCEPTION_FLAG"]) == "Y" && iString.ISNull(e.Row["EXCEPTION_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(EXCEPTION_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                EXCEPTION_DESC.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EXCEPTION_FLAG"]) == "N" && iString.ISNull(e.Row["EXCEPTION_DESC"]) != string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10569"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                EXCEPTION_DESC.Focus();
                e.Cancel = true;
                return;
            }
             
            //해지관련정보 입력시 메세지 처리
            //해지일자는 계약기간 이내이어야 함
            //해지관련 항목에 값을 입력하지 않은 경우(단, 환급금액의 경우 0을 입력한 경우는 유효한 값을 입력하지 안은 것으로 처리한다.)는 메세지를 띄우지 않는다.
            if (iString.ISNull(e.Row["CANCEL_DATE"]) == string.Empty && iString.ISNull(e.Row["CANCEL_REASON"]) != string.Empty)
            {
                //환급일자없이 환급금액 또는 환급사유 등록시 오류
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10572"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_DATE.Focus();
                e.Cancel = true;
                return;
            } 
            //해지관련 3개 항목 모두에 값이 입력되었는데 그 중 환급금액에 0(zero)을 입력한 경우 메세지 처리한다.
            else if (iString.ISNull(e.Row["CANCEL_DATE"]) != string.Empty && iString.ISNull(e.Row["CANCEL_REASON"]) == string.Empty)
            {
                //환급일자없이 환급금액 또는 환급사유 등록시 오류
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10571"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_REASON.Focus();
                e.Cancel = true;
                return;
            }  
        }

        private void IDA_PE_CANCEL_SLIP_ACC_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //해지 전표일자
            if (iString.ISNull(e.Row["CANCEL_SLIP_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CANCEL_SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CANCEL_SLIP_DATE.Focus();
                e.Cancel = true;
                return;
            } 
        }


        private void IDA_PREPAID_EXPENSE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if(pBindingManager.DataRow == null)
            {
                Init_IF_VALUE("N");

                CANCEL_SLIP_DATE.Insertable = false;
                CANCEL_SLIP_DATE.Updatable = false;
                CANCEL_SLIP_DATE.Refresh();

                CANCEL_EXCHANGE_RATE.Insertable = false;
                CANCEL_EXCHANGE_RATE.Updatable = false;
                CANCEL_EXCHANGE_RATE.Refresh();

                BTN_AR_ACC_CODE.Enabled = false;
                BTN_VAT_ACC.Enabled = false;
                BTN_PL_ACC.Enabled = false;

                BTN_EXEC_SLIP.Enabled = false;
                BTN_SLIP_OK.Enabled = false;
                BTN_SLIP_CANCEL.Enabled = false;
                return;
            }

            Init_IF_VALUE(pBindingManager.DataRow["IF_SLIP_FLAG"]);
             
            if (iString.ISNull(pBindingManager.DataRow["CANCEL_DATE"]) != string.Empty)
            {
                CANCEL_SLIP_DATE.Insertable = true;
                CANCEL_SLIP_DATE.Updatable = true;
                CANCEL_SLIP_DATE.Refresh();

                CANCEL_EXCHANGE_RATE.Insertable = false;
                CANCEL_EXCHANGE_RATE.Updatable = false;
                CANCEL_EXCHANGE_RATE.Refresh();
                if (iString.ISNull(mCurrency_Code) != iString.ISNull(CURRENCY_CODE.EditValue))
                {                    
                    CANCEL_EXCHANGE_RATE.Insertable = true;
                    CANCEL_EXCHANGE_RATE.Updatable = true;
                    CANCEL_EXCHANGE_RATE.Refresh(); 
                } 

                BTN_AR_ACC_CODE.Enabled = true;
            }
            else
            {
                CANCEL_SLIP_DATE.Insertable = false;
                CANCEL_SLIP_DATE.Updatable = false;
                CANCEL_SLIP_DATE.Refresh();

                CANCEL_EXCHANGE_RATE.Insertable = false;
                CANCEL_EXCHANGE_RATE.Updatable = false;
                CANCEL_EXCHANGE_RATE.Refresh();

                BTN_AR_ACC_CODE.Enabled = false;
            }

            if (iString.ISNull(pBindingManager.DataRow["INPUT_VAT_ACC_FLAG"]) == "Y")
            {
                BTN_VAT_ACC.Enabled = true;
            }
            else
            {
                BTN_VAT_ACC.Enabled = false; 
            }

            if (iString.ISNull(pBindingManager.DataRow["INPUT_PL_ACC_FLAG"]) == "Y")
            {
                BTN_PL_ACC.Enabled = true;
            }
            else
            {
                BTN_PL_ACC.Enabled = false;
            }

            if (iString.ISNull(pBindingManager.DataRow["CANCEL_SLIP_FLAG"]) == "Y")
            {
                BTN_EXEC_SLIP.Enabled = true;
                BTN_SLIP_OK.Enabled = true;
                BTN_SLIP_CANCEL.Enabled = true;
            }
            else if (iString.ISNull(pBindingManager.DataRow["CANCEL_DATE"]) != string.Empty)
            {
                BTN_EXEC_SLIP.Enabled = true;
                BTN_SLIP_OK.Enabled = true;
                BTN_SLIP_CANCEL.Enabled = true;
            }
            else
            {
                BTN_EXEC_SLIP.Enabled = false;
                BTN_SLIP_OK.Enabled = false;
                BTN_SLIP_CANCEL.Enabled = false;
            }

            //해지 전표 정보 조회//
            IDA_PE_CANCEL_SLIP_ACC.OraSelectData.AcceptChanges();
            IDA_PE_CANCEL_SLIP_ACC.Refillable = true;

            IDA_PE_CANCEL_SLIP_ACC.Fill();
        }

        private void IDA_PREPAID_EXPENSE_UpdateCompleted(object pSender)
        {
            V_PREPAID_EXPENSE_ID.EditValue = PREPAID_EXPENSE_ID.EditValue;

            Search();
        }


        #endregion

    }
}