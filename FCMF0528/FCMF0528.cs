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

namespace FCMF0528
{
    public partial class FCMF0528 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0528()
        {
            InitializeComponent();
        }

        public FCMF0528(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {            
            Set_Tab_Focus();
        }

        private void Set_Tab_Focus()
        {
            if (iConv.ISNull(W_ESTIMATE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ESTIMATE_DATE.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_EXCHANGE_RATE.TabIndex)
            {
                IGR_STATEMENT_EXCHANGE.LastConfirmChanges();
                IDA_STATEMENT_EXCHANGE.OraSelectData.AcceptChanges();
                IDA_STATEMENT_EXCHANGE.Refillable = true; 

                IDA_STATEMENT_EXCHANGE.Fill();
                IGR_STATEMENT_EXCHANGE.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_CURR_ESTIMATE.TabIndex)
            {
                IDA_BALANCE_ACCOUNT.SetSelectParamValue("W_ACCOUNT_ALL", "N");
                IDA_BALANCE_ACCOUNT.Fill();                
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_CURR_ESTIMATE_SLIP.TabIndex)
            {
                IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
                IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
                IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;

                IDA_BALANCE_ACCOUNT_SLIP.Fill();  
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_REV_SLIP.TabIndex)
            {
                IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
                IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
                IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;

                IDA_BALANCE_ACCOUNT_R_SLIP.Fill();
            }
        }

        private void INIT_MANAGEMENT_COLUMN_2(object pACCOUNT_CONTROL_ID)
        {
            int mStart_Column = 4;
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            string mENABLED_FLAG;       // 사용(표시)여부.
            string mCOLUMN_DESC;        // 헤더 프롬프트.

            IDA_ITEM_PROMPT_2.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_ITEM_PROMPT_2.Fill();
            if (IDA_ITEM_PROMPT_2.OraSelectData.Rows.Count == 0)
            {
                for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mENABLED_COLUMN = mMax_Column + mIDX_Column;
                    IGR_BALANCE_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0; 
                } 
                IGR_BALANCE_CURR_ESTIMATE.ResetDraw = true;
                return;
            }             

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT_2.CurrentRow[mENABLED_COLUMN], "N");

                if (mENABLED_FLAG == "N")
                {
                    IGR_BALANCE_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BALANCE_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                }
            }

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = iConv.ISNull(IDA_ITEM_PROMPT_2.CurrentRow[mIDX_Column]);
                if (mCOLUMN_DESC != string.Empty)
                {
                    IGR_BALANCE_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = mCOLUMN_DESC;
                    IGR_BALANCE_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = mCOLUMN_DESC;
                }
            } 
            IGR_BALANCE_CURR_ESTIMATE.ResetDraw = true;
        }

        private void INIT_MANAGEMENT_COLUMN_3(object pACCOUNT_CONTROL_ID)
        {
            IDA_ITEM_PROMPT_3.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_ITEM_PROMPT_3.Fill();

            int mStart_Column = 10;
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            string mENABLED_FLAG;       // 사용(표시)여부.
            string mCOLUMN_DESC;        // 헤더 프롬프트.

            if (IDA_ITEM_PROMPT_3.OraSelectData.Rows.Count == 0)
            {
                for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mENABLED_COLUMN = mMax_Column + mIDX_Column;  
                    IGR_CURR_ESTIMATE_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0; 
                } 
                IGR_CURR_ESTIMATE_SLIP.ResetDraw = true; 
                return;
            }
  
            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT_3.CurrentRow[mENABLED_COLUMN], "N");

                if (mENABLED_FLAG == "N")
                {
                    IGR_CURR_ESTIMATE_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_CURR_ESTIMATE_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                }
            }

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = iConv.ISNull(IDA_ITEM_PROMPT_3.CurrentRow[mIDX_Column]);
                if (mCOLUMN_DESC != string.Empty)
                {
                    IGR_CURR_ESTIMATE_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = mCOLUMN_DESC;
                    IGR_CURR_ESTIMATE_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = mCOLUMN_DESC;
                }
            }
            IGR_CURR_ESTIMATE_SLIP.ResetDraw = true;
        }

        private void INIT_MANAGEMENT_COLUMN_4(object pACCOUNT_CONTROL_ID)
        {
            IDA_ITEM_PROMPT_3.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_ITEM_PROMPT_3.Fill();

            int mStart_Column = 10;
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            string mENABLED_FLAG;       // 사용(표시)여부.
            string mCOLUMN_DESC;        // 헤더 프롬프트.
            if (IDA_ITEM_PROMPT_3.OraSelectData.Rows.Count == 0)
            {
                for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mENABLED_COLUMN = mMax_Column + mIDX_Column;
                    IGR_CURR_ESTIMATE_R_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0; 
                } 
                IGR_CURR_ESTIMATE_R_SLIP.ResetDraw = true;
                return;
            }
                        

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT_3.CurrentRow[mENABLED_COLUMN], "N");

                if (mENABLED_FLAG == "N")
                {
                    IGR_CURR_ESTIMATE_R_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_CURR_ESTIMATE_R_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                }
            }

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = iConv.ISNull(IDA_ITEM_PROMPT_3.CurrentRow[mIDX_Column]);
                if (mCOLUMN_DESC != string.Empty)
                {
                    IGR_CURR_ESTIMATE_R_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = mCOLUMN_DESC;
                    IGR_CURR_ESTIMATE_R_SLIP.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = mCOLUMN_DESC;
                }
            }
            IGR_CURR_ESTIMATE_R_SLIP.ResetDraw = true;
        }        

        private void SUM_SLIP_AMOUNT()
        {
            decimal vGL_AMOUNT = 0; 

            int vIDX_GL_AMOUNT = IGR_CURR_ESTIMATE_SLIP.GetColumnToIndex("GL_AMOUNT");
            for (int vROW = 0; vROW < IGR_CURR_ESTIMATE_SLIP.RowCount; vROW++)
            {
                vGL_AMOUNT = vGL_AMOUNT + iConv.ISDecimaltoZero(IGR_CURR_ESTIMATE_SLIP.GetCellValue(vROW, vIDX_GL_AMOUNT), 0); 
            }
            V_SUM_AMOUNT.EditValue = vGL_AMOUNT;
        }

        private void SUM_R_SLIP_AMOUNT()
        {
            decimal vGL_AMOUNT = 0;

            int vIDX_GL_AMOUNT = IGR_CURR_ESTIMATE_R_SLIP.GetColumnToIndex("GL_AMOUNT");
            for (int vROW = 0; vROW < IGR_CURR_ESTIMATE_R_SLIP.RowCount; vROW++)
            {
                vGL_AMOUNT = vGL_AMOUNT + iConv.ISDecimaltoZero(IGR_CURR_ESTIMATE_R_SLIP.GetCellValue(vROW, vIDX_GL_AMOUNT), 0);
            }
            V_SUM_R_AMOUNT.EditValue = vGL_AMOUNT;
        }

        private void Init_Check_3()
        {
            string vCheck_Value = V_CHECK_FLAG_3.CheckBoxString;

            int vIDX_CHECK = IGR_BALANCE_ACCOUNT_SLIP.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < IGR_BALANCE_ACCOUNT_SLIP.RowCount; i++)
            {
                IGR_BALANCE_ACCOUNT_SLIP.SetCellValue(i, vIDX_CHECK, vCheck_Value);
            }
            
            IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
            IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
            IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;
        }

        private void Init_Check_4()
        {
            string vCheck_Value = V_CHECK_FLAG_4.CheckBoxString;

            int vIDX_CHECK = IGR_BALANCE_ACCOUNT_R_SLIP.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < IGR_BALANCE_ACCOUNT_R_SLIP.RowCount; i++)
            {
                IGR_BALANCE_ACCOUNT_R_SLIP.SetCellValue(i, vIDX_CHECK, vCheck_Value);
            }

            IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
            IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
            IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;
        }

        private void Init_BTN_3()
        {
            if (iConv.ISNull(W_SLIP_FLAG.EditValue) == "N")
            {
                BTN_SET_SLIP.Enabled = true;
                BTN_CANCEL_SLIP.Enabled = false;
                BTN_SHOW_SLIP.Enabled = false;
                W_SLIP_NUM_3.Visible = false;
                W_SLIP_NUM_3.EditValue = DBNull.Value;
            }
            else
            {
                BTN_SET_SLIP.Enabled = false;
                BTN_CANCEL_SLIP.Enabled = true;
                BTN_SHOW_SLIP.Enabled = true;
                W_SLIP_NUM_3.Visible = true;
            }
        }

        private void Init_BTN_4()
        {
            if (iConv.ISNull(W_R_SLIP_FLAG.EditValue) == "N")
            {
                BTN_SET_R_SLIP.Enabled = true;
                BTN_CANCEL_R_SLIP.Enabled = false;
                BTN_SHOW_R_SLIP.Enabled = false;
                V_R_GL_DATE.Insertable = true;
                V_R_GL_DATE.Updatable = true;
                V_R_GL_DATE.ReadOnly = false;

                W_SLIP_NUM_4.Visible = false;
                W_SLIP_NUM_4.EditValue = DBNull.Value;
            }
            else
            {
                V_R_GL_DATE.Insertable = false;
                V_R_GL_DATE.Updatable = false;
                V_R_GL_DATE.ReadOnly = true;

                BTN_SET_R_SLIP.Enabled = false;
                BTN_CANCEL_R_SLIP.Enabled = true;
                BTN_SHOW_R_SLIP.Enabled = true;
                W_SLIP_NUM_4.Visible = true;
            }
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

        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID, object pSLIP_DATE, object pSLIP_NUM)
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
                            vParam[2] = pSLIP_DATE;
                            vParam[3] = pSLIP_DATE;
                            vParam[4] = DBNull.Value;
                            vParam[5] = pSLIP_NUM;

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

        #region ----- XL Print 1 (계정잔액명세서) Method ----

        //private void XLPrinting_1(string pOutChoice)
        //{// pOutChoice : 출력구분.
        //    string vMessageText = string.Empty;
        //    string vSaveFileName = string.Empty;

        //    object vBALANCE_DATE = iDate.ISGetDate(BALANCE_DATE.EditValue).ToShortDateString();
        //    object vACCOUNT_CODE = W_ACCOUNT_CODE.EditValue;
        //    object vACCOUNT_DESC = W_ACCOUNT_DESC.EditValue;
        //    object vTerritory = string.Empty;
        //    object vGROUPING_OPTION = null;
            
        //    if (iConv.ISNull(vBALANCE_DATE) == String.Empty)
        //    {//기준일자
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    if (iConv.ISNull(vACCOUNT_CODE) == String.Empty)
        //    {//계정과목코드
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    int vCountRow = igrBALANCE_STATEMENT.RowCount;
        //    if (vCountRow < 1)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //    vSaveFileName = string.Format("Balance_{0}_{1}", vBALANCE_DATE, vACCOUNT_DESC);

        //    saveFileDialog1.Title = "Excel Save";
        //    saveFileDialog1.FileName = vSaveFileName;
        //    saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
        //    saveFileDialog1.DefaultExt = "xls";
        //    if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        //    {
        //        return;
        //    }
        //    else
        //    {
        //        vSaveFileName = saveFileDialog1.FileName;
        //        System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
        //        try
        //        {
        //            if (vFileName.Exists)
        //            {
        //                vFileName.Delete();
        //            }
        //        }
        //        catch (Exception EX)
        //        {
        //            MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    System.Windows.Forms.Application.DoEvents();

        //    int vPageNumber = 0;

        //    vMessageText = string.Format(" Printing Starting...");
        //    isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //    System.Windows.Forms.Application.DoEvents();

        //    vTerritory = GetTerritory();
        //    XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //    try
        //    {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                
        //        // open해야 할 파일명 지정.
        //        //-------------------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = "FCMF0528_001.xls";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            //조회시 그룹핑 옵션.
        //            if (iConv.ISNull(GROUPING_DUE_DATE.CheckBoxValue) == "Y")
        //            {
        //                switch (iConv.ISNull(vTerritory))
        //                {
        //                    case "TL1_KR":
        //                        vGROUPING_OPTION = string.Format("내역서({0})", GROUPING_DUE_DATE.PromptTextElement[0].TL1_KR);
        //                        break;
        //                    case "TL2_CN":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", GROUPING_DUE_DATE.PromptTextElement[0].TL2_CN);
        //                        break;
        //                    case "TL3_VN":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", GROUPING_DUE_DATE.PromptTextElement[0].TL3_VN);
        //                        break;
        //                    case "TL4_JP":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", GROUPING_DUE_DATE.PromptTextElement[0].TL4_JP);
        //                        break;
        //                    case "TL5_XAA":
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", GROUPING_DUE_DATE.PromptTextElement[0].TL5_XAA);
        //                        break;
        //                    default:                                
        //                        vGROUPING_OPTION = string.Format("Detailed Statement({0})", GROUPING_DUE_DATE.PromptTextElement[0].Default);
        //                        break;
        //                }
        //            }
        //            else
        //            {
        //                switch (iConv.ISNull(vTerritory))
        //                {
        //                    case "TL1_KR":
        //                        vGROUPING_OPTION = "내역서";
        //                        break;
        //                    case "TL2_CN":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    case "TL3_VN":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    case "TL4_JP":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    case "TL5_XAA":
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                    default:                                
        //                        vGROUPING_OPTION = "Detailed Statement";
        //                        break;
        //                }
        //            }
        //            //날짜형식 변경.
        //            IDC_DATE_YYYYMMDD.SetCommandParamValue("P_DATE", vBALANCE_DATE);
        //            IDC_DATE_YYYYMMDD.ExecuteNonQuery();
        //            vBALANCE_DATE = IDC_DATE_YYYYMMDD.GetCommandParamValue("O_DATE");
        //            xlPrinting.HeaderWrite(vACCOUNT_CODE, vACCOUNT_DESC, vBALANCE_DATE, vGROUPING_OPTION, iConv.ISNull(vTerritory), igrBALANCE_STATEMENT);

        //            // 실제 인쇄
        //            //vPageNumber = xlPrinting.LineWrite(vBALANCE_DATE, iConv.ISNull(vTerritory), pGRID);
        //            vPageNumber = xlPrinting.LineWrite(igrBALANCE_STATEMENT);

        //            //출력구분에 따른 선택(인쇄 or file 저장)
        //            if (pOutChoice == "PRINT")
        //            {
        //                xlPrinting.Printing(1, vPageNumber);
        //            }
        //            else if (pOutChoice == "FILE")
        //            {

        //                xlPrinting.SAVE(vSaveFileName);
        //            }

        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.Dispose();
        //            //-------------------------------------------------------------------------------------

        //            vMessageText = "Printing End";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        else
        //        {
        //            vMessageText = "Excel File Open Error";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        //-------------------------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        xlPrinting.Dispose();

        //        vMessageText = ex.Message;
        //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Application.DoEvents();
        //}

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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BALANCE_CURR_ESTIMATE.IsFocused)
                    {
                        IDA_BALANCE_CURR_ESTIMATE.Cancel();
                    }
                    else if (IDA_CURR_ESTIMATE_SLIP.IsFocused)
                    {
                        IDA_CURR_ESTIMATE_SLIP.Cancel();
                    }
                    else if (IDA_CURR_ESTIMATE_R_SLIP.IsFocused)
                    {
                        IDA_CURR_ESTIMATE_R_SLIP.Cancel();
                    }
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

        private void FCMF0528_Load(object sender, EventArgs e)
        {
            W_ESTIMATE_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);
            V_EXCHNAGE_DATE.EditValue = W_ESTIMATE_DATE.EditValue;

            BTN_EXE_ESTIMATE.BringToFront();
            BTN_CANECL_EXE_ESTIMATE.BringToFront();
            BTN_CLOSED_EXE_ESTIMATE.BringToFront();
            BTN_CLOSED_CANCEL_ESTIMATE.BringToFront();

            IDA_STATEMENT_EXCHANGE.FillSchema();
            IDA_BALANCE_ACCOUNT.FillSchema();
            IDA_BALANCE_ACCOUNT_SLIP.FillSchema();
            IDA_BALANCE_ACCOUNT_R_SLIP.FillSchema();

            IDA_CURR_ESTIMATE_SLIP.FillSchema();
            IDA_CURR_ESTIMATE_R_SLIP.FillSchema();
        }

        private void FCMF0528_Shown(object sender, EventArgs e)
        {            
            V_SUM_AMOUNT.BringToFront();
            V_SUM_R_AMOUNT.BringToFront();

            W_RB_SLIP_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SLIP_FLAG.EditValue = W_RB_SLIP_N.RadioButtonString;
            Init_BTN_3();

            W_RB_R_SLIP_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_R_SLIP_FLAG.EditValue = W_RB_R_SLIP_N.RadioButtonString;
            Init_BTN_4();

        }

        private void W_ESTIMATE_DATE_EditValueChanged(object pSender)
        {
            V_R_GL_DATE.EditValue = iDate.ISDate_Add(W_ESTIMATE_DATE.EditValue, 1);
            V_EXCHNAGE_DATE.EditValue = W_ESTIMATE_DATE.EditValue;
        }
         
        private void BTN_EXE_ESTIMATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_ESTIMATE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ESTIMATE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0528_ESTIMATE_SET vFCMF0528_ESTIMATE_SET = new FCMF0528_ESTIMATE_SET(isAppInterfaceAdv1.AppInterface, W_ESTIMATE_DATE.EditValue, "ESTIMATE", "Y");
            vRESULT = vFCMF0528_ESTIMATE_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0528_ESTIMATE_SET.Dispose();
        }

        private void BTN_CANECL_EXE_ESTIMATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_ESTIMATE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ESTIMATE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0528_ESTIMATE_SET vFCMF0528_ESTIMATE_SET = new FCMF0528_ESTIMATE_SET(isAppInterfaceAdv1.AppInterface, W_ESTIMATE_DATE.EditValue, "CANCEL_ESTIMATE", "N");
            vRESULT = vFCMF0528_ESTIMATE_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0528_ESTIMATE_SET.Dispose();
        }

        private void BTN_CLOSED_EXE_ESTIMATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_ESTIMATE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ESTIMATE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0528_ESTIMATE_SET vFCMF0528_ESTIMATE_SET = new FCMF0528_ESTIMATE_SET(isAppInterfaceAdv1.AppInterface, W_ESTIMATE_DATE.EditValue, "CLOSED_ESTIMATE", "N");
            vRESULT = vFCMF0528_ESTIMATE_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0528_ESTIMATE_SET.Dispose();
        }

        private void BTN_CLOSED_CANCEL_ESTIMATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_ESTIMATE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ESTIMATE_DATE.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0528_ESTIMATE_SET vFCMF0528_ESTIMATE_SET = new FCMF0528_ESTIMATE_SET(isAppInterfaceAdv1.AppInterface, W_ESTIMATE_DATE.EditValue, "CANCEL_CLOSED_ESTIMATE", "N");
            vRESULT = vFCMF0528_ESTIMATE_SET.ShowDialog();
            if (vRESULT == DialogResult.OK)
            {
                SearchDB();
            }
            vFCMF0528_ESTIMATE_SET.Dispose();
        }

        private void V_RB_SLIP_N_Click(object sender, EventArgs e)
        {
            if (W_RB_SLIP_N.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_FLAG.EditValue = W_RB_SLIP_N.RadioButtonString;
                Init_BTN_3();
            }
        }

        private void V_RB_SLIP_Y_Click(object sender, EventArgs e)
        {
            if (W_RB_SLIP_Y.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_FLAG.EditValue = W_RB_SLIP_Y.RadioButtonString;
                Init_BTN_3();
            }
        }

        private void V_RB_R_SLIP_N_Click(object sender, EventArgs e)
        {
            if (W_RB_R_SLIP_N.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_R_SLIP_FLAG.EditValue = W_RB_R_SLIP_N.RadioButtonString;
                Init_BTN_4();
            }
        }

        private void V_RB_R_SLIP_Y_Click(object sender, EventArgs e)
        {
            if (W_RB_R_SLIP_Y.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_R_SLIP_FLAG.EditValue = W_RB_R_SLIP_Y.RadioButtonString;
                Init_BTN_4();
            }
        }

        private void V_CHECK_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Init_Check_3();
        }

        private void V_CHECK_FLAG_4_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Init_Check_4();
        }

        private void IGR_BALANCE_ACCOUNT_SLIP_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BALANCE_ACCOUNT_SLIP.GetColumnToIndex("CHECK_YN"))
            {
                IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
                IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
                IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;
            }
        }

        private void IGR_BALANCE_ACCOUNT_R_SLIP_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BALANCE_ACCOUNT_R_SLIP.GetColumnToIndex("CHECK_YN"))
            {
                IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
                IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
                IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;
            }
        }

        private void BTN_SHOW_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(V_SLIP_NUM.EditValue) != string.Empty)
            {
                AssmblyRun_Manual("FCMF0202", V_GL_DATE.EditValue, V_SLIP_NUM.EditValue);
            }
        }
        
        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_BALANCE_ACCOUNT_SLIP.RowCount < 1)
            {
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            int vIDX_CHECK_YN = IGR_BALANCE_ACCOUNT_SLIP.GetColumnToIndex("CHECK_YN");
            for (int r = 0; r < IGR_BALANCE_ACCOUNT_SLIP.RowCount; r++)
            {
                if (iConv.ISNull(IGR_BALANCE_ACCOUNT_SLIP.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                {
                    IGR_CURR_ESTIMATE_SLIP.LastConfirmChanges();
                    IDA_CURR_ESTIMATE_SLIP.OraSelectData.AcceptChanges();
                    IDA_CURR_ESTIMATE_SLIP.Refillable = true;

                    //포커스이동
                    IGR_BALANCE_ACCOUNT_SLIP.CurrentCellMoveTo(r, vIDX_CHECK_YN);
                    IGR_BALANCE_ACCOUNT_SLIP.CurrentCellActivate(r, vIDX_CHECK_YN);

                    if (iConv.ISNull(V_ESTIMATE_DATE.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_ESTIMATE_DATE.Focus();
                        return;
                    }
                    if (iConv.ISNull(V_GL_DATE.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10187"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_GL_DATE.Focus();
                        return;
                    }
                    if (iConv.ISNull(V_ORG_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        W_ACCOUNT_CODE_3.Focus();
                        return;
                    }

                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSTATUS = "F";
                    string vMESSAGE = string.Empty;

                    IDC_SET_CURR_ESTIMATE_SLIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_CURR_ESTIMATE_SLIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_CURR_ESTIMATE_SLIP.GetCommandParamValue("O_MESSAGE"));

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    if (IDC_SET_CURR_ESTIMATE_SLIP.ExcuteError || vSTATUS == "F")
                    {
                        IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
                        IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
                        IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }

                    IGR_BALANCE_ACCOUNT_SLIP.SetCellValue("CHECK_YN", "N");
                }
            }
            IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
            IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
            IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_CURR_ESTIMATE_SLIP.Fill();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_BALANCE_ACCOUNT_SLIP.RowCount < 1)
            {
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            int vIDX_CHECK_YN = IGR_BALANCE_ACCOUNT_SLIP.GetColumnToIndex("CHECK_YN");
            for (int r = 0; r < IGR_BALANCE_ACCOUNT_SLIP.RowCount; r++)
            {
                if (iConv.ISNull(IGR_BALANCE_ACCOUNT_SLIP.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                {
                    IGR_CURR_ESTIMATE_SLIP.LastConfirmChanges();
                    IDA_CURR_ESTIMATE_SLIP.OraSelectData.AcceptChanges();
                    IDA_CURR_ESTIMATE_SLIP.Refillable = true;
                    
                    //포커스이동
                    IGR_BALANCE_ACCOUNT_SLIP.CurrentCellMoveTo(r, vIDX_CHECK_YN);
                    IGR_BALANCE_ACCOUNT_SLIP.CurrentCellActivate(r, vIDX_CHECK_YN); 

                    if (iConv.ISNull(V_ESTIMATE_DATE.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_ESTIMATE_DATE.Focus();
                        return;
                    }                    
                    if (iConv.ISNull(V_ORG_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        W_ACCOUNT_CODE_3.Focus();
                        return;
                    }

                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSTATUS = "F";
                    string vMESSAGE = string.Empty;

                    IDC_CANCEL_CURR_ESTIMATE_SLIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_CANCEL_CURR_ESTIMATE_SLIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_CANCEL_CURR_ESTIMATE_SLIP.GetCommandParamValue("O_MESSAGE"));

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    if (IDC_CANCEL_CURR_ESTIMATE_SLIP.ExcuteError || vSTATUS == "F")
                    {
                        IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
                        IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
                        IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }

                    IGR_BALANCE_ACCOUNT_SLIP.SetCellValue("CHECK_YN", "N");
                }
            }

            IGR_BALANCE_ACCOUNT_SLIP.LastConfirmChanges();
            IDA_BALANCE_ACCOUNT_SLIP.OraSelectData.AcceptChanges();
            IDA_BALANCE_ACCOUNT_SLIP.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_CURR_ESTIMATE_SLIP.Fill();
        }

        private void BTN_SHOW_R_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(V_R_SLIP_NUM.EditValue) != string.Empty)
            {
                AssmblyRun_Manual("FCMF0202", V_R_GL_DATE.EditValue, V_R_SLIP_NUM.EditValue);
            }
        }

        private void BTN_SET_R_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_BALANCE_ACCOUNT_R_SLIP.RowCount < 1)
            {
                return;
            }
            if (iConv.ISNull(V_R_GL_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_R_GL_DATE.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            int vIDX_CHECK_YN = IGR_BALANCE_ACCOUNT_R_SLIP.GetColumnToIndex("CHECK_YN");
            for (int r = 0; r < IGR_BALANCE_ACCOUNT_R_SLIP.RowCount; r++)
            {
                if (iConv.ISNull(IGR_BALANCE_ACCOUNT_R_SLIP.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                {
                    IGR_CURR_ESTIMATE_R_SLIP.LastConfirmChanges();
                    IDA_CURR_ESTIMATE_R_SLIP.OraSelectData.AcceptChanges();
                    IDA_CURR_ESTIMATE_R_SLIP.Refillable = true;

                    //포커스이동
                    IGR_BALANCE_ACCOUNT_R_SLIP.CurrentCellMoveTo(r, vIDX_CHECK_YN);
                    IGR_BALANCE_ACCOUNT_R_SLIP.CurrentCellActivate(r, vIDX_CHECK_YN);
                    if (iConv.ISNull(V_R_GL_DATE.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_R_GL_DATE.Focus();
                        return;
                    }
                    if (iDate.ISGetDate(V_R_GL_DATE.EditValue) < iDate.ISGetDate(V_R_ESTIMATE_DATE.EditValue))
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10598"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_R_GL_DATE.Focus();
                        return;
                    }
                    if (iConv.ISNull(V_R_ORG_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSTATUS = "F";
                    string vMESSAGE = string.Empty;
                    IDC_SET_CURR_ESTIMATE_R_SLIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_CURR_ESTIMATE_R_SLIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_CURR_ESTIMATE_R_SLIP.GetCommandParamValue("O_MESSAGE"));

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    if (IDC_SET_CURR_ESTIMATE_R_SLIP.ExcuteError || vSTATUS == "F")
                    {
                        IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
                        IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
                        IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }

                    IGR_BALANCE_ACCOUNT_R_SLIP.SetCellValue("CHECK_YN", "N");
                }
            }

            IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
            IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
            IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_CURR_ESTIMATE_R_SLIP.Fill();
        }

        private void BTN_CANCEL_R_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_BALANCE_ACCOUNT_R_SLIP.RowCount < 1)
            {
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            int vIDX_CHECK_YN = IGR_BALANCE_ACCOUNT_R_SLIP.GetColumnToIndex("CHECK_YN");
            for (int r = 0; r < IGR_BALANCE_ACCOUNT_R_SLIP.RowCount; r++)
            {
                if (iConv.ISNull(IGR_BALANCE_ACCOUNT_R_SLIP.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                {
                    IGR_CURR_ESTIMATE_R_SLIP.LastConfirmChanges();
                    IDA_CURR_ESTIMATE_R_SLIP.OraSelectData.AcceptChanges();
                    IDA_CURR_ESTIMATE_R_SLIP.Refillable = true;

                    //포커스이동
                    IGR_BALANCE_ACCOUNT_R_SLIP.CurrentCellMoveTo(r, vIDX_CHECK_YN);
                    IGR_BALANCE_ACCOUNT_R_SLIP.CurrentCellActivate(r, vIDX_CHECK_YN);

                    if (iConv.ISNull(V_R_ORG_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                        
                        return;
                    }
                    
                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSTATUS = "F";
                    string vMESSAGE = string.Empty;

                    IDC_CANCEL_CURR_ESTIMATE_R_SLIP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_CANCEL_CURR_ESTIMATE_R_SLIP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_CANCEL_CURR_ESTIMATE_R_SLIP.GetCommandParamValue("O_MESSAGE"));

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    if (IDC_CANCEL_CURR_ESTIMATE_R_SLIP.ExcuteError || vSTATUS == "F")
                    {
                        IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
                        IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
                        IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                    IGR_BALANCE_ACCOUNT_R_SLIP.SetCellValue("CHECK_YN", "N");
                }
            }

            IGR_BALANCE_ACCOUNT_R_SLIP.LastConfirmChanges();
            IDA_BALANCE_ACCOUNT_R_SLIP.OraSelectData.AcceptChanges();
            IDA_BALANCE_ACCOUNT_R_SLIP.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_CURR_ESTIMATE_R_SLIP.Fill();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_FR_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", null);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_FR_2_SelectedRowData(object pSender)
        {
            W_ACCOUNT_CODE_TO_2.EditValue = W_ACCOUNT_CODE_FR_2.EditValue;
            W_ACCOUNT_DESC_TO_2.EditValue = W_ACCOUNT_DESC_FR_2.EditValue;
        }

        private void ILA_ACCOUNT_CONTROL_TO_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE", W_ACCOUNT_CODE_FR_2.EditValue);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_3_SelectedRowData(object pSender)
        {
            SearchDB();
        }

        private void ILA_ACCOUNT_CONTROL_4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_4_SelectedRowData(object pSender)
        {
            SearchDB();
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_STATEMENT_EXCHANGE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["BALANCE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_STATEMENT_EXCHANGE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iConv.ISNull(e.Row["BALANCE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
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
                INIT_MANAGEMENT_COLUMN_2(-1);
                Application.DoEvents();
                return;
            }
            IDA_BALANCE_CURR_ESTIMATE.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", -1);
            IDA_BALANCE_CURR_ESTIMATE.Fill();

            INIT_MANAGEMENT_COLUMN_2(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            Application.DoEvents();

            IDA_BALANCE_CURR_ESTIMATE.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            IDA_BALANCE_CURR_ESTIMATE.Fill();
        }

        private void IDA_BALANCE_ACCOUNT_SLIP_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                INIT_MANAGEMENT_COLUMN_3(-1);
                Application.DoEvents();
                return;
            }

            IDA_CURR_ESTIMATE_SLIP.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            IDA_CURR_ESTIMATE_SLIP.Fill(); 

            //전표 차/대합계 
            SUM_SLIP_AMOUNT(); 

            INIT_MANAGEMENT_COLUMN_3(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            Application.DoEvents();

            IDA_CURR_ESTIMATE_SLIP.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            IDA_CURR_ESTIMATE_SLIP.Fill(); 

            //전표 차/대합계 
            SUM_SLIP_AMOUNT(); 
        }

        private void IDA_BALANCE_ACCOUNT_R_SLIP_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                INIT_MANAGEMENT_COLUMN_4(-1);
                Application.DoEvents();
                return;
            }

            IDA_CURR_ESTIMATE_R_SLIP.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", -1);
            IDA_CURR_ESTIMATE_R_SLIP.Fill();

            //전표 차/대합계 
            SUM_R_SLIP_AMOUNT(); 

            INIT_MANAGEMENT_COLUMN_4(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            Application.DoEvents();

            IDA_CURR_ESTIMATE_R_SLIP.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
            IDA_CURR_ESTIMATE_R_SLIP.Fill(); 

            //전표 차/대합계 
            SUM_R_SLIP_AMOUNT(); 
        }

        #endregion

    }
}