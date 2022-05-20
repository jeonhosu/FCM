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

namespace FCMF0316
{
    public partial class FCMF0316 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        string mBase_Currency_Code;

        #endregion;

        #region ----- Constructor -----

        public FCMF0316()
        {
            InitializeComponent();
        }

        public FCMF0316(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void GetAccountBook()
        {
            IDC_BASE_CURRENCY.ExecuteNonQuery();
            mBase_Currency_Code = iConv.ISNull(IDC_BASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE"));
        }

        private void Search_DB()
        {
            if (iConv.ISNull(W_SALE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_SALE_DATE_FR.Focus();
                return;
            }

            if (iConv.ISNull(W_SALE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_SALE_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(W_SALE_DATE_FR.EditValue) > Convert.ToDateTime(W_SALE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_SALE_DATE_FR.Focus();
                return;
            }

            string vSALE_NUM = iConv.ISNull(IGR_ASSET_SALE_SLIP_LIST.GetCellValue("SALE_NUM"));
            int vCOL_IDX = IGR_ASSET_SALE_SLIP_LIST.GetColumnToIndex("SALE_NUM");
            IDA_ASSET_SALE_SLIP_LIST.Fill();
            if (iConv.ISNull(vSALE_NUM) != string.Empty)
            {
                for (int i = 0; i < IGR_ASSET_SALE_SLIP_LIST.RowCount; i++)
                {
                    if (vSALE_NUM == iConv.ISNull(IGR_ASSET_SALE_SLIP_LIST.GetCellValue(i, vCOL_IDX)))
                    {
                        IGR_ASSET_SALE_SLIP_LIST.CurrentCellMoveTo(i, vCOL_IDX);
                        IGR_ASSET_SALE_SLIP_LIST.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
        }

        private void Search_DB_DTL()
        {
            IGR_ASSET_SALE_LINE.LastConfirmChanges();
            IDA_ASSET_SALE_SLIP_LINE.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_SLIP_LINE.Refillable = true;

            IDA_ASSET_SALE_SLIP_HEADER.OraSelectData.AcceptChanges();
            IDA_ASSET_SALE_SLIP_HEADER.Refillable = true;

            IDA_ASSET_SALE_SLIP_HEADER.Fill();

            IDA_ASSET_SALE_SLIP_AR.Fill();
            IDA_ASSET_SALE_SLIP_VAT.Fill();
            IDA_ASSET_SALE_SLIP_ETC.Fill();
            IDC_GET_GAP_SALE_AMOUNT.ExecuteNonQuery();
        }

        private bool Save_Slip_Date()
        {
            if (iConv.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_DATE.Focus();
                return false; 
            }

            IDC_SAVE_ASSET_SALE_SLIP_HEADER.SetCommandParamValue("W_SALE_HEADER_ID", SALE_HEADER_ID.EditValue);
            IDC_SAVE_ASSET_SALE_SLIP_HEADER.SetCommandParamValue("P_SLIP_DATE", SLIP_DATE.EditValue);
            IDC_SAVE_ASSET_SALE_SLIP_HEADER.SetCommandParamValue("P_DESCRIPTION", DESCRIPTION.EditValue);
            IDC_SAVE_ASSET_SALE_SLIP_HEADER.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SAVE_ASSET_SALE_SLIP_HEADER.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SAVE_ASSET_SALE_SLIP_HEADER.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SAVE_ASSET_SALE_SLIP_HEADER.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SAVE_ASSET_SALE_SLIP_HEADER.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SLIP_DATE.Focus();
                return false; 
            }
            else if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SLIP_DATE.Focus();
                return false; 
            }  
            return true;
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

        private void AssmblyRun_Manual(object pAssembly_ID)
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
                            if (iConv.ISNull(pAssembly_ID) == "FCMF0206")
                            {
                                vParam = new object[5];
                                vParam[0] = this.MdiParent;
                                vParam[1] = isAppInterfaceAdv1.AppInterface;
                                vParam[2] = SLIP_DATE.EditValue;
                                vParam[3] = SLIP_DATE.EditValue; 
                                vParam[4] = SLIP_NUM.EditValue; 
                            }
                            else
                            {
                                vParam = new object[6];
                                vParam[0] = this.MdiParent;
                                vParam[1] = isAppInterfaceAdv1.AppInterface;
                                vParam[2] = SLIP_DATE.EditValue;
                                vParam[3] = SLIP_DATE.EditValue;
                                vParam[4] = DBNull.Value;
                                vParam[5] = SLIP_NUM.EditValue;
                            }
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


        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    if (TB_MAIN.SelectedTab.TabIndex == TP_DETAIL.TabIndex)
                    {
                        Search_DB_DTL();
                    }
                    else
                    {
                        Search_DB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (Save_Slip_Date() == true)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10010"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_SALE_SLIP_HEADER.IsFocused)
                    {
                        IDA_ASSET_SALE_SLIP_HEADER.Cancel();
                    }
                    else if (IDA_ASSET_SALE_SLIP_LINE.IsFocused)
                    {
                        IDA_ASSET_SALE_SLIP_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event ----

        private void FCMF0316_Load(object sender, EventArgs e)
        {
            W_SALE_DATE_FR.EditValue = iDate.ISYearMonth(iDate.ISGetDate());
            W_SALE_DATE_TO.EditValue = iDate.ISGetDate();

            GetAccountBook();

            BTN_SLIP_OK.BringToFront();
            BTN_SLIP_CANCEL.BringToFront();
            BTN_DPR_VIEW.BringToFront();
            V_GAP_AMOUNT.BringToFront();

            BTN_AR_ACC_CODE.Enabled = false;
            BTN_VAT_ACC.Enabled = false;
            BTN_ETC_ACC.Enabled = false;

            IDA_ASSET_SALE_SLIP_HEADER.FillSchema();
        }

        private void FCMF0316_Shown(object sender, EventArgs e)
        {
            V_RB_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SLIP_IF_FLAG.EditValue = V_RB_NO.RadioCheckedString;

            TB_MAIN.SelectedIndex = (TP_LIST.TabIndex - 1);
            TB_MAIN.SelectedTab.Focus();
        }

        private void V_RB_NO_Click(object sender, EventArgs e)
        {
            if (V_RB_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_NO.RadioCheckedString;
            }
        }

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            if (V_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_ALL.RadioCheckedString;
            }
        }

        private void V_RB_YES_Click(object sender, EventArgs e)
        {
            if (V_RB_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SLIP_IF_FLAG.EditValue = V_RB_YES.RadioCheckedString;
            }
        }

        private void IGR_ASSET_SALE_DTL_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_ASSET_SALE_SLIP_LIST.RowIndex < 0)
            {
                return; 
            }

            W_SALE_HEADER_ID.EditValue = IGR_ASSET_SALE_SLIP_LIST.GetCellValue("SALE_HEADER_ID");

            TB_MAIN.SelectedIndex = (TP_DETAIL.TabIndex -1);
            TB_MAIN.SelectedTab.Focus();
            Application.DoEvents();

            Search_DB_DTL();
        } 

        private void BTN_CONFIRM_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //
            IDA_ASSET_SALE_SLIP_HEADER.Update();

            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            } 
            
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_SLIP.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_SLIP.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_SLIP.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_SLIP.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

            Search_DB_DTL();
        }

        private void BTN_CONFIRM_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_SLIP.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_CANCEL_SLIP.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_CANCEL_SLIP.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_SLIP.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_CANCEL_SLIP.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

            Search_DB_DTL();
        }
        
        private void BTN_EXEC_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SLIP_NUM.EditValue) != string.Empty)
            {
                if (iConv.ISNull(SLIP_TABLE.EditValue) == "FI_SLIP_HEADER_BUDGET")
                {
                    AssmblyRun_Manual("FCMF0206");
                }
                else
                {
                    AssmblyRun_Manual("FCMF0202");
                }
            }
        }

        private void BTN_DPR_VIEW_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISDecimaltoZero(IGR_ASSET_SALE_LINE.GetCellValue("ASSET_ID"),0) != 0)
            {
                DialogResult vResult = DialogResult.None;
                FCMF0316_DPR vFCMF0316_DPR = new FCMF0316_DPR(MdiParent, isAppInterfaceAdv1.AppInterface, SALE_HEADER_ID.EditValue
                                                            , IGR_ASSET_SALE_LINE.GetCellValue("ASSET_ID")
                                                            , IGR_ASSET_SALE_LINE.GetCellValue("ASSET_CODE")
                                                            , IGR_ASSET_SALE_LINE.GetCellValue("ASSET_DESC")
                                                            , DPR_TYPE.EditValue, DPR_TYPE_DESC.EditValue);
                mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0316_DPR, isAppInterfaceAdv1.AppInterface);
                vResult = vFCMF0316_DPR.ShowDialog();
                if (vResult == DialogResult.OK)
                {
                     
                }
                vFCMF0316_DPR.Dispose();
            }
        }

        private void BTN_AR_ACC_CODE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SALE_NUM.Focus();
                return;
            } 
            if (iConv.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_DATE.Focus(); 
                return;  
            }

            if (Save_Slip_Date() == false)
            {
                return;
            }

            decimal vSALE_AMOUNT = iConv.ISDecimaltoZero(SALE_AMOUNT.EditValue, 0) +
                                    iConv.ISDecimaltoZero(SALE_VAT_AMOUNT.EditValue, 0);
            DialogResult vResult = DialogResult.None;
            FCMF0316_ACCOUNT vFCMF0316_ACCOUNT = new FCMF0316_ACCOUNT(MdiParent, isAppInterfaceAdv1.AppInterface, SALE_HEADER_ID.EditValue
                                                                    , "AR", SLIP_DATE.EditValue, DEPT_ID.EditValue, DEPT_NAME.EditValue
                                                                    , DEPT_ID.EditValue, DEPT_ID.EditValue, CURRENCY_CODE.EditValue
                                                                    , EXCHANGE_RATE.EditValue, CURR_AMOUNT.EditValue, vSALE_AMOUNT
                                                                    , VENDOR_CODE.EditValue, VENDOR_NAME.EditValue
                                                                    , VAT_TAX_TYPE.EditValue, VAT_TAX_TYPE_DESC.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0316_ACCOUNT, isAppInterfaceAdv1.AppInterface);
            vResult = vFCMF0316_ACCOUNT.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                IDA_ASSET_SALE_SLIP_AR.Fill();
                IDA_ASSET_SALE_SLIP_LINE.Fill();
                IDC_GET_GAP_SALE_AMOUNT.ExecuteNonQuery();
            }
            vFCMF0316_ACCOUNT.Dispose();
        }

        private void BTN_VAT_ACC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SALE_NUM.Focus();
                return;
            } 
            if (iConv.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_DATE.Focus();
                return;
            }
            if (Save_Slip_Date() == false)
            {
                return;
            }

            DialogResult vResult = DialogResult.None;
            FCMF0316_ACCOUNT vFCMF0316_ACCOUNT = new FCMF0316_ACCOUNT(MdiParent, isAppInterfaceAdv1.AppInterface, SALE_HEADER_ID.EditValue
                                                                    , "VAT", SLIP_DATE.EditValue, DEPT_ID.EditValue, DEPT_NAME.EditValue
                                                                    , DEPT_ID.EditValue, DEPT_ID.EditValue, CURRENCY_CODE.EditValue
                                                                    , EXCHANGE_RATE.EditValue, CURR_AMOUNT.EditValue, SALE_VAT_AMOUNT.EditValue
                                                                    , VENDOR_CODE.EditValue, VENDOR_NAME.EditValue
                                                                    , VAT_TAX_TYPE.EditValue, VAT_TAX_TYPE_DESC.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0316_ACCOUNT, isAppInterfaceAdv1.AppInterface);
            vResult = vFCMF0316_ACCOUNT.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                IDA_ASSET_SALE_SLIP_VAT.Fill();
                IDA_ASSET_SALE_SLIP_LINE.Fill();
                IDC_GET_GAP_SALE_AMOUNT.ExecuteNonQuery();
            }
            vFCMF0316_ACCOUNT.Dispose(); 
        }

        private void BTN_ETC_ACC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SALE_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SALE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SALE_NUM.Focus();
                return;
            }
            if (iConv.ISNull(SLIP_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLIP_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SLIP_DATE.Focus();
                return;
            }
            if (Save_Slip_Date() == false)
            {
                return;
            }

            object vCURR_AMOUNT = 0;
            if(mBase_Currency_Code != iConv.ISNull(CURRENCY_CODE.EditValue))
            {
                vCURR_AMOUNT = Math.Round(iConv.ISDecimaltoZero(ETC_AMOUNT.EditValue) / iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue), 2);
            }
            DialogResult vResult = DialogResult.None;
            FCMF0316_ACCOUNT vFCMF0316_ACCOUNT = new FCMF0316_ACCOUNT(MdiParent, isAppInterfaceAdv1.AppInterface, SALE_HEADER_ID.EditValue
                                                                    , "ETC", SLIP_DATE.EditValue, DEPT_ID.EditValue, DEPT_NAME.EditValue
                                                                    , DEPT_ID.EditValue, DEPT_ID.EditValue, CURRENCY_CODE.EditValue
                                                                    , EXCHANGE_RATE.EditValue, vCURR_AMOUNT, ETC_AMOUNT.EditValue
                                                                    , VENDOR_CODE.EditValue, VENDOR_NAME.EditValue
                                                                    , VAT_TAX_TYPE.EditValue, VAT_TAX_TYPE_DESC.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0316_ACCOUNT, isAppInterfaceAdv1.AppInterface);
            vResult = vFCMF0316_ACCOUNT.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                IDA_ASSET_SALE_SLIP_ETC.Fill();
                IDA_ASSET_SALE_SLIP_LINE.Fill();
                IDC_GET_GAP_SALE_AMOUNT.ExecuteNonQuery();
            }
            vFCMF0316_ACCOUNT.Dispose();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ASSET_SALE_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_SALE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_ASSET_SALE_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ASSET_SALE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_VAT_TAX_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACC_CONTROL_FR_TO.SetLookupParamValue("W_AP_AR_TYPE", "AP");
            ILD_ACC_CONTROL_FR_TO.SetLookupParamValue("W_ENABLED_FLAG", "Y"); 
        }

        private void ILA_VAT_TAX_TYPE_SelectedRowData(object pSender)
        {
             
        }

        private void ILA_VENDOR_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }	 
        
        private void ILA_VENDOR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VENDOR_LIST.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_STATUS.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_DEPT_CODE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT_CODE.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_CURRENCY_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CURRENCY.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ILD_CURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CURRENCY_SelectedRowData(object pSender)
        {
             
        }

        private void ILA_ACC_CONTROL_FR_TO_AR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACC_CONTROL_FR_TO.SetLookupParamValue("W_ACCOUNT_CODE_FR", DBNull.Value);
            ILD_ACC_CONTROL_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACC_CONTROL_FR_TO_VAT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ACC_CONTROL_FR_TO.SetLookupParamValue("W_ACCOUNT_CODE_FR", DBNull.Value);
            ILD_ACC_CONTROL_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----
        
        private void IDA_ASSET_SALE_SLIP_HEADER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (iConv.ISNull(STATUS.EditValue) == "CONFIRM")
            {
                BTN_AR_ACC_CODE.Enabled = true;
                BTN_VAT_ACC.Enabled = true;
                BTN_ETC_ACC.Enabled = true;
            }
            else
            {
                BTN_AR_ACC_CODE.Enabled = false;
                BTN_VAT_ACC.Enabled = false;
                BTN_ETC_ACC.Enabled = false;
            }
        }



        #endregion

    }
}