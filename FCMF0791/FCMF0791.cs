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

namespace FCMF0791
{
    public partial class FCMF0791 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private bool mIsSearch = false;
        private int mPageTotal = 0;

        #endregion;

        #region ----- Constructor -----

        public FCMF0791()
        {
            InitializeComponent();
        }

        public FCMF0791(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Search_DB()
        {
            IDA_VOUCH_BS.Fill();
            IDA_VOUCH_IS.Fill();
            IDA_VOUCH_MS.Fill();
            IDA_VOUCH_SUM.Fill();
        }

        private void Show_Detail(object pFS_FORM_TYPE, object pITEM_CODE
                                , object pGL_DATE_FR, object pGL_DATE_TO)
        {
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();
            Application.DoEvents(); 

            IDA_VOUCH_LIST.SetSelectParamValue("P_FS_FORM_TYPE", pFS_FORM_TYPE);
            IDA_VOUCH_LIST.SetSelectParamValue("P_ITEM_CODE", pITEM_CODE);
            IDA_VOUCH_LIST.SetSelectParamValue("P_GL_DATE_FR", pGL_DATE_FR);
            IDA_VOUCH_LIST.SetSelectParamValue("P_GL_DATE_TO", pGL_DATE_TO);
            IDA_VOUCH_LIST.SetSelectParamValue("P_ALL_VIEW_FLAG", "N");
            IDA_VOUCH_LIST.Fill();

            IGR_VOUCH_LIST.Focus();
        }

        private void Show_Slip_Detail()
        {
            try
            {
                int mSLIP_HEADER_ID = iString.ISNumtoZero(IGR_VOUCH_LIST.GetCellValue("SLIP_HEADER_ID"));
                if (mSLIP_HEADER_ID != Convert.ToInt32(0))
                {
                    Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    FCMF0205.FCMF0205 vFCMF0205 = new FCMF0205.FCMF0205(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
                    vFCMF0205.Show();

                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.UseWaitCursor = false;
                }
            }
            catch
            {
            }
        }
        
        #endregion;


        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID, 
                                        object pACCOUNT_CONTROL_ID, object pACCOUNT_CODE, object pACCOUNT_DESC,
                                        object pGL_DATE_FR, object pGL_DATE_TO)
        {
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vCurrAssemblyFileVersion = string.Empty;
            object vFS_TYPE_ID = W_ACCOUNT_CONTROL_ID.EditValue;
            IDC_GET_FS_TYPE_NAME.SetCommandParamValue("W_COMMON_ID", vFS_TYPE_ID);
            IDC_GET_FS_TYPE_NAME.ExecuteNonQuery();
            object vFS_TYPE_NAME = IDC_GET_FS_TYPE_NAME.GetCommandParamValue("O_RETURN_VALUE");

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

                            object[] vParam = new object[12];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = "지출증빙명세서";
                            vParam[3] = "0";
                            vParam[4] = pACCOUNT_CODE;
                            vParam[5] = "0";
                            vParam[6] = pACCOUNT_DESC;
                            vParam[7] = pACCOUNT_CODE;
                            vParam[8] = iDate.ISMonth_1st(W_GL_DATE_FR.EditValue);
                            vParam[9] = iDate.ISMonth_Last(W_GL_DATE_TO.EditValue);
                            vParam[10] = "VOUCH";
                            vParam[11] = pACCOUNT_CONTROL_ID;  

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

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx vGrid)
        {
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.RestoreDirectory = true;
            vSaveFileDialog.Filter = "Excel file(*.xls)|*.xls";
            vSaveFileDialog.DefaultExt = "xls";

            if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                vExport.GridToExcel(vGrid.BaseGrid, vSaveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = vSaveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

        #region ----- MDi Main ToolBar Button Event -----

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
                    if (IDA_VOUCH_LIST.IsFocused)
                    {
                        IDA_VOUCH_LIST.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_VOUCH_LIST.IsFocused)
                    {
                        IDA_VOUCH_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_VOUCH("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (TB_MAIN.SelectedTab.TabIndex == TP_DETAIL.TabIndex)
                    {
                        ExcelExport(IGR_VOUCH_LIST);
                    }
                    else
                    {
                        XLPrinting_VOUCH("FILE");
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0791_Load(object sender, EventArgs e)
        {
            int vYear = System.DateTime.Today.Year;
            System.DateTime vDate = new System.DateTime(vYear, 1, 1);
            W_GL_DATE_FR.EditValue = vDate;
            W_GL_DATE_TO.EditValue = iDate.ISMonth_Last(DateTime.Today);
        }

        private void FCMF0791_Shown(object sender, EventArgs e)
        {
            W_ACCOUNT_CODE.Focus();
        }
         
        private void IGR_VOUCH_BS_CellDoubleClick(object pSender)
        {
            if (IGR_VOUCH_BS.RowIndex < 0)
            {
                return;
            }

            object vFS_FORM_TYPE = IGR_VOUCH_BS.GetCellValue("FS_FORM_TYPE");
            object vITEM_CODE = IGR_VOUCH_BS.GetCellValue("ITEM_CODE");
            Show_Detail(vFS_FORM_TYPE,vITEM_CODE, W_GL_DATE_FR.EditValue, W_GL_DATE_TO.EditValue);

            //if (iString.ISNull(vACCOUNT_CONTROL_ID) != string.Empty)
            //{
            //    AssmblyRun_Manual("FCMF0295", vACCOUNT_CONTROL_ID, IGR_VOUCH_BS.GetCellValue("ACCOUNT_CODE"), IGR_VOUCH_BS.GetCellValue("ACCOUNT_DESC"),
            //                        W_GL_DATE_FR.EditValue, W_GL_DATE_TO.EditValue);
            //}
        }

        private void IGR_VOUCH_IS_CellDoubleClick(object pSender)
        {
            if (IGR_VOUCH_IS.RowIndex < 0)
            {
                return;
            }

            object vFS_FORM_TYPE = IGR_VOUCH_IS.GetCellValue("FS_FORM_TYPE");
            object vITEM_CODE = IGR_VOUCH_IS.GetCellValue("ITEM_CODE");
            Show_Detail(vFS_FORM_TYPE, vITEM_CODE, W_GL_DATE_FR.EditValue, W_GL_DATE_TO.EditValue); 

            //if (iString.ISNull(vACCOUNT_CONTROL_ID) != string.Empty)
            //{
            //    AssmblyRun_Manual("FCMF0295", vACCOUNT_CONTROL_ID, IGR_VOUCH_IS.GetCellValue("ACCOUNT_CODE"), IGR_VOUCH_IS.GetCellValue("ACCOUNT_DESC"),
            //                        W_GL_DATE_FR.EditValue, W_GL_DATE_TO.EditValue);
            //}
        }

        private void IGR_VOUCH_MS_CellDoubleClick(object pSender)
        {
            if (IGR_VOUCH_MS.RowIndex < 0)
            {
                return;
            }

            object vFS_FORM_TYPE = IGR_VOUCH_MS.GetCellValue("FS_FORM_TYPE");
            object vITEM_CODE = IGR_VOUCH_MS.GetCellValue("ITEM_CODE");
            Show_Detail(vFS_FORM_TYPE, vITEM_CODE, W_GL_DATE_FR.EditValue, W_GL_DATE_TO.EditValue);  

            //if (iString.ISNull(vACCOUNT_CONTROL_ID) != string.Empty)
            //{
            //    AssmblyRun_Manual("FCMF0295", vACCOUNT_CONTROL_ID, IGR_VOUCH_MS.GetCellValue("ACCOUNT_CODE"), IGR_VOUCH_MS.GetCellValue("ACCOUNT_DESC"),
            //                        W_GL_DATE_FR.EditValue, W_GL_DATE_TO.EditValue);
            //}
        }

        private void IGR_VOUCH_LIST_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        private void BTN_INQUIRY_DTL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_VOUCH_LIST.SetSelectParamValue("P_GL_DATE_FR", W_GL_DATE_FR.EditValue);
            IDA_VOUCH_LIST.SetSelectParamValue("P_GL_DATE_TO", W_GL_DATE_TO.EditValue);
            IDA_VOUCH_LIST.SetSelectParamValue("P_ALL_VIEW_FLAG", V_ALL_VIEW_FLAG.CheckBoxValue);
            IDA_VOUCH_LIST.Fill();

            IGR_VOUCH_LIST.Focus();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaFS_SET_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_SET", "Y");
        }

        private void ILA_VOUCH_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "VOUCH_CODE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");  
        }

        #endregion

        #region ----- EditAdv Event -----

        private void PERIOD_TO_EditValueChanged(object pSender)
        {
            mIsSearch = false;
            object vObject = W_GL_DATE_TO.EditValue;
            string vDate = iString.ISNull(vObject);
            int vLength = vDate.Length;

            string vYear = string.Empty;
            string vYearMonth = string.Empty;

            if (vLength == 7)
            {
                vYear = vDate.Substring(0, 4);
                vYearMonth = string.Format("{0}-01", vYear);

                W_GL_DATE_FR.EditValue = vYearMonth;

                mIsSearch = true;
            }
        }

        #endregion

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice, InfoSummit.Win.ControlAdv.ISGridAdvEx p_grid_TRIAL_BALANCE)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = p_grid_TRIAL_BALANCE.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            idaPRINT_TITLE.Fill();

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting_1 xlPrinting = new XLPrinting_1(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                vMessageText = string.Format("XL Opening...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0791_001.xls";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    vPageNumber = xlPrinting.LineWrite(p_grid_TRIAL_BALANCE, idaPRINT_TITLE);

                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE("TRIAL_");
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
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

        private void XLPrinting_VOUCH(string pOutput_Type)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSaveFileName = string.Empty;
            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            if (pOutput_Type == "PRINT")
            {

            }
            else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                //파일명 지정//
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
                saveFileDialog1.DefaultExt = "xls";
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
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            vMessageText = string.Format("Printing Starting", vPageTotal);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0791_001.xls";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    string vDate_Fr = string.Format("{0}. {1:D2}. {2:D2}", iDate.ISGetDate(W_GL_DATE_FR.EditValue).Year,
                                                                            iDate.ISGetDate(W_GL_DATE_FR.EditValue).Month,
                                                                            iDate.ISGetDate(W_GL_DATE_FR.EditValue).Day);
                    string vDate_To = string.Format("{0}. {1:D2}. {2:D2}", iDate.ISGetDate(W_GL_DATE_TO.EditValue).Year,
                                                                            iDate.ISGetDate(W_GL_DATE_TO.EditValue).Month,
                                                                            iDate.ISGetDate(W_GL_DATE_TO.EditValue).Day);

                    IDC_DV_TAX_VALUE.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
                    IDC_DV_TAX_VALUE.ExecuteNonQuery();
                    object vCORP_NAME = IDC_DV_TAX_VALUE.GetCommandParamValue("O_CODE_NAME");
                    object vTAX_REG_NO = IDC_DV_TAX_VALUE.GetCommandParamValue("O_VALUE1"); 


                    //1.표준대차대조표                    
                    if (IGR_VOUCH_BS.RowCount != 0)
                    {
                        vPageNumber = xlPrinting.LineWrite(xlPrinting, vDate_Fr, vDate_To, vCORP_NAME, vTAX_REG_NO, 
                                                            IDA_VOUCH_BS, IDA_VOUCH_IS, IDA_VOUCH_MS, IDA_VOUCH_SUM);
                    }
                    mPageTotal = vPageNumber;
                                        
                     
                    vMessageText = string.Format("Printing Completed.. ", vPageTotal);
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }

                if (pOutput_Type == "PRINT")
                {
                    xlPrinting.PreView(1, mPageTotal);
                }
                else
                {
                    //-------------------------------------------------------------------------
                    xlPrinting.Save(vSaveFileName); //SAVE
                    //-------------------------------------------------------------------------
                }
                xlPrinting.Dispose();
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                string vMessage = ex.Message;
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }

            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (pOutput_Type == "PRINT")
            {

            }
            else if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10173"), "Qeustion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(vSaveFileName);
            }
        }

        #endregion;

    }
}