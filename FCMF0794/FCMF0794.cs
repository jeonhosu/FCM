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

using System.IO;
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace FCMF0794
{
    public partial class FCMF0794 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        string vMULTI_LANG_FLAG = "N";

        #endregion;


        #region ----- Constructor -----

        public FCMF0794(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search()
        {

            //처리기간은 필수입니다.
            if (iString.ISNull(V_FS_FORM_TYPE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10529"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_FS_FORM_TYPE_DESC.Focus();
                return;
            }

            if (vMULTI_LANG_FLAG == "Y" && iString.ISNull(V_LANG_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10004"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_LANG_DESC.Focus();
                return;
            }


            if (iString.ISNull(W1_PERIOD_YEAR.EditValue) == string.Empty)
            {
                //년도는 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W1_PERIOD_YEAR.Focus();
                return;
            }

            //조회기간관련 정합성을 체크한다. 기간은 필수 조회조건이다.
            if (iString.ISNull(W1_MONTH_FR.EditValue) == string.Empty)
            {
                //시작기간은 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10548"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W1_MONTH_DESC_FR.Focus();
                return;
            }

            if (iString.ISNull(W1_MONTH_TO.EditValue) == string.Empty)
            {
                //종료기간은 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10549"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W1_MONTH_DESC_TO.Focus();
                return;
            }

            if (ConvertInteger(iString.ISNull(W1_MONTH_FR.EditValue)) > ConvertInteger(iString.ISNull(W1_MONTH_TO.EditValue)))
            {
                //종료기간은 시작기간 보다 커야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10345"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W1_MONTH_DESC_TO.Focus();
                return;
            }

            //출력구분은 필수사항입니다.
            if (iString.ISNull(W1_ACCOUNT_LEVEL.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10550"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W1_ACCOUNT_LEVEL_NAME.Focus();
                return;
            }  

            IDA_MS_MONTH.Fill();
            IGR_MS_MONTH.Focus(); 
        }

        #endregion;

        #region ----- Convert decimal  Method ----

            private int ConvertInteger(object pObject)
            {
                bool vIsConvert = false;
                int vConvertInteger = 0;

                try
                {
                    if (pObject != null)
                    {
                        vIsConvert = pObject is string;
                        if (vIsConvert == true)
                        {
                            string vString = pObject as string;
                            vConvertInteger = int.Parse(vString);
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                    System.Windows.Forms.Application.DoEvents();
                }

                return vConvertInteger;
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

        //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_NAME_L))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
                GridExcelConverterControl converter = new GridExcelConverterControl();
                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2007;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2007;
                IWorksheet sheet = workBook.Worksheets[0];
                //used to convert grid to excel 
                converter.GridToExcel(pGrid.BaseGrid, sheet, ConverterOptions.ColumnHeaders);
                //used to save the file
                workBook.SaveAs(saveFileDialog.FileName);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                        "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID, DateTime pPERIOD_DATE_FR, DateTime pPERIOD_DATE_TO,
                                        object pITEM_CODE, object pITEM_NAME, object pACCOUNT_LEVEL)
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

                            //    Form pMainForm, ISAppInterface pAppInterface,
                            //object pFS_FORM_TYPE_NAME, object pFS_FORM_TYPE_ID,
                            //object pFS_TYPE, object pACCOUNT_LEVEL, 
                            //object pFORM_ITEM_NAME, object pFORM_ITEM_CODE,
                            //object pGL_DATE_FR, object pGL_DATE_TO) 

                            object[] vParam = new object[10];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = V_FS_FORM_TYPE_DESC.EditValue;
                            vParam[3] = V_FS_FORM_TYPE_ID.EditValue;
                            vParam[4] = V_FS_TYPE.EditValue;
                            vParam[5] = pACCOUNT_LEVEL;
                            vParam[6] = iString.ISNull(pITEM_NAME).Trim();
                            vParam[7] = pITEM_CODE;
                            vParam[8] = pPERIOD_DATE_FR;
                            vParam[9] = pPERIOD_DATE_TO;

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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder) //아래에 새레코드 추가
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)   //저장
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)   //취소
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)   //삭제
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)    //인쇄
                {
                    //XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                {
                    ExcelExport(IGR_MS_MONTH); 
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0794_Load(object sender, EventArgs e)
        {
            IDC_GET_DEFAULT_FS_FORM_TYPE.ExecuteNonQuery();
            V_FS_FORM_TYPE_ID.EditValue = IDC_GET_DEFAULT_FS_FORM_TYPE.GetCommandParamValue("O_FS_FORM_TYPE_ID");
            V_FS_FORM_TYPE.EditValue = IDC_GET_DEFAULT_FS_FORM_TYPE.GetCommandParamValue("O_FS_FORM_TYPE");
            V_FS_FORM_TYPE_DESC.EditValue = IDC_GET_DEFAULT_FS_FORM_TYPE.GetCommandParamValue("O_FS_FORM_TYPE_DESC");
             
            IDC_GET_ACCOUNT_LEVEL.ExecuteNonQuery();
            W1_ACCOUNT_LEVEL_NAME.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE_NAME");
            W1_ACCOUNT_LEVEL.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE"); 

            IDC_GET_MULTI_LANG_P.ExecuteNonQuery();
            vMULTI_LANG_FLAG = iString.ISNull(IDC_GET_MULTI_LANG_P.GetCommandParamValue("O_MULTI_LANG_FLAG"));
            if (vMULTI_LANG_FLAG == "Y")
            {
                V_LANG_DESC.Visible = true;
                V_LANG_DESC.BringToFront();
                IDC_GET_LANG_CODE.ExecuteNonQuery();
                V_LANG_DESC.EditValue = IDC_GET_LANG_CODE.GetCommandParamValue("O_LANG_DESC");
                V_LANG_CODE.EditValue = IDC_GET_LANG_CODE.GetCommandParamValue("O_LANG_CODE");
            }
            else
            {
                V_LANG_DESC.Visible = false;
                V_LANG_CODE.EditValue = null;
            }

            W1_PERIOD_YEAR.EditValue = iDate.ISYear(DateTime.Today); 

            //월별//
            IDC_DEFAULT_VALUE2_W.SetCommandParamValue("W_WHERE", "GROUP_CODE = 'MONTH' AND CODE = '01'");
            IDC_DEFAULT_VALUE2_W.ExecuteNonQuery();
            W1_MONTH_DESC_FR.EditValue = IDC_DEFAULT_VALUE2_W.GetCommandParamValue("O_CODE_NAME");
            W1_MONTH_FR.EditValue = IDC_DEFAULT_VALUE2_W.GetCommandParamValue("O_CODE");

            IDC_DEFAULT_VALUE2_W.SetCommandParamValue("W_WHERE", "GROUP_CODE = 'MONTH' AND CODE = '12'");
            IDC_DEFAULT_VALUE2_W.ExecuteNonQuery();
            W1_MONTH_DESC_TO.EditValue = IDC_DEFAULT_VALUE2_W.GetCommandParamValue("O_CODE_NAME");
            W1_MONTH_TO.EditValue = IDC_DEFAULT_VALUE2_W.GetCommandParamValue("O_CODE");  
        }

        private void IGR_MS_MONTH_CellDoubleClick(object pSender)
        {
            //월
            if (IGR_MS_MONTH.RowIndex < 0)
            {
                return;
            }

            int vIDX_Month = IGR_MS_MONTH.ColIndex;
            DateTime vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-01", W1_PERIOD_YEAR.EditValue));
            DateTime vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-12", W1_PERIOD_YEAR.EditValue));
            if (vIDX_Month >= 2 && vIDX_Month <= 6)  //합계
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-{1}", W1_PERIOD_YEAR.EditValue, W1_MONTH_FR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-{1}", W1_PERIOD_YEAR.EditValue, W1_MONTH_TO.EditValue));
            }
            else if (vIDX_Month >= 7 && vIDX_Month <= 11)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-01", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-01", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 12 && vIDX_Month <= 16)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-02", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-02", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 17 && vIDX_Month <= 21)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-03", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-03", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 22 && vIDX_Month <= 26)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-04", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-04", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 27 && vIDX_Month <= 31)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-05", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-05", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 32 && vIDX_Month <= 36)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-06", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-06", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 37 && vIDX_Month <= 41)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-07", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-07", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 42 && vIDX_Month <= 46)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-08", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-08", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 47 && vIDX_Month <= 51)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-09", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-09", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 52 && vIDX_Month <= 56)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-10", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-10", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 57 && vIDX_Month <= 61)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-11", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-11", W1_PERIOD_YEAR.EditValue));
            }
            else if (vIDX_Month >= 62 && vIDX_Month <= 66)
            {
                vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-12", W1_PERIOD_YEAR.EditValue));
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-12", W1_PERIOD_YEAR.EditValue));
            }
            object vITEM_CODE = IGR_MS_MONTH.GetCellValue("HEADER_CODE");
            object vITEM_NAME = IGR_MS_MONTH.GetCellValue("ITEM_NAME");
            if (iString.ISNull(vITEM_CODE) != string.Empty)
            {
                AssmblyRun_Manual("FCMF0295", vPERIOD_DATE_FR, vPERIOD_DATE_TO, vITEM_CODE, vITEM_NAME, W1_ACCOUNT_LEVEL.EditValue);
            }
        }
         
        #endregion

        #region ----- Lookup event ----- 
          
        private void ILA_YEAR_YYYY_W1_RefreshLookupData_1(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_YEAR_YYYY.SetLookupParamValue("W_END_YEAR", iDate.ISDate_Month_Add(DateTime.Today, 12).Year);
        } 

        private void ILA_MONTH_FR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "MONTH");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_MONTH_TO_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "MONTH");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_MONTH_FR_SelectedRowData(object pSender)
        {
            W1_MONTH_DESC_TO.EditValue = W1_MONTH_DESC_FR.EditValue;
            W1_MONTH_TO.EditValue = W1_MONTH_FR.EditValue;
        }
           
        private void ILA_ACCOUNT_LEVEL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_LEVEL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        #endregion

        #region ----- Grid Event -----



        #endregion
          
        #region ----- Adapter Lookup Event -----


        #endregion



    }
}