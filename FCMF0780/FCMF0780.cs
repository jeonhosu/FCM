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
using System.Drawing.Printing;
using System.Runtime.InteropServices;

/*
 * 
 * Project      : FLEX ERP
 * Module       : Financial(회계관리)
 * Program Name : FCMF0780
 * Description  : 원가보고서
 *
 * relevant program  : 
 * 
 * Program History :
 * 
 ------------------------------------------------------------------------------
   Date         Worker                  Description
------------------------------------------------------------------------------
 * 2016-03-10   J.LAKE                  최초 생성
 * 
 * 
 * 
 */


namespace FCMF0780
{
    public partial class FCMF0780 : Office2007Form
    {
        #region ----- Variables -----

            private ISFunction.ISConvert iString = new ISFunction.ISConvert();
            private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
            string vMULTI_LANG_FLAG = "N";

            System.Windows.Forms.PrintDialog PD = new PrintDialog();
            System.Drawing.Printing.PrinterSettings PS = new System.Drawing.Printing.PrinterSettings();

            [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool SetDefaultPrinter(string Name);

        #endregion;

        #region ----- Constructor -----

        public FCMF0780(Form pMainForm, ISAppInterface pAppInterface)
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

            if (iString.ISNull(W_PERIOD_YEAR.EditValue) == string.Empty)
            {
                //년도는 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_YEAR.Focus();
                return;
            }

            //출력구분은 필수사항입니다.
            if (iString.ISNull(W_MONTH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10550"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_MONTH_DESC.Focus();
                return;
            }

            //전년도자료조회구분은 필수사항입니다.
            if (iString.ISNull(W_FS_PRE_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10552"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_FS_PRE_YEAR_DESC.Focus();
                return;
            }

            Set_Period_Term();

            IDA_MS_YEAR_RPT.Fill();
            IGR_MS_YEAR_RPT.Focus();
        }

        //그리드 칼럼헤더에 조회되는 년의 회계기수를 표현한다.
        private void Set_Period_Term()
        {
            string vPeriod_Year;
            int vIDX_CURRENT_YEAR = IGR_MS_YEAR_RPT.GetColumnToIndex("THIS_LEFT_AMOUNT");
            int vIDX_PRIOR_YEAR = IGR_MS_YEAR_RPT.GetColumnToIndex("PRE_LEFT_AMOUNT");

            IDC_PROMPT_YEAR_COUNT_P.SetCommandParamValue("W_PERIOD_YEAR", W_PERIOD_YEAR.EditValue);
            IDC_PROMPT_YEAR_COUNT_P.ExecuteNonQuery();             

            //당기 회계기수
            vPeriod_Year = iString.ISNull(IDC_PROMPT_YEAR_COUNT_P.GetCommandParamValue("O_THIS_PROMPT"));

            IGR_MS_YEAR_RPT.GridAdvExColElement[vIDX_CURRENT_YEAR].HeaderElement[0].TL1_KR = vPeriod_Year;
            IGR_MS_YEAR_RPT.GridAdvExColElement[vIDX_CURRENT_YEAR].HeaderElement[0].Default = vPeriod_Year;
            W_THIS_PROMPT.EditValue = vPeriod_Year;
            W_THIS_LEFT.EditValue = vPeriod_Year.Substring(1, 2);

            //전기 회계기수
            vPeriod_Year = iString.ISNull(IDC_PROMPT_YEAR_COUNT_P.GetCommandParamValue("O_PRE_PROMPT"));
            IGR_MS_YEAR_RPT.GridAdvExColElement[vIDX_PRIOR_YEAR].HeaderElement[0].TL1_KR = vPeriod_Year;
            IGR_MS_YEAR_RPT.GridAdvExColElement[vIDX_PRIOR_YEAR].HeaderElement[0].Default = vPeriod_Year;
            W_PRE_PROMPT.EditValue = vPeriod_Year;
            W_PRE_LEFT.EditValue = vPeriod_Year.Substring(1, 2);

            IGR_MS_YEAR_RPT.ResetDraw = true;
            Application.DoEvents();
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
                                        object pITEM_CODE, object pITEM_NAME)
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
                            vParam[5] = W_ACCOUNT_LEVEL.EditValue;
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
                        string defaultPrint = GetDefaultPrinter();

                        System.Windows.Forms.DialogResult vResult = printDialog1.ShowDialog();
                        short vInput_Copies = printDialog1.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
                        PD.PrinterSettings = printDialog1.PrinterSettings;

                        SetDefaultPrinter(PD.PrinterSettings.PrinterName.ToString());

                        if (Convert.ToString(vResult).Equals("OK"))
                        {
                            XLPrinting1("PRINT");
                        }

                        SetDefaultPrinter(defaultPrint);
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                    {
                        ExcelExport(IGR_MS_YEAR_RPT);
                    }
                }
            }

        #endregion;

        public string GetDefaultPrinter()
        {
            PrintDocument PD = new PrintDocument();
            return PD.PrinterSettings.PrinterName;
        }

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1(string pOutput_Type)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            int vCountRowGrid = IGR_MS_YEAR_RPT.RowCount;
            //if ((itbSLIP.SelectedIndex == 0 && vCountRowGrid > 0) ||
            //    (itbSLIP.SelectedIndex == 1 && iString.ISNull(H_SLIP_HEADER_ID.EditValue) != string.Empty))
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting", vPageTotal);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0780_001.xlsx";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        int vCountRow = 0;
                        int vRow = IGR_MS_YEAR_RPT.RowIndex;

                        object vNAME = V_FS_FORM_TYPE_DESC.EditValue; // 제목
                        //object vYEAR_DATE = W_PERIOD_YEAR.EditValue; // 년도
                        //object vMONTH_DATE = W_MONTH.EditValue; // 월
                        object vTHIS_LEFT = W_THIS_LEFT.EditValue; // 당기기수
                        object vPRE_LEFT = W_PRE_LEFT.EditValue; // 전기기수
                        object vTHIS_YEAR = W_THIS_YEAR.EditValue; // 기간 마지막날
                        object vPRE_YEAR = W_PRE_YEAR.EditValue; // 전년도 기간 마지막날
                        object vORG_NAME = W_ORG_NAME.EditValue; // 법인명
                        object vTHIS_PROMPT = W_THIS_PROMPT.EditValue; // 당기기수명
                        object vPRE_PROMPT = W_PRE_PROMPT.EditValue; // 전기기수명

                        xlPrinting.HeaderWrite(vNAME, vTHIS_LEFT, vPRE_LEFT, vTHIS_YEAR, vPRE_YEAR, vORG_NAME, vTHIS_PROMPT, vPRE_PROMPT);

                        //인쇄일자 //////////////////////////////////////////////////////////////////////////////////////////////////
                        //IDC_GET_DATE.ExecuteNonQuery();
                        //object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");
                        //// 계정코드 분류에 따른 값 가져오기. Start ////////////////////////////////
                        //object vACCOUNT_CODE = W_ACCOUNT_CODE.EditValue;
                        //object vACCOUNT_DESC = W_ACCOUNT_DESC.EditValue;

                        //object vACCOUNT_CODE_R = V_CODE.EditValue;
                        //object vACCOUNT_DESC_R = V_DESC.EditValue;

                        //object vAccount_default = V_ACCOUNT_LEVEL.EditValue;



                        //if (vAccount_default.ToString() == "10")
                        //{
                        //    xlPrinting.HeaderWrite(vLOCAL_DATE, vACCOUNT_CODE, vACCOUNT_DESC);
                        //}
                        //else
                        //{
                        //    xlPrinting.HeaderWrite(vLOCAL_DATE, vACCOUNT_CODE_R, vACCOUNT_DESC_R);
                        //}
                        // 계정코드 분류에 따른 값 가져오기. End ///////////////////////////////////
                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        vCountRow = IDA_MS_YEAR_RPT.CurrentRows.Count;
                        if (vCountRow > 0)
                        {
                            vPageNumber = xlPrinting.LineWrite(IDA_MS_YEAR_RPT, vTHIS_PROMPT, vPRE_PROMPT);
                        }

                        if (pOutput_Type == "PRINT")
                        {//[PRINT]
                            ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                            xlPrinting.PreView(1, vPageNumber);

                        }
                        else if (pOutput_Type == "EXCEL")
                        {
                            ////[SAVE]
                            xlPrinting.Save("SLIP_"); //저장 파일명
                        }

                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();

                    System.Windows.Forms.Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();

                    return;
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0780_Load(object sender, EventArgs e)
        {
            IDC_GET_DEFAULT_FS_FORM_TYPE.ExecuteNonQuery();
            V_FS_FORM_TYPE_ID.EditValue = IDC_GET_DEFAULT_FS_FORM_TYPE.GetCommandParamValue("O_FS_FORM_TYPE_ID");
            V_FS_FORM_TYPE.EditValue = IDC_GET_DEFAULT_FS_FORM_TYPE.GetCommandParamValue("O_FS_FORM_TYPE");
            V_FS_FORM_TYPE_DESC.EditValue = IDC_GET_DEFAULT_FS_FORM_TYPE.GetCommandParamValue("O_FS_FORM_TYPE_DESC");

            IDC_GET_ACCOUNT_LEVEL.ExecuteNonQuery();
            W_ACCOUNT_LEVEL_NAME.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE_NAME");
            W_ACCOUNT_LEVEL.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE"); 

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
            V_LANG_DESC.BringToFront();

            W_PERIOD_YEAR.EditValue = iDate.ISYear(DateTime.Today); 

            //월별//
            IDC_DEFAULT_VALUE2_W.SetCommandParamValue("W_WHERE", "GROUP_CODE = 'MONTH' AND CODE = '01'");
            IDC_DEFAULT_VALUE2_W.ExecuteNonQuery();
            W_MONTH_DESC.EditValue = IDC_DEFAULT_VALUE2_W.GetCommandParamValue("O_CODE_NAME");
            W_MONTH.EditValue = IDC_DEFAULT_VALUE2_W.GetCommandParamValue("O_CODE");

            //전년도 조회 구분
            IDC_GET_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "FS_PRE_YEAR");
            IDC_GET_DEFAULT_VALUE.ExecuteNonQuery();
            W_FS_PRE_YEAR_DESC.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_FS_PRE_YEAR.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE");

            //기수 설정//
            Set_Period_Term();

            //년도//
            W_THIS_YEAR.EditValue = Convert.ToString(W_PERIOD_YEAR.EditValue) + '-' + Convert.ToString(W_MONTH.EditValue) + '-' + '3' + '1';
            W_PRE_YEAR.EditValue = Convert.ToString((Convert.ToInt32(W_PERIOD_YEAR.EditValue) - 1)) + '-' + Convert.ToString(W_MONTH.EditValue) + '-' + '3' + '1';

            //법인//
            IDC_ORG_P.ExecuteNonQuery();
            W_ORG_NAME.EditValue = IDC_ORG_P.GetCommandParamValue("O_ORG_NAME");
        }

        private void IGR_MS_YEAR_RPT_CellDoubleClick(object pSender)
        {
            if (IGR_MS_YEAR_RPT.RowIndex < 0)
            {
                return;
            }

            DateTime vPERIOD_DATE_FR = iDate.ISMonth_1st(string.Format("{0}-{1}", W_PERIOD_YEAR.EditValue, W_MONTH.EditValue));
            DateTime vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-{1}", W_PERIOD_YEAR.EditValue, W_MONTH.EditValue));
            if (IGR_MS_YEAR_RPT.ColIndex == 2 || IGR_MS_YEAR_RPT.ColIndex == 3)  //합계
            {
                //종료일
                vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-{1}", W_PERIOD_YEAR.EditValue, W_MONTH.EditValue));
            }
            else if (IGR_MS_YEAR_RPT.ColIndex == 4 || IGR_MS_YEAR_RPT.ColIndex == 5)
            {
                //종료일
                if (iString.ISNull(W_FS_PRE_YEAR.EditValue) == "02")
                {
                    vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-{1}", W_PERIOD_YEAR.EditValue, W_MONTH.EditValue));
                }
                else
                {
                    vPERIOD_DATE_TO = iDate.ISMonth_Last(string.Format("{0}-{1}", W_PERIOD_YEAR.EditValue, W_MONTH.EditValue));
                }
                //전년도 변경.
                vPERIOD_DATE_FR = iDate.ISDate_Month_Add(vPERIOD_DATE_FR, -12);
                vPERIOD_DATE_TO = iDate.ISDate_Month_Add(vPERIOD_DATE_TO, -12);
            }

            object vITEM_CODE = IGR_MS_YEAR_RPT.GetCellValue("HEADER_CODE");
            object vITEM_NAME = IGR_MS_YEAR_RPT.GetCellValue("ITEM_DESC");
            if (iString.ISNull(vITEM_CODE) != string.Empty)
            {
                AssmblyRun_Manual("FCMF0295", vPERIOD_DATE_FR, vPERIOD_DATE_TO, vITEM_CODE, vITEM_NAME);
            }
        }
        
        #endregion
         

        #region ----- Lookup Event -----

        private void ILA_MONTH_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "MONTH");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_FS_PRE_YEAR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FS_PRE_YEAR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_LEVEL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_LEVEL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }




        #endregion

        #region ----- Adapter Lookup Event -----


        #endregion

        private void ILA_YEAR_YYYY_SelectedRowData(object pSender)
        {
            //년도//
            W_THIS_YEAR.EditValue = Convert.ToString(W_PERIOD_YEAR.EditValue) + '-' + Convert.ToString(W_MONTH.EditValue) + '-' + '3' + '1';
            W_PRE_YEAR.EditValue = Convert.ToString((Convert.ToInt32(W_PERIOD_YEAR.EditValue) - 1)) + '-' + Convert.ToString(W_MONTH.EditValue) + '-' + '3' + '1';
        }

        private void ILA_MONTH_SelectedRowData(object pSender)
        {
            //년도//
            W_THIS_YEAR.EditValue = Convert.ToString(W_PERIOD_YEAR.EditValue) + '-' + Convert.ToString(W_MONTH.EditValue) + '-' + '3' + '1';
            W_PRE_YEAR.EditValue = Convert.ToString((Convert.ToInt32(W_PERIOD_YEAR.EditValue) - 1)) + '-' + Convert.ToString(W_MONTH.EditValue) + '-' + '3' + '1';
        }
    }
}