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

namespace FCMF0601
{
    public partial class FCMF0601 : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0601()
        {
            InitializeComponent();
        }

        public FCMF0601(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iString.ISNull(BUDGET_YEAR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_YEAR_0.Focus();
                return;
            }

            string vACCOUNT_CODE = iString.ISNull(igrBUDGET_ACCOUNT.GetCellValue("ACCOUNT_CODE"));
            int vIDX_ACCOUNT_CODE = igrBUDGET_ACCOUNT.GetColumnToIndex("ACCOUNT_CODE");

            idaBUDGET.SetSelectParamValue("P_BUDGET_CONTROL_YN", "N");
            idaBUDGET.SetSelectParamValue("P_CHECK_CAPACITY", "Y");

            idaBUDGET_ACCOUNT.SetSelectParamValue("P_CHECK_CAPACITY", "C");
            idaBUDGET_ACCOUNT.SetSelectParamValue("P_ENABLED_YN", "Y");
            idaBUDGET_ACCOUNT.Fill();

            if (iString.ISNull(vACCOUNT_CODE) != string.Empty)
            {
                for (int i = 0; i < igrBUDGET_ACCOUNT.RowCount; i++)
                {
                    if (vACCOUNT_CODE == iString.ISNull(igrBUDGET_ACCOUNT.GetCellValue(i, vIDX_ACCOUNT_CODE)))
                    {
                        igrBUDGET_ACCOUNT.CurrentCellMoveTo(i, vIDX_ACCOUNT_CODE);
                        igrBUDGET_ACCOUNT.CurrentCellActivate(i, vIDX_ACCOUNT_CODE);
                        return;
                    }
                }
            }
            //Set_Total_Amount();
        }

        //private void Set_Total_Amount()
        //{
        //    decimal vTotal_Base_Amount = 0;
        //    decimal vTotal_Add_Amount = 0;
        //    decimal vTotal_Move_Amount = 0;
        //    decimal vTotal_Next_Amount = 0;
        //    object vAmount;
        //    int vIDX_Base_Col = igrBUDGET.GetColumnToIndex("BASE_AMOUNT");
        //    int vIDX_Add_Col = igrBUDGET.GetColumnToIndex("ADD_AMOUNT");
        //    int vIDX_Move_Col = igrBUDGET.GetColumnToIndex("MOVE_AMOUNT");
        //    int vIDX_Next_Col = igrBUDGET.GetColumnToIndex("NEXT_AMOUNT");

        //    for (int r = 0; r < idaBUDGET.SelectRows.Count; r++)
        //    {
        //        vAmount = 0;
        //        vAmount = igrBUDGET.GetCellValue(r, vIDX_Base_Col);
        //        vTotal_Base_Amount = vTotal_Base_Amount + iString.ISDecimaltoZero(vAmount);

        //        vAmount = 0;
        //        vAmount = igrBUDGET.GetCellValue(r, vIDX_Add_Col);
        //        vTotal_Add_Amount = vTotal_Add_Amount + iString.ISDecimaltoZero(vAmount);

        //        vAmount = 0;
        //        vAmount = igrBUDGET.GetCellValue(r, vIDX_Move_Col);
        //        vTotal_Move_Amount = vTotal_Move_Amount + iString.ISDecimaltoZero(vAmount);

        //        vAmount = 0;
        //        vAmount = igrBUDGET.GetCellValue(r, vIDX_Next_Col);
        //        vTotal_Next_Amount = vTotal_Next_Amount + iString.ISDecimaltoZero(vAmount);
        //    }
        //    TOTAL_BASE_AMOUNT.EditValue = vTotal_Base_Amount;
        //    TOTAL_ADD_AMOUNT.EditValue = vTotal_Add_Amount;
        //    TOTAL_MOVE_AMOUNT.EditValue = vTotal_Move_Amount;
        //    TOTAL_NEXT_AMOUNT.EditValue = vTotal_Next_Amount;
        //}

        private void Set_CheckBox(int pIDX_Col)
        {// 그리드 체크박스 전체 선택/ 선택 취소 기능.
            string mCheckBox_Value;
            for (int r = 0; r < igrBUDGET.RowCount; r++)
            {
                mCheckBox_Value = iString.ISNull(igrBUDGET.GetCellValue(r, pIDX_Col), "N");
                if (mCheckBox_Value == "Y".ToString())
                {
                    igrBUDGET.SetCellValue(r, pIDX_Col, "N");
                }
                else
                {
                    igrBUDGET.SetCellValue(r, pIDX_Col, "Y");
                }
            }
        }

        private void Show_Detail(object pPERIOD_NAME, object pACCOUNT_DESC, object pACCOUNT_CODE, object pACCOUNT_CONTROL_ID
                                , object pBUDGET_DEPT_NAME, object pBUDGET_DEPT_ID)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            FCMF0601_DETAIL vFCMF0601_DETAIL = new FCMF0601_DETAIL(isAppInterfaceAdv1.AppInterface, pPERIOD_NAME
                                                                , pACCOUNT_DESC, pACCOUNT_CODE, pACCOUNT_CONTROL_ID
                                                                , pBUDGET_DEPT_NAME, pBUDGET_DEPT_ID);

            dlgRESULT = vFCMF0601_DETAIL.ShowDialog();
            vFCMF0601_DETAIL.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print Method -----

        private void XLPrinting(string pOutChoice)
        {
            object vPRINT_TYPE = string.Empty;
            object vDept_Code_Fr = string.Empty;
            object vDept_Code_To = string.Empty;
            object vAccount_Code_Fr = string.Empty;
            object vAccount_Code_To = string.Empty;

            DialogResult dlgResult;
            FCMF0601_PRINT vFCMF0601_PRINT = new FCMF0601_PRINT(isAppInterfaceAdv1.AppInterface, 
                                                                DEPT_NAME_FR_0.EditValue, DEPT_CODE_FR_0.EditValue, DEPT_ID_FR_0.EditValue,
                                                                DEPT_NAME_TO_0.EditValue, DEPT_CODE_TO_0.EditValue, DEPT_ID_TO_0.EditValue);
            dlgResult = vFCMF0601_PRINT.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                vPRINT_TYPE = vFCMF0601_PRINT.Get_Print_Type;
                vDept_Code_Fr = vFCMF0601_PRINT.Get_Dept_Code_Fr;
                vDept_Code_To = vFCMF0601_PRINT.Get_Dept_Code_To;
                vAccount_Code_Fr = vFCMF0601_PRINT.Get_Account_Code_Fr;
                vAccount_Code_To = vFCMF0601_PRINT.Get_Account_Code_To;
                if (iString.ISNull(vPRINT_TYPE) == "D")
                {
                    //부서별
                    XLPrinting_1(pOutChoice, vDept_Code_Fr, vDept_Code_To, vAccount_Code_Fr, vAccount_Code_To);
                }
                else if (iString.ISNull(vPRINT_TYPE) == "A")
                {
                    //계정별
                    XLPrinting_2(pOutChoice, vDept_Code_Fr, vDept_Code_To, vAccount_Code_Fr, vAccount_Code_To);
                }
            }
            vFCMF0601_PRINT.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_1(string pOutChoice, object pDept_Code_Fr, object pDept_Code_To,
                                    object pAccount_Code_Fr, object pAccount_Code_To)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            IDA_PRINT_BUDGET_DEPT.SetSelectParamValue("W_DEPT_CODE_FR", pDept_Code_Fr);
            IDA_PRINT_BUDGET_DEPT.SetSelectParamValue("W_DEPT_CODE_TO", pDept_Code_To);
            IDA_PRINT_BUDGET_DEPT.SetSelectParamValue("W_ACCOUNT_CODE_FR", pAccount_Code_Fr);
            IDA_PRINT_BUDGET_DEPT.SetSelectParamValue("W_ACCOUNT_CODE_TO", pAccount_Code_To);
            IDA_PRINT_BUDGET_DEPT.Fill();
            int vCountRow = IDA_PRINT_BUDGET_DEPT.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Budget_assign_depart";

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
                vMessageText = string.Format(" Writing Starting...");
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0601_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    object vBUDGET_YEAR = BUDGET_YEAR_0.EditValue;

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_1(vBUDGET_YEAR);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(IDA_PRINT_BUDGET_DEPT);

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

        private void XLPrinting_2(string pOutChoice, object pDept_Code_Fr, object pDept_Code_To,
                                    object pAccount_Code_Fr, object pAccount_Code_To)
        {
            //예산신청내역 - 계정별
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            IDA_PRINT_BUDGET_ACCOUNT.SetSelectParamValue("W_DEPT_CODE_FR", pDept_Code_Fr);
            IDA_PRINT_BUDGET_ACCOUNT.SetSelectParamValue("W_DEPT_CODE_TO", pDept_Code_To);
            IDA_PRINT_BUDGET_ACCOUNT.SetSelectParamValue("W_ACCOUNT_CODE_FR", pAccount_Code_Fr);
            IDA_PRINT_BUDGET_ACCOUNT.SetSelectParamValue("W_ACCOUNT_CODE_TO", pAccount_Code_To);
            IDA_PRINT_BUDGET_ACCOUNT.Fill();
            int vCountRow = IDA_PRINT_BUDGET_ACCOUNT.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Budget_assign_account";

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
                vMessageText = string.Format(" Writing Starting...");
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0601_002.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    object vBUDGET_YEAR = BUDGET_YEAR_0.EditValue;

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_2(vBUDGET_YEAR);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_2(IDA_PRINT_BUDGET_ACCOUNT);

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

        #endregion;

        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID
                                    , object pPERIOD_NAME, object pACCOUNT_DESC, object pACCOUNT_CODE, object pACCOUNT_CONTROL_ID
                                    , object pBUDGET_DEPT_NAME, object pBUDGET_DEPT_ID)
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
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive == "N")
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

                            object[] vParam = new object[10];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = string.Format("{0:yyyy-MM-dd}", iDate.ISMonth_1st(pPERIOD_NAME));    //전표일자 시작
                            vParam[3] = string.Format("{0:yyyy-MM-dd}", iDate.ISMonth_Last(pPERIOD_NAME));   //전표일자 종료
                            vParam[4] = DBNull.Value;                   //기표번호
                            vParam[5] = pACCOUNT_CONTROL_ID;            //계정id
                            vParam[6] = pACCOUNT_CODE;              //계정
                            vParam[7] = pACCOUNT_DESC;              //계정
                            vParam[8] = pBUDGET_DEPT_NAME;          //예산부서
                            vParam[9] = pBUDGET_DEPT_ID;            //예산부서

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
                    if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive == "N")
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
                    if (idaBUDGET.IsFocused)
                    {
                        idaBUDGET.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBUDGET.IsFocused)
                    {
                        idaBUDGET.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting("FILE");
                }
            }
        }

        #endregion;

        #region ----- Forms Event ------
        
        private void FCMF0601_Load(object sender, EventArgs e)
        {
            idaBUDGET_ACCOUNT.FillSchema();
        }

        private void FCMF0601_Shown(object sender, EventArgs e)
        {
            BUDGET_YEAR_0.EditValue = DateTime.Today.Year;

            IDC_DEFAULT_DEPT.ExecuteNonQuery();
            DEPT_ID_FR_0.EditValue = IDC_DEFAULT_DEPT.GetCommandParamValue("O_DEPT_ID");
            DEPT_CODE_FR_0.EditValue = IDC_DEFAULT_DEPT.GetCommandParamValue("O_DEPT_CODE");
            DEPT_NAME_FR_0.EditValue = IDC_DEFAULT_DEPT.GetCommandParamValue("O_DEPT_NAME");

            DEPT_ID_TO_0.EditValue = DEPT_ID_FR_0.EditValue;
            DEPT_CODE_TO_0.EditValue = DEPT_CODE_FR_0.EditValue;
            DEPT_NAME_TO_0.EditValue = DEPT_NAME_FR_0.EditValue;
        }

        private void igrBUDGET_CellDoubleClick(object pSender)
        {
            if (igrBUDGET.Row > 0)
            {
                if (iDate.ISDate(string.Format("{0}-01", igrBUDGET.GetCellValue("BUDGET_PERIOD"))) == false)
                {
                    return;
                }
                
                AssmblyRun_Manual("FCMF0206"
                                , igrBUDGET.GetCellValue("BUDGET_PERIOD")
                                , igrBUDGET.GetCellValue("ACCOUNT_DESC"), igrBUDGET.GetCellValue("ACCOUNT_CODE"), igrBUDGET.GetCellValue("ACCOUNT_CONTROL_ID")
                                , igrBUDGET.GetCellValue("DEPT_NAME"), igrBUDGET.GetCellValue("DEPT_ID")); 
            } 
        }

        //private void ibtnEXE_NEXT_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    string mMESSAGE;
        //    if (iString.ISNull(PERIOD_NAME_0.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        PERIOD_NAME_0.Focus();
        //        return;
        //    }
        //    idcEXE_BUDGET_NEXT_PERIOD.ExecuteNonQuery();
        //    mMESSAGE = iString.ISNull(idcEXE_BUDGET_NEXT_PERIOD.GetCommandParamValue("O_MESSAGE"));
        //    if (mMESSAGE != string.Empty)
        //    {
        //        MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }           
        //}

        //private void ibtnEXE_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    string mMESSAGE;
        //    if (iString.ISNull(PERIOD_NAME_0.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        PERIOD_NAME_0.Focus();
        //        return;
        //    }
        //    idcBUDGET_CLOSE.ExecuteNonQuery();
        //    mMESSAGE = iString.ISNull(idcBUDGET_CLOSE.GetCommandParamValue("O_MESSAGE"));
        //    if (mMESSAGE != string.Empty)
        //    {
        //        MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //}

        #endregion

        #region ----- Lookup Event ------

        private void ilaACCOUNT_CONTROL_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", ACCOUNT_CODE_FR_0.EditValue);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDEPT_FR_0_SelectedRowData(object pSender)
        {
            DEPT_NAME_TO_0.EditValue = DEPT_NAME_FR_0.EditValue;
            DEPT_CODE_TO_0.EditValue = DEPT_CODE_FR_0.EditValue;
            DEPT_ID_TO_0.EditValue = DEPT_ID_FR_0.EditValue;
        }

        private void ilaACCOUNT_CONTROL_FR_0_SelectedRowData(object pSender)
        {
            ACCOUNT_DESC_TO_0.EditValue = ACCOUNT_DESC_FR_0.EditValue;
            ACCOUNT_CODE_TO_0.EditValue = ACCOUNT_CODE_FR_0.EditValue;
            ACCOUNT_CONTROL_ID_TO_0.EditValue = ACCOUNT_CONTROL_ID_FR_0.EditValue;
        }

        private void ilaDEPT_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(string.Format("{0}-01", BUDGET_YEAR_0.EditValue)));
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(string.Format("{0}-12", BUDGET_YEAR_0.EditValue)));
        }

        private void ilaDEPT_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", DEPT_CODE_FR_0.EditValue);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(string.Format("{0}-01", BUDGET_YEAR_0.EditValue)));
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(string.Format("{0}-12", BUDGET_YEAR_0.EditValue)));
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBUDGET_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_PERIOD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Budget Period(예산년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Department(부서)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Account Code(계정)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaBUDGET_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}