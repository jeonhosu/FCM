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

namespace FCMF0505
{
    public partial class FCMF0505 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCONFIRM_CHECK = "N";

        #endregion;


        #region ----- Constructor -----

            public FCMF0505(Form pMainForm, ISAppInterface pAppInterface)
            {
                InitializeComponent();
                this.MdiParent = pMainForm;
                isAppInterfaceAdv1.AppInterface = pAppInterface;
            }

        #endregion;


        #region ----- Private Methods ----
            
        //조건에 부합되는 자료를 조회한다.
        private void Search()
        {
            //회계일자는 계정기준(1번째 탭)과 거래처기준(2번째 탭) 모두 필수이다.
            if (iString.ISNull(V_GL_DATE_FR.EditValue) == string.Empty)
            {
                //시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_FR.Focus();
                return;
            }

            if (iString.ISNull(V_GL_DATE_TO.EditValue) == string.Empty)
            {
                //종료일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_TO.Focus();
                return;
            }

            if (Convert.ToDateTime(V_GL_DATE_FR.EditValue) > Convert.ToDateTime(V_GL_DATE_TO.EditValue))
            {
                //종료일은 시작일 이후이어야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10345"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE_FR.Focus();
                return;
            }

            //조회조건의 필요한 값이 모두 채워졌을 경우 조건에 부합되는 자료를 조회한다.
            if (TB_MAIN.SelectedTab.TabIndex == TP_LEDGER_ALL.TabIndex) {
                IDA_ACCOUNT_LEDGER_ALL.Fill();
                IGR_ACCOUNT_LEDGER_ALL.Focus();
            }
            else{
                IDA_ACCOUNT_LIST.Fill();
                IGR_LIST_LEDGER_ACCOUNT.Focus();
            }                

        }


        //계정구분 조건에 값이 있는지 체크
        private void Check_ACCOUNT_LEVEL()
        {
            if (iString.ISNull(V_ACCOUNT_LEVEL.EditValue) == string.Empty)
            {
                //계정구분은 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10523"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_ACCOUNT_LEVEL_DESC.Focus();
                return;
            }
        }

        //조회된 자료에서 더블클릭하면 전표팝업 띄워준다.
        private void Show_Slip_Detail(decimal pSLIP_HEADER_ID)
        {
            if (pSLIP_HEADER_ID != Convert.ToDecimal(0))
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                Application.DoEvents();

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, pSLIP_HEADER_ID);
                vFCMF0204.Show(); 
            }
        }

        #endregion;


        #region ----- Convert decimal  Method ----
            
            //아래 문은 범위조건인 계정과목을 비교하기 위해 사용하는 것이나 굳이 필요치는 않다.
            //향후 사용할 지 몰라 지우지 않는다.
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

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "CSV File(*.csv)|*.csv|Excel file(*.xlsx)|*.xlsx|Excel file(*.xls)|*.xls";
            saveFileDialog.DefaultExt = ".csv";

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
                //string vFileName = saveFileDialog1.FileName;
                //converter.GridToExcel(pGrid.BaseGrid, vFileName, ConverterOptions.ColumnHeaders);
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

        #region ----- Excel Export -----

        private void Xls_Export(ISDataAdapter pAdapter, ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Account List";

            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.Filter = "CSV File(*.csv)|*.csv|Excel file(*.xlsx)|*.xlsx|Excel file(*.xls)|*.xls";
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

            System.Windows.Forms.Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //DATA 조회   
            int vCountRow = pAdapter.CurrentRows.Count;

            if (vCountRow < 1)
            {
                vMessageText = isMessageAdapter1.ReturnText("EAPP_10106");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            try
            {
                //Step 1 : Instantiate the spreadsheet creation engine.
                ExcelEngine ExcelEngine = new ExcelEngine();

                //Step 2 : Instantiate the excel application object.
                IApplication Exc_App = ExcelEngine.Excel;

                //set 2.1 : file Extension check =>xlsx, xls 
                if(Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
                }
                else
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;
                }

                //A new workbook is created.[Equivalent to creating a new workbook in MS Excel]
                //The new workbook will have 3 worksheets
                IWorkbook Exc_WorkBook = Exc_App.Workbooks.Create(1);
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel97to2003; 
                }
                else
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel2007;
                }

                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet sheet = Exc_WorkBook.Worksheets[0];
                 
                //Export DataTable.
                sheet.ImportDataTable(pAdapter.OraDataTable(), false, 1, 1, pAdapter.CurrentRows.Count, pAdapter.OraSelectData.Columns.Count);

                //1.title insert
                int vHeaderCount = pGrid.GridAdvExColElement[0].HeaderElement.Count;
                for (int h = 1; h <= vHeaderCount; h++)
                {
                    sheet.InsertRow(h);
                    object vTitle = string.Empty;
                    for (int c = 0; c < pGrid.ColCount; c++)
                    {
                        if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL1_KR)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount- h].TL1_KR;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL2_CN)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL2_CN;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL3_VN)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL3_VN;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL4_JP)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL4_JP;
                        }
                        else
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].Default;
                        }

                        sheet.Range[1, c + 1].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                        sheet.Range[1, c + 1].Value = iString.ISNull(vTitle);
                        sheet.AutofitColumn(c + 1);
                        if (iString.ISNull(pGrid.GridAdvExColElement[c].Visible) == "0")
                        {
                            sheet.HideColumn(c + 1);
                        }
                    }
                }

                ////2.prompt insert
                //sheet.InsertRow(2);
                //sheet.ImportDataTable(IDA_REJECT_DETAIL_TITLE.OraDataTable(), false, 2, 1); 
                //Exc_WorkBook.ActiveSheet.AutofitColumn(1);

                //Saving the workbook to disk.
                Exc_WorkBook.SaveAs(vSaveFileName);

                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(vSaveFileName);
                }

            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;


        #region ----- XL Print 1 Method -----

        //private void XLPrinting_Main(string pOutput_Type)
        //{
        //    object vSlip_Header_id;
        //    object vGL_Date;
        //    object vGL_Num;

        //    Application.UseWaitCursor = true;
        //    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    if (itbSLIP.SelectedTab.TabIndex == 2)
        //    {
        //        vSlip_Header_id = H_SLIP_HEADER_ID.EditValue;
        //        vGL_Date = GL_DATE.EditValue;
        //        vGL_Num = GL_NUM.EditValue;
        //    }
        //    else
        //    {
        //        vSlip_Header_id = igrSLIP_LIST.GetCellValue("SLIP_HEADER_ID");
        //        vGL_Date = igrSLIP_LIST.GetCellValue("GL_DATE");
        //        vGL_Num = igrSLIP_LIST.GetCellValue("GL_NUM");
        //    }

        //    AssmblyRun_Manual("FCMF0211", vSlip_Header_id, vGL_Date, vGL_Num);
        //    //IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", SLIP_DATE.EditValue);
        //    //IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0202");
        //    //IDC_GET_REPORT_SET_P.ExecuteNonQuery();
        //    //string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
        //    //if (vREPORT_TYPE.ToUpper() == "BSK")
        //    //{
        //    //    XLPrinting_BSK(pOutput_Type);
        //    //}
        //    //else
        //    //{
        //    //    XLPrinting(pOutput_Type);
        //    //}

        //    Application.UseWaitCursor = false;
        //    System.Windows.Forms.Cursor.Current = Cursors.Default;
        //    Application.DoEvents();

        //}

        //엑셀로 출력 시 사용하는 것인데, 현 프로그램에서는 사용안한다.

        private void XLPrinting(string pOutChoice)
        {
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                //계정별 인쇄
                XLPrinting_1(pOutChoice);
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                //전체 인쇄
                XLPrinting_2(pOutChoice);
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_1(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = IGR_ACCOUNT_LEDGER_DETAIL.RowCount;
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
                vSaveFileName = "Account_ledger";

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0505_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_FR.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    object vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    object vPeriod = vDATE_FORMAT;

                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_TO.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    vPeriod = String.Format("{0} ~ {1}", vPeriod, vDATE_FORMAT);

                    object vACCOUNT_CODE = IGR_LIST_LEDGER_ACCOUNT.GetCellValue("ACCOUNT_CODE");
                    object vACCOUNT_DESC = IGR_LIST_LEDGER_ACCOUNT.GetCellValue("ACCOUNT_DESC");
                    vACCOUNT_DESC = string.Format("({0}){1}", vACCOUNT_CODE, vACCOUNT_DESC);

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_1(vPeriod, vACCOUNT_DESC);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(IGR_ACCOUNT_LEDGER_DETAIL, vPeriod);

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_2(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = IGR_ACCOUNT_LEDGER_ALL.RowCount;
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
                vSaveFileName = "Account_ledger_All";

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0505_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_FR.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    object vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    object vPeriod = vDATE_FORMAT;

                    IDC_DATE_FORMAT.SetCommandParamValue("P_DATE", V_GL_DATE_TO.EditValue);
                    IDC_DATE_FORMAT.ExecuteNonQuery();
                    vDATE_FORMAT = IDC_DATE_FORMAT.GetCommandParamValue("O_DATE");
                    vPeriod = String.Format("{0} ~ {1}", vPeriod, vDATE_FORMAT);


                    object vACCOUNT_CODE = IGR_ACCOUNT_LEDGER_ALL.GetCellValue("ACCOUNT_GROUP_CODE");
                    object vACCOUNT_DESC = IGR_ACCOUNT_LEDGER_ALL.GetCellValue("ACCOUNT_GROUP_DESC");
                    vACCOUNT_DESC = string.Format("({0}){1}", vACCOUNT_CODE, vACCOUNT_DESC);

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_2(vPeriod, vACCOUNT_DESC);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_2(IGR_ACCOUNT_LEDGER_ALL, vPeriod, vACCOUNT_DESC);

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
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
                        XLPrinting("PRINT");
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                    {
                        if (TB_MAIN.SelectedTab.TabIndex == TP_LEDGER_ALL.TabIndex)
                        {
                            //ExcelExport(IGR_ACCOUNT_LEDGER_ALL);
                            Xls_Export(IDA_ACCOUNT_LEDGER_ALL, IGR_ACCOUNT_LEDGER_ALL);
                        }
                        else if (TB_MAIN.SelectedTab.TabIndex == TP_LEDGER_ACCOUNT.TabIndex)
                        {
                            //ExcelExport(IGR_ACCOUNT_LEDGER_DETAIL);
                            Xls_Export(IDA_ACCOUNT_LEDGER_DETAIL, IGR_ACCOUNT_LEDGER_DETAIL);
                        }
                    }
                }
            }

        #endregion;


        #region ----- Form Event -----

        private void FCMF0505_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iString.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;

            int vIDX_ACC_CONFIRM_FLAG = IGR_ACCOUNT_LEDGER_ALL.GetColumnToIndex("CONFIRM_FLAG");
            int vIDX_CUST_CONFIRM_FLAG = IGR_ACCOUNT_LEDGER_DETAIL.GetColumnToIndex("CONFIRM_FLAG");
            if (mCONFIRM_CHECK == "Y")
            {
                GB_CONFIRM_STATUS.Visible = true;

                IGR_ACCOUNT_LEDGER_ALL.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 1;
                IGR_ACCOUNT_LEDGER_DETAIL.GridAdvExColElement[vIDX_CUST_CONFIRM_FLAG].Visible = 1;
            }
            else
            {
                GB_CONFIRM_STATUS.Visible = false;

                IGR_ACCOUNT_LEDGER_ALL.GridAdvExColElement[vIDX_ACC_CONFIRM_FLAG].Visible = 0;
                IGR_ACCOUNT_LEDGER_DETAIL.GridAdvExColElement[vIDX_CUST_CONFIRM_FLAG].Visible = 0;
            }

            IDC_GET_OPERATION_DIV_FLAG_P.ExecuteNonQuery();
            string vOPERATION_DIV_FLAG = iString.ISNull(IDC_GET_OPERATION_DIV_FLAG_P.GetCommandParamValue("O_OPERATION_DIV_FLAG"));
            if (vOPERATION_DIV_FLAG == "Y")
            {
                V_OPERATION_DIV_NAME.Visible = true;
            }
            else
            {
                V_OPERATION_DIV_NAME.Visible = false;
            }

            IGR_ACCOUNT_LEDGER_ALL.ResetDraw = true;
            IGR_ACCOUNT_LEDGER_DETAIL.ResetDraw = true;
        }

        private void FCMF0505_Shown(object sender, EventArgs e)
        {
            V_GL_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            V_GL_DATE_TO.EditValue = System.DateTime.Today;

            IDC_GET_ACCOUNT_LEVEL.ExecuteNonQuery();
            V_ACCOUNT_LEVEL_DESC.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE_NAME");
            V_ACCOUNT_LEVEL.EditValue = IDC_GET_ACCOUNT_LEVEL.GetCommandParamValue("O_CODE");
        }

        private void V_ACCOUNT_CODE_FR_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(V_ACCOUNT_CODE_FR.EditValue) == string.Empty)
            {
                V_ACCOUNT_CODE_TO.EditValue = null;
                V_ACCOUNT_DESC_TO.EditValue = null;
            }
        }

        private void V_RB_CONFIRM_ALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;

            V_CONFIRM_STATUS.EditValue = iStatus.RadioCheckedString;
        }

        #endregion

        #region ----- Grid Event -----

        //실 자료가 있는 행을 더블클릭하면 전표팝업을 띄워준다.
        private void IGR_ACCOUNT_LEDGER_DETAIL_CellDoubleClick(object pSender)
        {
            //if (igrLIST_GENERAL_LEDGER_ACCOUNT.RowIndex > -1)
            if (IGR_ACCOUNT_LEDGER_DETAIL.Row > 0)
            {
                decimal vSLIP_HEADER_ID = iString.ISDecimaltoZero(IGR_ACCOUNT_LEDGER_DETAIL.GetCellValue("SLIP_HEADER_ID"));

                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        //실 자료가 있는 행을 더블클릭하면 전표팝업을 띄워준다.
        private void IGR_ACCOUNT_LEDGER_ALL_CellDoubleClick(object pSender)
        {
            if (IGR_ACCOUNT_LEDGER_ALL.Row > 0)
            {
                decimal vSLIP_HEADER_ID = iString.ISDecimaltoZero(IGR_ACCOUNT_LEDGER_ALL.GetCellValue("SLIP_HEADER_ID"));

                Show_Slip_Detail(vSLIP_HEADER_ID);
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_ACCOUNT_CODE_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Check_ACCOUNT_LEVEL();  //계정구분이 설정되어 있는지를 체크한다.

            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_LEVEL", V_ACCOUNT_LEVEL.EditValue);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_CODE", null);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ENABLED_YN", "N");
        }

        private void ILA_ACCOUNT_CODE_FR_SelectedRowData(object pSender)
        {
            V_ACCOUNT_CODE_TO.EditValue = V_ACCOUNT_CODE_FR.EditValue;
            V_ACCOUNT_DESC_TO.EditValue = V_ACCOUNT_DESC_FR.EditValue;
        }

        private void ILA_ACCOUNT_CODE_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Check_ACCOUNT_LEVEL();  //계정구분이 설정되어 있는지를 체크한다.

            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_LEVEL", V_ACCOUNT_LEVEL.EditValue);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ACCOUNT_CODE", V_ACCOUNT_CODE_FR.EditValue);
            ILD_ACCOUNT_CODE.SetLookupParamValue("P_ENABLED_YN", "N");
        }

        private void ILA_ACCOUNT_LEVEL_SelectedRowData(object pSender)
        {
            V_ACCOUNT_CODE_FR.EditValue = String.Empty;
            V_ACCOUNT_DESC_FR.EditValue = String.Empty;

            V_ACCOUNT_CODE_TO.EditValue = String.Empty;
            V_ACCOUNT_DESC_TO.EditValue = String.Empty;
        } 

        private void ILA_ACCOUNT_LEVEL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_LEVEL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_OPERATION_DIVISION_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OPERATION_DIVISION");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion


        #region ----- Adapter Lookup Event -----

        #endregion


    }
}