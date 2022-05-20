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
using Syncfusion.XlsIO;

namespace FCMF0209
{
    public partial class FCMF0209 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        string mCONFIRM_CHECK = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0209()
        {
            InitializeComponent();
        }

        public FCMF0209(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            //int vCountRow = ((ISGridAdvEx)(pObject)).RowCount;
            //((mdiMMPS52)(this.MdiParent)).StatusSTRIP_Form_Open_iF_Value.Text = "0";
            //(()(this.MdiParent)).

            //System.Type vType = this.MdiParent.GetType();
            //object vO1 = Convert.ChangeType(pMainForm, System.Type.GetType(vType.FullName));
            string vPathReport = string.Empty;
            object vObject = this.MdiParent.Tag;
            if (vObject != null)
            {
                bool isConvert = vObject is string;
                if (isConvert == true)
                {
                    vPathReport = vObject as string;
                }
            }
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
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void DefaultValue()
        {
            DateTime vGetDate = GetDate();
                        
            GL_DATE_FR_0.EditValue = iDate.ISMonth_1st(vGetDate);
            GL_DATE_TO_0.EditValue = vGetDate;
        }

        private void SEARCH_DB()
        {
            if (iConv.ISNull(GL_DATE_FR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                GL_DATE_FR_0.Focus();
                return;
            }

            if (iConv.ISNull(GL_DATE_TO_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                GL_DATE_TO_0.Focus();
                return;
            }

            if (Convert.ToDateTime(GL_DATE_FR_0.EditValue) > Convert.ToDateTime(GL_DATE_TO_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                GL_DATE_FR_0.Focus();
                return;
            }

            //데이터 그리드 초기화.
            idaSLIP_MANAGEMENT.SetSelectParamValue("P_SOB_ID", -1);
            idaSLIP_MANAGEMENT.Fill();

            SET_GRID_COL_VISIBLE();  // 그리드 보이기/감추기 설정.

            Application.DoEvents();
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            

            object mMANAGEMENT_DESC = Get_Management_Desc();
            idaSLIP_MANAGEMENT.SetSelectParamValue("P_MANAGEMENT_DESC", mMANAGEMENT_DESC);
            idaSLIP_MANAGEMENT.SetSelectParamValue("P_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
            idaSLIP_MANAGEMENT.Fill();
            igrSLIP_MANAGEMENT.Focus();
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
        }

        private object Get_Management_Desc()
        {
            object mMANAGEMENT_DESC = null;
            if (iConv.ISNull(MANAGEMENT_DESC_0.EditValue) == String.Empty)
            {
                mMANAGEMENT_DESC = null;
            }
            else if (iConv.ISNull(DATA_TYPE_0.EditValue) == "DATE")
            {
                mMANAGEMENT_DESC = MANAGEMENT_DESC_0.DateTimeValue.ToShortDateString().ToString();
            }
            else
            {
                mMANAGEMENT_DESC = MANAGEMENT_DESC_0.EditValue;
            }
            return mMANAGEMENT_DESC;
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            idaMANAGEMENT_PROMPT.Fill();
            if (idaMANAGEMENT_PROMPT.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            // Adapter Column.
            int mIDX_Column;            // 시작 COLUMN.
            int mMax_Column = idaMANAGEMENT_PROMPT.SelectColumns.Count - 2; // 종료 COLUMN.
            object mCOLUMN_DESC;        // 헤더 프롬프트.
            
            //Grid Column.
            int mGrid_Column = 14;     // 그리드 시작 Column.
            for (mIDX_Column = 1; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = idaMANAGEMENT_PROMPT.CurrentRow[mIDX_Column];
                if (iConv.ISNull(mCOLUMN_DESC, ":=") == ":=".ToString())
                {
                    igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].Visible = 0;
                }
                else
                {
                    igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].Visible = 1;
                    igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].HeaderElement[0].Default = iConv.ISNull(mCOLUMN_DESC);
                    igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].HeaderElement[0].TL1_KR = iConv.ISNull(mCOLUMN_DESC);
                }
                mGrid_Column = mGrid_Column + 1;
            }
            igrSLIP_MANAGEMENT.ResetDraw = true;
        }

        private void INIT_EDIT_TYPE()
        {
            MANAGEMENT_DESC_0.EditValue = null;
            MANAGEMENT_DESC_0.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
            MANAGEMENT_DESC_0.NumberDecimalDigits = 0;
            if (iConv.ISNull(DATA_TYPE_0.EditValue) == "NUMBER".ToString())
            {
                MANAGEMENT_DESC_0.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iConv.ISNull(DATA_TYPE_0.EditValue) == "RATE".ToString())
            {
                MANAGEMENT_DESC_0.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                MANAGEMENT_DESC_0.NumberDecimalDigits = 4;
            }
            else if (iConv.ISNull(DATA_TYPE_0.EditValue) == "DATE".ToString())
            {
                MANAGEMENT_DESC_0.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }

            if (iConv.ISNull(LOOKUP_YN_0.EditValue, "N") == "N")
            {
                MANAGEMENT_DESC_0.LookupAdapter = null;
            }
            else
            {
                MANAGEMENT_DESC_0.LookupAdapter = ilaMANAGEMENT_ITEM;
            }
            MANAGEMENT_DESC_0.Refresh();
        }

        private void SET_GRID_COL_VISIBLE()
        {
            object mMANAGEMENT_DESC = Get_Management_Desc();
            idaSLIP_MANAGEMENT_YN.SetSelectParamValue("P_MANAGEMENT_DESC", mMANAGEMENT_DESC);
            idaSLIP_MANAGEMENT_YN.Fill();

            // Adapter Column.
            int mIDX_Column;            // 시작 COLUMN.
            int mMax_Column = 0;        // 종료 COLUMN.
            int mGrid_Column = 14;      // 그리드 시작 Column.
            object mVISIBLE_YN = ":=";   // 보이기 여부.

            if (idaSLIP_MANAGEMENT_YN.OraSelectData.Rows.Count == 0)
            {
                // Adapter Column.
                mMax_Column = idaMANAGEMENT_PROMPT.SelectColumns.Count - 2; // 종료 COLUMN.
                for (mIDX_Column = 1; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mVISIBLE_YN = idaMANAGEMENT_PROMPT.CurrentRow[mIDX_Column];
                    if (iConv.ISNull(mVISIBLE_YN, ":=") == ":=".ToString())
                    {
                        igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].Visible = 0;
                    }
                    else
                    {
                        igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].Visible = 1;
                    }
                    mGrid_Column = mGrid_Column + 1;
                }
            }
            else
            {
                // Adapter Column.
                mMax_Column = idaSLIP_MANAGEMENT_YN.SelectColumns.Count - 2; // 종료 COLUMN.
                for (mIDX_Column = 1; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mVISIBLE_YN = idaSLIP_MANAGEMENT_YN.CurrentRow[mIDX_Column];
                    if (iConv.ISNull(mVISIBLE_YN, ":=") == ":=".ToString())
                    {
                        igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].Visible = 0;
                    }
                    else
                    {
                        igrSLIP_MANAGEMENT.GridAdvExColElement[mGrid_Column].Visible = 1;
                    }
                    mGrid_Column = mGrid_Column + 1;
                }
            }
            igrSLIP_MANAGEMENT.ResetDraw = true;
        }

        private void Show_Slip_Detail()
        {
            decimal mSLIP_HEADER_ID = iConv.ISDecimaltoZero(igrSLIP_MANAGEMENT.GetCellValue("SLIP_HEADER_ID"));
            if (mSLIP_HEADER_ID != Convert.ToDecimal(0))
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
                vFCMF0204.Show();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.UseWaitCursor = false;
            }
        }

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter_W(string pGroup_Code, string pWhere, string pEnabled_YN)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private object GetTerritory()
        {

            object vTerritory = "Default";
            vTerritory = isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage;
            return vTerritory;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice, object pGL_DATE_FR, object pGL_DATE_TO, ISGridAdvEx pGRID)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;
            object vTerritory = string.Empty;
                        
            int vCountRow = pGRID.RowCount;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = String.Format("Document_List_{0}~{1}", pGL_DATE_FR.ToString().Substring(0, 10), pGL_DATE_TO.ToString().Substring(0, 10));

            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.DefaultExt = "xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = saveFileDialog1.FileName;
                System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                if (vFileName.Exists)
                {
                    try
                    {
                        vFileName.Delete();
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

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            vTerritory = GetTerritory();
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.

                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0209_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    // 실제 인쇄
                    vPageNumber = xlPrinting.LineWrite(iConv.ISNull(vTerritory), pGRID);

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

                    vMessageText = "Printing End";
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

        #region ----- Excel Export II -----

        private void ExcelExport(ISDataAdapter pAdapter, ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.RestoreDirectory = true;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Slip List";     //기본 파일명. 수정필요.

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
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
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
                sheet.ImportDataTable(pAdapter.OraDataTable(), false, 1, 1, pAdapter.CurrentRows.Count, pAdapter.OraSelectData.Columns.Count, true);

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
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL1_KR;
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

                        sheet.Range[h, c + 1].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                        sheet.Range[h, c + 1].Value = iConv.ISNull(vTitle);
                        sheet.AutofitColumn(c + 1);
                        if (iConv.ISNull(pGrid.GridAdvExColElement[c].Visible) == "0")
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
                    if (idaSLIP_MANAGEMENT.IsFocused)
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaSLIP_MANAGEMENT.IsFocused)
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaSLIP_MANAGEMENT.IsFocused)
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaSLIP_MANAGEMENT.IsFocused)
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaSLIP_MANAGEMENT.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("FILE", GL_DATE_FR_0.EditValue, GL_DATE_TO_0.EditValue, igrSLIP_MANAGEMENT);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(idaSLIP_MANAGEMENT, igrSLIP_MANAGEMENT);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0209_Load(object sender, EventArgs e)
        {
            // 전표저장시 자동 승인 여부
            IDC_SLIP_CONFIRM_CHECK_P.ExecuteNonQuery();
            mCONFIRM_CHECK = iConv.ISNull(IDC_SLIP_CONFIRM_CHECK_P.GetCommandParamValue("O_CONFIRM_CHECK"));

            GB_CONFIRM_STATUS.BringToFront();
            V_RB_CONFIRM_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            if (mCONFIRM_CHECK == "Y")
            {
                GB_CONFIRM_STATUS.Visible = true;
            }
            else
            {
                GB_CONFIRM_STATUS.Visible = false;
            }

            GB_CONFIRM_STATUS.BringToFront();
        }
        
        private void FCMF0209_Shown(object sender, EventArgs e)
        {
            DefaultValue();
            INIT_MANAGEMENT_COLUMN();
        }

        private void igrSLIP_MANAGEMENT_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail(); 
        }

        private void V_RB_CONFIRM_ALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;

            V_CONFIRM_STATUS.EditValue = iStatus.RadioCheckedString;
        }

        #endregion

        #region ------ Lookup Event ------

        private void ilaGL_NUM_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildGL_NUM.SetLookupParamValue("W_GL_NUM", GL_NUM_0.EditValue);
        }

        private void ilaACCOUNT_CODE_FR_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CODE_TO_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ACCOUNT_CODE_FR", ACCOUNT_CODE_FR_0.EditValue);
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SLIP_TYPE_SLIP_DOCU_ALL.SetLookupParamValue("P_ENABLED_FLAG", "Y"); 
        }

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "N");
        }
          
        private void ilaMANAGEMENT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "MANAGEMENT_CODE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT_ITEM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildMANAGEMENT_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaMANAGEMENT_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "MANAGEMENT_CODE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ilaMANAGEMENT_TYPE_SelectedRowData(object pSender)
        {
            INIT_EDIT_TYPE();
        }
        
        private void ilaMANAGEMENT_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildMANAGEMENT_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        #endregion


    }
}