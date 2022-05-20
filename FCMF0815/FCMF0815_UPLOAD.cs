using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.IO;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil; 
using Syncfusion.XlsIO;

namespace FCMF0815
{
    public partial class FCMF0815_UPLOAD : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0815_UPLOAD()
        {
            InitializeComponent();
        }

        public FCMF0815_UPLOAD(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            
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

        #region ----- Excel Upload -----

        private void Select_Excel_File()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    V_UPLOAD_FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    V_UPLOAD_FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        // ----- Excel Import : 수급사업자 포함 ----- 
        private bool GridConvert_Upload_Consignment()
        {
            bool vImport_Status = false;

            string vSelectFullPath = iString.ISNull(V_UPLOAD_FILE_PATH.EditValue);
            string vSelectDirectoryPath = string.Empty;

            string vFileName = string.Empty;
            string vFileExtension = string.Empty;
            int vRow_CNT = 0;
            int vStart_Row = iString.ISNumtoZero(V_START_ROW.EditValue, 0);

            if (vSelectFullPath == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10026", string.Format("&&FIELD_NAME:=Upload File")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            PB_RATE.BarFillPercent = 0;
            PB_RATE.Visible = true; 

            IDA_VAT_NTS_UPLOAD_C.Cancel();
            IDA_VAT_NTS_UPLOAD_C.SetSelectParamValue("P_SOB_ID", -1);
            IDA_VAT_NTS_UPLOAD_C.Fill();
 
            //1. 사용자 선택 파일  
            vSelectDirectoryPath = Path.GetDirectoryName(vSelectFullPath);

            vFileName = Path.GetFileName(vSelectFullPath);
            vFileExtension = Path.GetExtension(vSelectFullPath).ToUpper();
        

            //--------------------------------------------------------------------------------------
            //excel 개체 생성
            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine ExcelEngine = new ExcelEngine();
            if (Path.GetExtension(vSelectFullPath).ToUpper() == ".XLSX")
            {
                ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;
            }
            else
            {
                ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
            }

            IApplication Exc_App = ExcelEngine.Excel;

            //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
            //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
            IWorkbook Exc_WorkBook = null;
            if (Path.GetExtension(vSelectFullPath).ToUpper() == ".XLSX")
            {
                Exc_WorkBook = Exc_App.Workbooks.Open(@vSelectFullPath, ExcelVersion.Excel2007);
            }
            else
            {
                Exc_WorkBook = Exc_App.Workbooks.Open(@vSelectFullPath, ExcelVersion.Excel97to2003);
            }

            try
            {
                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet Exc_Sheet = Exc_WorkBook.Worksheets[0];

                //Read data from spreadsheet.
                DataTable customersTable = Exc_Sheet.ExportDataTable(Exc_Sheet.Range, ExcelExportDataTableOptions.DefaultStyleColumnTypes);
                 
                IGR_VAT_NTS_UPLOAD_C.BeginUpdate();

                foreach (System.Data.DataRow vRow in customersTable.Rows)
                {
                    if (vRow_CNT < (vStart_Row - 1)) //index 0 부터 시작하고, 1번째 열은 프롬프트로 인식해서 2를 빼준다//
                    {
                        //
                    }
                    else
                    {
                        IDA_VAT_NTS_UPLOAD_C.AddUnder();

                        for (int vCol = 0; vCol < customersTable.Columns.Count; vCol++)
                        {
                            if (IDA_VAT_NTS_UPLOAD_C.OraSelectData.Columns[vCol].DataType.Name == "DateTime")
                            {
                                IDA_VAT_NTS_UPLOAD_C.CurrentRow[vCol] = iDate.ISGetDate(vRow[vCol]);
                            }
                            else if (IDA_VAT_NTS_UPLOAD_C.OraSelectData.Columns[vCol].DataType.Name == "Decimal")
                            {
                                IDA_VAT_NTS_UPLOAD_C.CurrentRow[vCol] = iString.ISDecimaltoZero(vRow[vCol]);
                            }
                            else
                            {
                                IDA_VAT_NTS_UPLOAD_C.CurrentRow[vCol] = iString.ISNull(vRow[vCol]);
                            }
                        } 
                    }
                    PB_RATE.BarFillPercent = (Convert.ToSingle(vRow_CNT) / Convert.ToSingle(customersTable.Rows.Count)) * 100F;
                    vRow_CNT++;
                }

                IGR_VAT_NTS_UPLOAD_C.EndUpdate();
            }
            catch (Exception Ex)
            {

                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }
            try
            {
                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }
            //CLEAR
            V_UPLOAD_FILE_PATH.EditValue = string.Empty;

            try
            {
                IDA_VAT_NTS_UPLOAD_C.Update();
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }
            PB_RATE.Visible = false;
            GB_UPLOAD_FILE.Enabled = true;
            vImport_Status = true; 

            return vImport_Status;
        }

        // ----- Excel Import : 수급사업자 미포함 ----- 
        private bool GridConvert_Upload()
        {
            bool vImport_Status = false;

            string vSelectFullPath = iString.ISNull(V_UPLOAD_FILE_PATH.EditValue);
            string vSelectDirectoryPath = string.Empty;

            string vFileName = string.Empty;
            string vFileExtension = string.Empty;
            int vRow_CNT = 0;
            int vStart_Row = iString.ISNumtoZero(V_START_ROW.EditValue, 0);

            if (vSelectFullPath == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10026", string.Format("&&FIELD_NAME:=Upload File")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }
            
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            PB_RATE.BarFillPercent = 0;
            PB_RATE.Visible = true; 

            IDA_VAT_NTS_UPLOAD.Cancel();
            IDA_VAT_NTS_UPLOAD.SetSelectParamValue("P_SOB_ID", -1);
            IDA_VAT_NTS_UPLOAD.Fill();

            //1. 사용자 선택 파일  
            vSelectDirectoryPath = Path.GetDirectoryName(vSelectFullPath);

            vFileName = Path.GetFileName(vSelectFullPath);
            vFileExtension = Path.GetExtension(vSelectFullPath).ToUpper();


            //--------------------------------------------------------------------------------------
            //excel 개체 생성
            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine ExcelEngine = new ExcelEngine();
            if (Path.GetExtension(vSelectFullPath).ToUpper() == ".XLSX")
            {
                ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;
            }
            else
            {
                ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
            }

            IApplication Exc_App = ExcelEngine.Excel;

            //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
            //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
            IWorkbook Exc_WorkBook = null;
            if (Path.GetExtension(vSelectFullPath).ToUpper() == ".XLSX")
            {
                Exc_WorkBook = Exc_App.Workbooks.Open(@vSelectFullPath, ExcelVersion.Excel2007);
            }
            else
            {
                Exc_WorkBook = Exc_App.Workbooks.Open(@vSelectFullPath, ExcelVersion.Excel97to2003);
            }

            try
            {
                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet Exc_Sheet = Exc_WorkBook.Worksheets[0];

                //Read data from spreadsheet.
                DataTable customersTable = Exc_Sheet.ExportDataTable(Exc_Sheet.Range, ExcelExportDataTableOptions.ColumnNames);

                IGR_VAT_NTS_UPLOAD.BeginUpdate();

                foreach (System.Data.DataRow vRow in customersTable.Rows)
                {
                    if (vRow_CNT < (vStart_Row - 2)) //index 0 부터 시작하고, 1번째 열은 프롬프트로 인식해서 2를 빼준다//
                    {
                        //
                    }
                    else
                    {
                        IDA_VAT_NTS_UPLOAD.AddUnder();

                        for (int vCol = 0; vCol < customersTable.Columns.Count; vCol++)
                        {
                            if (IDA_VAT_NTS_UPLOAD.OraSelectData.Columns[vCol].DataType.Name == "DateTime")
                            {
                                IDA_VAT_NTS_UPLOAD.CurrentRow[vCol] = iDate.ISGetDate(vRow[vCol]);
                            }
                            else if (IDA_VAT_NTS_UPLOAD.OraSelectData.Columns[vCol].DataType.Name == "Decimal")
                            {
                                IDA_VAT_NTS_UPLOAD.CurrentRow[vCol] = iString.ISDecimaltoZero(vRow[vCol]);
                            }
                            else
                            {
                                IDA_VAT_NTS_UPLOAD.CurrentRow[vCol] = iString.ISNull(vRow[vCol]);
                            }
                        }               
                    }
                    PB_RATE.BarFillPercent = (Convert.ToSingle(vRow_CNT) / Convert.ToSingle(customersTable.Rows.Count)) * 100F;                    
                    vRow_CNT++;
                }                

                IGR_VAT_NTS_UPLOAD.EndUpdate();
            }
            catch (Exception Ex)
            {
                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }
            try
            {
                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }
            //CLEAR
            V_UPLOAD_FILE_PATH.EditValue = string.Empty;

            try
            {
                IDA_VAT_NTS_UPLOAD.Update();
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return vImport_Status;
            }

            PB_RATE.Visible = false;
            GB_UPLOAD_FILE.Enabled = true;
            vImport_Status = true;
             
            return vImport_Status;
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
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
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

        private void FCMF0815_UPLOAD_Load(object sender, EventArgs e)
        {
            
        }

        private void FCMF0815_UPLOAD_Shown(object sender, EventArgs e)
        {
            V_RB_VAT_2.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_VAT_GUBUN.EditValue = V_RB_VAT_2.RadioCheckedString;

            V_RB_SAVE_NEW.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_SAVE_TYPE.EditValue = V_RB_SAVE_NEW.RadioCheckedString;

            V_START_ROW.EditValue = 2;

            IDA_VAT_NTS_UPLOAD_C.FillSchema();
            IDA_VAT_NTS_UPLOAD.FillSchema();
        }

        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNumtoZero(V_START_ROW.EditValue) < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10139"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_START_ROW.Focus();
                return;
            }

            GB_UPLOAD_FILE.Enabled = false;

            //수급사업자 포함 여부에 따른 메서드 호출
            if (V_CB_CONSIGNMENT_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                if (GridConvert_Upload_Consignment() == false)
                {
                    PB_RATE.Visible = false;
                    GB_UPLOAD_FILE.Enabled = true;

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();
                }
            }
            else
            {
                if (GridConvert_Upload() == false)
                {
                    PB_RATE.Visible = false;
                    GB_UPLOAD_FILE.Enabled = true;

                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();
                }
            }

            PB_RATE.Visible = false;
            GB_UPLOAD_FILE.Enabled = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }
         
        private void V_RB_VAT_2_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_VAT_2.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_VAT_GUBUN.EditValue = V_RB_VAT_2.RadioCheckedString;
            }
        }

        private void V_RB_VAT_1_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_VAT_1.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_VAT_GUBUN.EditValue = V_RB_VAT_1.RadioCheckedString;
            }
        }

        #endregion

        private void V_RB_SAVE_NEW_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_SAVE_NEW.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SAVE_TYPE.EditValue = V_RB_SAVE_NEW.RadioCheckedString;
            }
        }

        private void V_RB_SAVE_OVERWRITE_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_SAVE_OVERWRITE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_SAVE_TYPE.EditValue = V_RB_SAVE_OVERWRITE.RadioCheckedString;
            }
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #region ----- Lookup Event -----
         
        #endregion

        #region ----- Adapter Event -----
         
        #endregion


    }
}