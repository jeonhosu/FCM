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
using Syncfusion.XlsIO;

namespace EAPF0111
{
    public partial class EAPF0111 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public EAPF0111()
        {
            InitializeComponent();
        }

        public EAPF0111(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_NAVI_AUTHORITY.Fill();
            IGR_NAVI_AUTHORITY.Focus();
        }

        private void Init_Insert_User()
        {
            IGR_EAPP_USER.SetCellValue("READ_FLAG", "Y");
            IGR_EAPP_USER.Focus();
        }

        private void Sync_CheckBox(int pIDX_Col, object vCheckValue)
        {
            for (int r = 0; r < IGR_EAPP_USER.RowCount; r++)
            {
                IGR_EAPP_USER.CurrentCellMoveTo(r, pIDX_Col);
                IGR_EAPP_USER.CurrentCellActivate(r, pIDX_Col);
                IGR_EAPP_USER.SetCellValue(r, pIDX_Col, vCheckValue);
            }
            IGR_EAPP_USER.CurrentCellMoveTo(0, pIDX_Col);
            IGR_EAPP_USER.CurrentCellActivate(0, pIDX_Col);
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
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.AddOver();
                        Init_Insert_User();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.AddUnder();
                        Init_Insert_User();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_NAVI_AUTHORITY); 
                }
            }
        }

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

        #region ----- Form Event -----

        private void EAPF0111_Load(object sender, EventArgs e)
        {
            IDA_NAVI_AUTHORITY.FillSchema();
            IDA_EAPP_USER.FillSchema();

            V_READ.BringToFront();
            V_WRITE.BringToFront();
            V_PRINT.BringToFront();
        }

        private void V_READ_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Sync_CheckBox(IGR_EAPP_USER.GetColumnToIndex("READ_FLAG"), V_READ.CheckBoxValue);
        }

        private void V_WRITE_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Sync_CheckBox(IGR_EAPP_USER.GetColumnToIndex("WRITE_FLAG"), V_WRITE.CheckBoxValue);
        }

        private void V_PRINT_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Sync_CheckBox(IGR_EAPP_USER.GetColumnToIndex("PRINT_FLAG"), V_PRINT.CheckBoxValue);
        }

        #endregion

        #region ---- Lookup Event -----

        private void ILA_DEPT_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_EAPP_USER_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_EAPP_USER.SetLookupParamValue("W_ASSEMBLY_INFO_ID", DBNull.Value);
        }

        private void ILA_EAPP_USER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_EAPP_USER.SetLookupParamValue("W_ASSEMBLY_INFO_ID", IGR_NAVI_AUTHORITY.GetCellValue("ASSEMBLY_INFO_ID"));
        }

        #endregion

        #region ---- Adatper Event -----

        private void IDA_EAPP_USER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull("USER_ID") == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iString.ISNull("READ_FLAG") == "N" && iString.ISNull("WRITE_FLAG") == "N" && iString.ISNull("PRINT_FLAG") == "N")
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10120"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion


    }
}