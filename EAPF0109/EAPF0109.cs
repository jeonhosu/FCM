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


namespace EAPF0109
{
    public partial class EAPF0109 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public EAPF0109()
        {
            InitializeComponent();
        }

        public EAPF0109(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {

                    if (ITB_REJECT_LIST.SelectedTab.TabIndex == tabPageAdv1.TabIndex)
                    {
                        IDA_AUTHORITY_HEADER.Fill();
                        IDA_AUTHORITY_LINE.Fill();
                        IDA_USER_AUTHORITY.Fill();     
                    }
                    else if (ITB_REJECT_LIST.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        IDA_USER_ASSEMBLY_H.Fill();
                        IDA_USER_ASSEMBLY_L.Fill();
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
                    //if (ITB_REJECT_LIST.SelectedTab.TabIndex == tabPageAdv1.TabIndex)
                    //{
                    //    ExcelExport();
                    //}
                    //else if (ITB_REJECT_LIST.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    //{
                    //    ExcelExport();
                    //}
          

                }
            }
        }

        #endregion;




        #region ----- Excel Export -----
        private void ExcelExport(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            GridExcelConverterControl vExport = new GridExcelConverterControl();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog.DefaultExt = ".xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ////데이터 테이블을 이용한 export
                //Syncfusion.XlsIO.ExcelEngine vEng = new Syncfusion.XlsIO.ExcelEngine();
                //Syncfusion.XlsIO.IApplication vApp = vEng.Excel;
                //string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();
                //if (vFileExtension == "XLSX")
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel2007;
                //}
                //else
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel97to2003;
                //}
                //Syncfusion.XlsIO.IWorkbook vWorkbook = vApp.Workbooks.Create(1);
                //Syncfusion.XlsIO.IWorksheet vSheet = vWorkbook.Worksheets[0];
                //foreach(System.Data.DataRow vRow in IDA_MATERIAL_LIST_ALL.CurrentRows)
                //{
                //    vSheet.ImportDataTable(vRow.Table, true, 1, 1, -1, -1);
                //}
                //vWorkbook.SaveAs(saveFileDialog.FileName);
                vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);
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


        #region ----- For Event -----
        
        private void EAPF0109_Load(object sender, EventArgs e)
        {
            //W_DATE_FROM.EditValue = iDate.ISGetDate(DateTime.Today); //iDate.ISDate_Add(iDate.ISGetDate(DateTime.Today), -3);
            //W_DATE_TO.EditValue = iDate.ISGetDate(DateTime.Today);
            //W_DATE_TO.EditValue = iDate.ISGetDate(DateTime.Today);
                     
        }

        #endregion







        private void IDA_MM_PRINT2_FillCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {

        }

        private void IGR_MM_PRINT2_Click(object sender, EventArgs e)
        {

        }

        private void W_DATE_TO_Load(object sender, EventArgs e)
        {

        }

        private void W_DATE_FROM_Load(object sender, EventArgs e)
        {

        }

        private void splitContainerAdv2_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void GB_AUTHORITY_GROUP_Paint(object sender, PaintEventArgs e)
        {

        }

        
    }
}