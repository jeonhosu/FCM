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


namespace EAPF0110
{
    public partial class EAPF0110 : Office2007Form
    {
        ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----


        private bool mIsAllSelectRead = false;
        private bool mIsAllSelectWrite = false;
        private bool mIsAllSelectPrint = false;
        private bool mIsAllSelect = false;

        #endregion;

        #region ----- Constructor -----

        public EAPF0110()
        {
            InitializeComponent();
        }

        public EAPF0110(Form pMainForm, ISAppInterface pAppInterface)
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
                        IDA_PROGRAM.Fill();
                 }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_USER.IsFocused)
                    {
                        IDA_USER.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_USER.IsFocused)
                    {
                        IDA_USER.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {                    
                    IDA_PROGRAM.Update();
                //    IDA_USER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                   
                        IDA_USER.Cancel();
                    
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
        
        private void EAPF0110_Load(object sender, EventArgs e)
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

      

        private void IGR_USER_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        
        {
                        
    

        }


        #region ----- Select All Assembly Method -----

        private void SelectAllAssembly(ISGridAdvEx pGrid)
        {
            int vCountRows = pGrid.RowCount;
            if (vCountRows > 0)
            {
                int vIndexCheckBox1 = pGrid.GetColumnToIndex("READ_FLAG");
                string vCheckedString1 = pGrid.GridAdvExColElement[vIndexCheckBox1].CheckedString;
                string vUnCheckedString1 = pGrid.GridAdvExColElement[vIndexCheckBox1].UncheckedString;

                int vIndexCheckBox2 = pGrid.GetColumnToIndex("WRITE_FLAG");
                string vCheckedString2 = pGrid.GridAdvExColElement[vIndexCheckBox2].CheckedString;
                string vUnCheckedString2 = pGrid.GridAdvExColElement[vIndexCheckBox2].UncheckedString;

                int vIndexCheckBox3 = pGrid.GetColumnToIndex("PRINT_FLAG");
                string vCheckedString3 = pGrid.GridAdvExColElement[vIndexCheckBox3].CheckedString;
                string vUnCheckedString3 = pGrid.GridAdvExColElement[vIndexCheckBox3].UncheckedString;

                for (int vRow = 0; vRow < vCountRows; vRow++)
                {
                    if (mIsAllSelect == true)
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox1, vCheckedString1);
                        pGrid.SetCellValue(vRow, vIndexCheckBox2, vCheckedString2);
                        pGrid.SetCellValue(vRow, vIndexCheckBox3, vCheckedString3);
                    }
                    else
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox1, vUnCheckedString1);
                        pGrid.SetCellValue(vRow, vIndexCheckBox2, vUnCheckedString2);
                        pGrid.SetCellValue(vRow, vIndexCheckBox3, vUnCheckedString3);
                    }
                }

                if (mIsAllSelect == true)
                {
                    int vMoveRow = vCountRows - 1;
                    pGrid.CurrentCellMoveTo(vMoveRow, vIndexCheckBox1);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(vMoveRow, vIndexCheckBox1);
                }
                else
                {
                    pGrid.CurrentCellMoveTo(0, vIndexCheckBox1);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(0, vIndexCheckBox1);
                }


            }
        }

        #endregion;

        #region ----- Select All Assembly Print Method -----

        private void SelectAllAssemblyPrint(ISGridAdvEx pGrid)
        {
            int vCountRows = pGrid.RowCount;
            if (vCountRows > 0)
            {
                int vIndexCheckBox = pGrid.GetColumnToIndex("PRINT_FLAG");
                string vCheckedString = pGrid.GridAdvExColElement[vIndexCheckBox].CheckedString;
                string vUnCheckedString = pGrid.GridAdvExColElement[vIndexCheckBox].UncheckedString;

                for (int vRow = 0; vRow < vCountRows; vRow++)
                {
                    if (mIsAllSelectPrint == true)
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox, vCheckedString);
                    }
                    else
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox, vUnCheckedString);
                    }
                }

                if (mIsAllSelectPrint == true)
                {
                    int vMoveRow = vCountRows - 1;
                    pGrid.CurrentCellMoveTo(vMoveRow, vIndexCheckBox);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(vMoveRow, vIndexCheckBox);
                }
                else
                {
                    pGrid.CurrentCellMoveTo(0, vIndexCheckBox);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(0, vIndexCheckBox);
                }
            }
        }

        #endregion;

        #region ----- Select All Assembly Read Method -----

        private void SelectAllAssemblyRead(ISGridAdvEx pGrid)
        {
            int vCountRows = pGrid.RowCount;
            if (vCountRows > 0)
            {
                int vIndexCheckBox = pGrid.GetColumnToIndex("READ_FLAG");
                string vCheckedString = pGrid.GridAdvExColElement[vIndexCheckBox].CheckedString;
                string vUnCheckedString = pGrid.GridAdvExColElement[vIndexCheckBox].UncheckedString;

                for (int vRow = 0; vRow < vCountRows; vRow++)
                {
                    if (mIsAllSelectRead == true)
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox, vCheckedString);
                    }
                    else
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox, vUnCheckedString);
                    }
                }

                if (mIsAllSelectRead == true)
                {
                    int vMoveRow = vCountRows - 1;
                    pGrid.CurrentCellMoveTo(vMoveRow, vIndexCheckBox);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(vMoveRow, vIndexCheckBox);
                }
                else
                {
                    pGrid.CurrentCellMoveTo(0, vIndexCheckBox);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(0, vIndexCheckBox);
                }
            }
        }

        #endregion;

        #region ----- Select All Assembly Write Method -----

        private void SelectAllAssemblyWrite(ISGridAdvEx pGrid)
        {
            int vCountRows = pGrid.RowCount;
            if (vCountRows > 0)
            {
                int vIndexCheckBox = pGrid.GetColumnToIndex("WRITE_FLAG");
                string vCheckedString = pGrid.GridAdvExColElement[vIndexCheckBox].CheckedString;
                string vUnCheckedString = pGrid.GridAdvExColElement[vIndexCheckBox].UncheckedString;

                for (int vRow = 0; vRow < vCountRows; vRow++)
                {
                    if (mIsAllSelectWrite == true)
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox, vCheckedString);
                    }
                    else
                    {
                        pGrid.SetCellValue(vRow, vIndexCheckBox, vUnCheckedString);
                    }
                }

                if (mIsAllSelectWrite == true)
                {
                    int vMoveRow = vCountRows - 1;
                    pGrid.CurrentCellMoveTo(vMoveRow, vIndexCheckBox);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(vMoveRow, vIndexCheckBox);
                }
                else
                {
                    pGrid.CurrentCellMoveTo(0, vIndexCheckBox);
                    pGrid.Focus();
                    pGrid.CurrentCellActivate(0, vIndexCheckBox);
                }
            }
        }

        #endregion;

        private void IGR_USER_Click(object sender, EventArgs e)
        {

        }

        private void isCheckBoxAdv3_CheckedChange(object pSender, ISCheckEventArgs e)  //PRINT
        {
            if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                mIsAllSelectPrint = true;
            }
            else
            {
                mIsAllSelectPrint = false;
            }
            SelectAllAssemblyPrint(IGR_USER);
        }

        private void isCheckBoxAdv2_CheckedChange(object pSender, ISCheckEventArgs e)  //WRITE
        {
            if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                mIsAllSelectWrite = true;
            }
         
            else
            {
                mIsAllSelectWrite = false;
            }
            SelectAllAssemblyWrite(IGR_USER);
        }


        private void isCheckBoxAdv1_CheckedChange(object pSender, ISCheckEventArgs e) //READ
        {
            if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                mIsAllSelectRead = true;
            }
            else
            {
                mIsAllSelectRead = false;
            }
            SelectAllAssemblyRead(IGR_USER);
        }

        private void isCheckBoxAdv4_CheckedChange(object pSender, ISCheckEventArgs e)  //ALL
        {
            if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                mIsAllSelect = true;
            }
            else
            {
                mIsAllSelect = false;
            }
            SelectAllAssembly(IGR_USER);
        }

        private void IGR_ASSEMBLY_Click(object sender, EventArgs e)
        {

        }

        private void ILA_PERSON_SelectedRowData(object pSender)
        {
            string vUserId = Convert.ToString(IGR_USER.GetCellValue("USER_ID"));
            string vLineUserId = string.Empty;

            for (int i = 0; i < IGR_USER.RowCount; i++)
            {
                if (i != IGR_USER.RowIndex)
                {
                    vLineUserId = Convert.ToString(IGR_USER.GetCellValue(i, IGR_USER.GetColumnToIndex("USER_ID")));

                    if (vUserId == vLineUserId)
                    {
                        string vMessageGet = isMessageAdapter1.ReturnText("EAPP_10191"); //이미 있는 사용자입니다.
                        string vMessageString = string.Format("{0}", vMessageGet);
                        MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        IGR_USER.SetCellValue(IGR_USER.RowIndex, IGR_USER.GetColumnToIndex("USER_ID"), "");
                        IGR_USER.SetCellValue(IGR_USER.RowIndex, IGR_USER.GetColumnToIndex("USER_NO"), "");
                        IGR_USER.SetCellValue(IGR_USER.RowIndex, IGR_USER.GetColumnToIndex("PERSON_NUM"), "");
                        IGR_USER.SetCellValue(IGR_USER.RowIndex, IGR_USER.GetColumnToIndex("DEPT_NAME"), "");
                        IGR_USER.SetCellValue(IGR_USER.RowIndex, IGR_USER.GetColumnToIndex("DESCRIPTION"), "");
                    }
                    

                }
            }
        }

        private void ilaDEPT_MASTER_SelectedRowData(object pSender)
        {
            IDA_USER.Fill();
        }

        private void IDA_PROGRAM_ExcuteKeySearch(object pSender)
        {

        }

        private void ILA_PERSON_W_SelectedRowData(object pSender)
        {
            IDA_USER.Fill();
        }


    }
}