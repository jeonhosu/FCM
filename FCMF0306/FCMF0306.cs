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


namespace FCMF0306
{
    public partial class FCMF0306 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0306()
        {
            InitializeComponent();
        }

        public FCMF0306(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                if (iString.ISNull(W_PERIOD_FR_1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_PERIOD_FR_1.Focus();
                    return;
                }
                if (iString.ISNull(W_PERIOD_TO_1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10219"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_PERIOD_TO_1.Focus();
                    return;
                }
                if (iString.ISNull(W_DPR_TYPE_1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10097"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_DPR_TYPE_DESC_1.Focus();
                    return;                    
                }

                IDA_DPR_STATEMENT_PERIOD.Fill();
                igrDPR_STATEMENT.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                if (iString.ISNull(W_PERIOD_NAME_2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_PERIOD_NAME_2.Focus();
                    return;
                }
                if (iString.ISNull(W_DPR_TYPE_2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10097"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_DPR_TYPE_DESC_2.Focus();
                    return;
                }
                IDA_DPR_STATEMENT_STANDARD.Fill();
                IGR_DPR_STATEMENT_STANDARD.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
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

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (TB_MAIN.SelectedTab.TabIndex == TP_PERIOD.TabIndex)
                    {
                        ExcelExport(igrDPR_STATEMENT);
                    }
                    else if(TB_MAIN.SelectedTab.TabIndex == TP_IN_TIME.TabIndex)
                    {
                        ExcelExport(IGR_DPR_STATEMENT_STANDARD);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0306_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0306_Shown(object sender, EventArgs e)
        {
            W_PERIOD_FR_1.EditValue = string.Format("{0}-{1}", iDate.ISYear(DateTime.Today), "01");
            W_PERIOD_TO_1.EditValue = iDate.ISYearMonth(DateTime.Today);

            //W_CIP_NO.CheckedString = "Checked";
            W_CIP_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            T2_W_CIP_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_CIP_FLAG.EditValue = W_CIP_NO.RadioButtonString;
            T2_W_CIP_FLAG.EditValue = T2_W_CIP_NO.RadioButtonString;


            W_PERIOD_NAME_2.EditValue = iDate.ISYearMonth(DateTime.Today);

            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "DPR_TYPE");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            W_DPR_TYPE_DESC_1.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_DPR_TYPE_1.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");

            W_DPR_TYPE_DESC_2.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_DPR_TYPE_2.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaPERIOD_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", null);
        }

        private void ilaPERIOD_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", W_PERIOD_FR_1.EditValue);
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddYears(1));
        }

        private void ilaDPR_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DPR_TYPE", "N");
        }

        private void ILA_ASSET_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ASSET_TYPE", "N");
        }

        private void ILA_EXPENSE_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EXPENSE_TYPE", "N");
        }

        private void ILA_ASSET_CATEGORY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_PERIOD_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", null);
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddYears(1));
        }

        private void ILA_DPR_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DPR_TYPE", "N");
        }

        private void ILA_ASSET_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ASSET_TYPE", "N");
        }

        private void ILA_EXPENSE_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EXPENSE_TYPE", "N");
        }

        private void ILA_ASSET_CODE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_ASSET_CATEGORY_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        #endregion

        private void W_CIP_NO_Click(object sender, EventArgs e)
        {
            if (W_CIP_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_NO.RadioButtonString;
            }
        }

        private void W_CIP_YES_Click(object sender, EventArgs e)
        {
            if (W_CIP_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_YES.RadioButtonString;
            }
        }

        private void T2_W_CIP_NO_Click(object sender, EventArgs e)
        {
            if (T2_W_CIP_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                T2_W_CIP_FLAG.EditValue = T2_W_CIP_NO.RadioButtonString;
            }
        }

        private void T2_W_CIP_YES_Click(object sender, EventArgs e)
        {
            if (T2_W_CIP_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                T2_W_CIP_FLAG.EditValue = T2_W_CIP_YES.RadioButtonString;
            }
        }

    }
}