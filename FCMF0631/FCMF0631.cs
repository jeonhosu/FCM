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

namespace FCMF0631
{
    public partial class FCMF0631 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0631()
        {
            InitializeComponent();
        }

        public FCMF0631(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
         
        private void SearchDB()
        {
            if (iString.ISNull(W_BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(W_WEEK_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_WEEK_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_NAME.Focus();
                return;
            }
            INIT_WEEK_COLUMN();
            Application.DoEvents();

            IDA_BUDGET_MONTH_USE.Fill();
            IGR_BUDGET_MONTH_USE.Focus();
        }
         
        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ILD_COMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }
         
        private void Set_Budget_Month_Week()
        {
            IGR_BUDGET_MONTH_USE.SetCellValue("BUDGET_PERIOD", W_BUDGET_PERIOD.EditValue);
            IGR_BUDGET_MONTH_USE.Focus();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void INIT_WEEK_COLUMN()
        {
            string vTITLE = string.Empty;
            IDA_BUDGET_MONTH_WEEK_TIT.Fill(); 
            if (IDA_BUDGET_MONTH_WEEK_TIT.CurrentRows.Count == 0)
            {
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_PLAN_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_USE_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_GAP_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_USE_RATE")].Visible = 0;

                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_PLAN_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_USE_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_GAP_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_USE_RATE")].Visible = 0;

                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_PLAN_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_USE_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_GAP_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_USE_RATE")].Visible = 0;

                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_PLAN_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_USE_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_GAP_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_USE_RATE")].Visible = 0;

                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_PLAN_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_USE_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_GAP_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_USE_RATE")].Visible = 0;

                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_PLAN_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_USE_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_GAP_AMOUNT")].Visible = 0;
                IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_USE_RATE")].Visible = 0;
            }
            else
            {
                vTITLE = iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_1"]);
                if(iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_1_FLAG"]) == "Y")
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_PLAN_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_USE_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_GAP_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_USE_RATE")].Visible = 1;
                }
                else
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_PLAN_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_USE_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_GAP_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_1_USE_RATE")].Visible = 0;
                }

                vTITLE = iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_2"]);
                if (iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_2_FLAG"]) == "Y")
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_PLAN_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_USE_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_GAP_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_USE_RATE")].Visible = 1;
                }
                else
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_PLAN_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_USE_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_GAP_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_2_USE_RATE")].Visible = 0;
                }

                vTITLE = iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_3"]);
                if (iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_3_FLAG"]) == "Y")
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_PLAN_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_USE_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_GAP_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_USE_RATE")].Visible = 1;
                }
                else
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_PLAN_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_USE_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_GAP_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_3_USE_RATE")].Visible = 0;
                }

                vTITLE = iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_4"]);
                if (iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_4_FLAG"]) == "Y")
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_PLAN_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_USE_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_GAP_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_USE_RATE")].Visible = 1;
                }
                else
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_PLAN_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_USE_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_GAP_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_4_USE_RATE")].Visible = 0;
                }

                vTITLE = iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_5"]);
                if (iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_5_FLAG"]) == "Y")
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_PLAN_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_USE_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_GAP_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_USE_RATE")].Visible = 1;
                }
                else
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_PLAN_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_USE_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_GAP_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_5_USE_RATE")].Visible = 0;
                }

                vTITLE = iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_6"]);
                if (iString.ISNull(IDA_BUDGET_MONTH_WEEK_TIT.CurrentRow["WEEK_6_FLAG"]) == "Y")
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_PLAN_AMOUNT")].HeaderElement[1].Default = vTITLE;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_PLAN_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_USE_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_GAP_AMOUNT")].Visible = 1;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_USE_RATE")].Visible = 1;
                }
                else
                {
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_PLAN_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_USE_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_GAP_AMOUNT")].Visible = 0;
                    IGR_BUDGET_MONTH_USE.GridAdvExColElement[IGR_BUDGET_MONTH_USE.GetColumnToIndex("WEEK_6_USE_RATE")].Visible = 0;
                }
            }
            IGR_BUDGET_MONTH_USE.ResetDraw = true;
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
        
        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx vGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog.DefaultExt = ".xls";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                vExport.GridToExcel(vGrid.BaseGrid, saveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);

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
                    if (IDA_BUDGET_MONTH_USE.IsFocused)
                    {
                        IDA_BUDGET_MONTH_USE.AddOver();
                        Set_Budget_Month_Week();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BUDGET_MONTH_USE.IsFocused)
                    {
                        IDA_BUDGET_MONTH_USE.AddUnder();
                        Set_Budget_Month_Week();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_MONTH_USE.Update();
                    }
                    catch
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_MONTH_USE.IsFocused)
                    {
                        IDA_BUDGET_MONTH_USE.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_MONTH_USE.IsFocused)
                    {
                        IDA_BUDGET_MONTH_USE.Delete();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_BUDGET_MONTH_USE);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0631_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_MONTH_USE.FillSchema();  
        }

        private void FCMF0631_Shown(object sender, EventArgs e)
        {
            W_BUDGET_PERIOD.EditValue = iDate.ISYearMonth(DateTime.Today); 
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }  

        #endregion
        
        #region ----- Lookup Event -----
         
        private void ILA_PERIOD_W_SelectedRowData(object pSender)
        {
             
        }

        private void ILA_DEPT_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "N");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(W_BUDGET_PERIOD.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(W_BUDGET_PERIOD.EditValue));
        }

        private void ILA_MONTH_WEEK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("MONTH_WEEK", DBNull.Value, "Y");
        }

        private void ILA_PERIOD_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_ACCOUNT_CONTROL_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_CONTROL_YN", "N");
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_MONTH_WEEK_SelectedRowData(object pSender)
        {
            INIT_WEEK_COLUMN();
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_BUDGET_MONTH_WEEK_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        } 
          
        #endregion

    }
}