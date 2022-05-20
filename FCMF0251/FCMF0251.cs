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


namespace FCMF0251
{
    public partial class FCMF0251 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        //object mAccount_Book_ID;
        //object mAccount_Set_ID;
        //object mFiscal_Calendar_ID;
        //object mDept_Level;
        //object mAccount_Book_Name;
        //object mCurrency_Code;
        //object mBudget_Control_YN; 

        #endregion;
        
        #region ----- Constructor -----

        public FCMF0251(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void GetAccountBook()
        {
             
        }

        private void Search()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_SLIP_MONTH_COMPARE.TabIndex)
            {
                if (iString.ISNull(P_COMPARE_PERIOD.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(P_COMPARE_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    P_COMPARE_PERIOD.Focus();
                    return;
                }
                if (iString.ISNull(P_BASE_PERIOD.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(P_COMPARE_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    P_BASE_PERIOD.Focus();
                    return;
                }
                IDA_SLIP_MONTH_COMPARE.Fill();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        ////조회된 자료에서 더블클릭하면 전표팝업 띄워준다.
        //private void Show_Slip_Detail(int pSLIP_HEADER_ID)
        //{
        //    if (pSLIP_HEADER_ID != 0)
        //    {
        //        Application.UseWaitCursor = true;
        //        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

        //        FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, pSLIP_HEADER_ID);
        //        vFCMF0204.Show();

        //        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        //        Application.UseWaitCursor = false;
        //    }
        //}
         
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

        #region ----- XLS Print Method ----

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
                vExport.ExportStyle = false;
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
                    //XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                {
                    ExcelExport(IGR_SLIP_MONTH_COMPARE);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0251_Load(object sender, EventArgs e)
        {
              
        }

        private void FCMF0251_Shown(object sender, EventArgs e)
        {
            GetAccountBook();

            P_COMPARE_PERIOD.EditValue = iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, -1));
            P_BASE_PERIOD.EditValue = iDate.ISYearMonth(DateTime.Today);
        }
           
        #endregion


        #region ----- Lookup Event -----
  
        private void ILA_ACCOUNT_PROPERTY_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("ACCOUNT_PROPERTY", "Y");
        }


        private void ILA_COMPARE_PERIOD_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", null);
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddMonths(3));
        }

        private void ILA_COMPARE_PERIOD_SelectedRowData(object pSender)
        {
            P_BASE_PERIOD.EditValue = P_COMPARE_PERIOD.EditValue;
        }

        private void ILA_BASE_PERIOD_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_START_YYYYMM", P_COMPARE_PERIOD.EditValue);
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", DateTime.Today.AddMonths(3));
        } 

        #endregion

        #region ----- Adapter Lookup Event -----
                    
        #endregion

        

    }
}