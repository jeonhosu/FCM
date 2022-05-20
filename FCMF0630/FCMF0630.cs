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

namespace FCMF0630
{
    public partial class FCMF0630 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0630()
        {
            InitializeComponent();
        }

        public FCMF0630(Form pMainForm, ISAppInterface pAppInterface)
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

            IDA_BUDGET_MONTH_WEEK.Fill();
            IGR_BUDGET_MONTH_WEEK.Focus();
        }
         
        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ILD_COMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }
         
        private void Set_Budget_Month_Week()
        {
            IGR_BUDGET_MONTH_WEEK.SetCellValue("BUDGET_PERIOD", W_BUDGET_PERIOD.EditValue);
            IGR_BUDGET_MONTH_WEEK.Focus();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
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
                    if (IDA_BUDGET_MONTH_WEEK.IsFocused)
                    {
                        IDA_BUDGET_MONTH_WEEK.AddOver();
                        Set_Budget_Month_Week();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BUDGET_MONTH_WEEK.IsFocused)
                    {
                        IDA_BUDGET_MONTH_WEEK.AddUnder();
                        Set_Budget_Month_Week();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_MONTH_WEEK.Update();
                    }
                    catch
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_MONTH_WEEK.IsFocused)
                    {
                        IDA_BUDGET_MONTH_WEEK.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_MONTH_WEEK.IsFocused)
                    {
                        IDA_BUDGET_MONTH_WEEK.Delete();
                    } 
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

        private void FCMF0630_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_MONTH_WEEK.FillSchema();  
        }

        private void FCMF0630_Shown(object sender, EventArgs e)
        {
            W_BUDGET_PERIOD.EditValue = iDate.ISYearMonth(DateTime.Today); 
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }  

        #endregion
        
        #region ----- Lookup Event -----
         
        private void ILA_PERIOD_W_SelectedRowData(object pSender)
        {
            SearchDB();
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

        private void ILA_PERIOD_NAME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        } 
          
        #endregion

        #region ----- Adapter Event -----

        private void IDA_BUDGET_MONTH_WEEK_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_PERIOD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10442"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WEEK_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("WIP_10325"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WEEK_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WEEK_DATE_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        } 
          
        #endregion

    }
}