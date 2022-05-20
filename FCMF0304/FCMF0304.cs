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

namespace FCMF0304
{
    public partial class FCMF0304 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0304()
        {
            InitializeComponent();
        }

        public FCMF0304(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(PERIOD_FR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PERIOD_FR_0.Focus();
                return;
            }
            if (iString.ISNull(PERIOD_TO_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10219"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PERIOD_TO_0.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                IDA_DPR_LIST_LINE.SetSelectParamValue("W_SOB_ID", -1);
                IDA_DPR_LIST_LINE.Fill();

                IDA_DPR_LIST_HEADER.Fill();
                igrDPR_LIST_HEADER.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                IDA_DPR_LIST_ASSET.Fill();
                IGR_DPR_LIST_ASSET.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
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
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0304_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0304_Shown(object sender, EventArgs e)
        {
            PERIOD_FR_0.EditValue = string.Format("{0}-{1}", iDate.ISYear(DateTime.Today), "01");
            PERIOD_TO_0.EditValue = iDate.ISYearMonth(DateTime.Today);

            //idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "EXPENSE_TYPE");
            //idcDEFAULT_VALUE.ExecuteNonQuery();
            //EXPENSE_TYPE_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            //EXPENSE_TYPE_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "DPR_TYPE");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            DPR_TYPE_NAME_0.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            DPR_TYPE_0.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
        }

        private void TB_MAIN_Click(object sender, EventArgs e)
        {
            //SearchDB();
        }

        private void igrDPR_LIST_HEADER_CellDoubleClick(object pSender)
        {
            if (igrDPR_LIST_HEADER.RowCount < 1)
            {
                return;
            }
            IDA_DPR_LIST_LINE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_DPR_LIST_LINE.Fill();
        }
                    
        #endregion

        #region ----- Lookup Event -----

        private void ilaPERIOD_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD.SetLookupParamValue("W_START_YYYYMM", null);
            ildPERIOD.SetLookupParamValue("W_END_YYYYMM", null);
        }

        private void ilaPERIOD_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD.SetLookupParamValue("W_START_YYYYMM", PERIOD_FR_0.EditValue);
            ildPERIOD.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISGetDate(string.Format("{0}-01", PERIOD_FR_0.EditValue)).AddYears(6)));
        }

        private void ilaEXPENSE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EXPENSE_TYPE", "N");
        }

        private void ilaDPR_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DPR_TYPE", "N");
        }

        private void ilaASSET_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ASSET_TYPE", "N");
        }

        private void ilaASSET_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        #endregion

    }
}