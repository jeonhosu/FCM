using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0951
{
    public partial class FCMF0951 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0951()
        {
            InitializeComponent();
        }

        public FCMF0951(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            IDA_OPERATION_COST.Fill();
            ISG_OPERATION_COST.Focus();
        }

        private void Init_Insert()
        {
            ISG_OPERATION_COST_DIST.SetCellValue("ENABLED_FLAG", "Y");
            ISG_OPERATION_COST_DIST.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            ISG_OPERATION_COST_DIST.Focus();
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
                    if (IDA_OPERATION_COST_DIST.IsFocused)
                    {
                        IDA_OPERATION_COST_DIST.AddOver();
                        Init_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_OPERATION_COST_DIST.IsFocused)
                    {
                        IDA_OPERATION_COST_DIST.AddUnder();
                        Init_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_OPERATION_COST.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_OPERATION_COST_DIST.IsFocused)
                    {
                        IDA_OPERATION_COST_DIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_OPERATION_COST_DIST.IsFocused)
                    {
                        IDA_OPERATION_COST_DIST.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ---- Lookup Event -----

        private void ILA_FR_COST_CENTER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_FR_COST_CENTER.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_TO_COST_CENTER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_TO_COST_CENTER.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_OPERATION_COST_DIST_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["FR_COST_CENTER_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(string.Format("From C/C :: {0}", isMessageAdapter1.ReturnText("FCM_10524")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["COST_CENTER_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(string.Format("To C/C :: {0}", isMessageAdapter1.ReturnText("FCM_10524")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["ENABLED_FLAG"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10085"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void IDA_OPERATION_COST_DIST_PreDelete(ISPreDeleteEventArgs e)
        {
            if (IDA_OPERATION_COST_DIST.CurrentRow.RowState != DataRowState.Added)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10307"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion
        
    }
}