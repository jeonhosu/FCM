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

namespace FCMF0991
{
    public partial class FCMF0991 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        #endregion;

        #region ----- Constructor -----

        public FCMF0991()
        {
            InitializeComponent();
        }

        public FCMF0991(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_TAX_BILL_USER.Fill();
            IGR_TAX_BILL_USER.Focus();
        }

        private void Init_Insert()
        {
            IGR_TAX_BILL_USER.SetCellValue("ENABLED_FLAG", "Y");
            IGR_TAX_BILL_USER.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));

            IGR_TAX_BILL_USER.Focus();
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
                    if (IDA_TAX_BILL_USER.IsFocused)
                    {
                        IDA_TAX_BILL_USER.AddOver();
                        Init_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_TAX_BILL_USER.IsFocused)
                    {
                        IDA_TAX_BILL_USER.AddUnder();
                        Init_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                   IDA_TAX_BILL_USER.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_TAX_BILL_USER.IsFocused)
                    {
                        IDA_TAX_BILL_USER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_TAX_BILL_USER.IsFocused)
                    {
                        if (IDA_TAX_BILL_USER.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_TAX_BILL_USER.Delete();
                        }
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Evevnt -----

        private void FCMF0991_Load(object sender, EventArgs e)
        {
            IDA_TAX_BILL_USER.FillSchema();
        }
        
        #endregion

        #region ----- Lookup Event ------
        
        private void ILA_CUSTOMER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event ------

        private void IDA_TAX_BILL_USER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["USER_NO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USER_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10126"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}