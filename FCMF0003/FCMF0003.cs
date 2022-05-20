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

namespace FCMF0003
{
    public partial class FCMF0003 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public FCMF0003()
        {
            InitializeComponent();
        }

        public FCMF0003(Form pMainForm, ISAppInterface pAppInterface)
        { 
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        private void Search_DB()
        {
            idaDOCUMENT_NUM.Fill();
            igrDOCUMENT_NUM.Focus();
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
                    if (idaDOCUMENT_NUM.IsFocused)
                    {
                        idaDOCUMENT_NUM.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaDOCUMENT_NUM.IsFocused)
                    {
                        idaDOCUMENT_NUM.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaDOCUMENT_NUM.IsFocused)
                    {
                        idaDOCUMENT_NUM.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaDOCUMENT_NUM.IsFocused)
                    {
                        idaDOCUMENT_NUM.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaDOCUMENT_NUM.IsFocused)
                    {
                        idaDOCUMENT_NUM.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        private void FCMF0003_Load(object sender, EventArgs e)
        {
            idaDOCUMENT_NUM.FillSchema();
        }

        #endregion

        #region ----- Adapter Event ----
        private void idaDOCUMENT_NUM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DOCU_NUM_CLASS"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10084"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DOCUMENT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10104"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void idaDOCUMENT_NUM_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ---- Lookup Event ----

        private void ILA_DOCU_NUM_CLASS_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DOCU_NUM_CLASS");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_DOCU_NUM_CLASS_V_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DOCU_NUM_CLASS");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_DATE_FORMAT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DATE_FORMAT");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        #endregion

    }
}