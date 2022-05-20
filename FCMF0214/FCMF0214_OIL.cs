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

namespace FCMF0214
{
    public partial class FCMF0214_OIL : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0214_OIL()
        {
            InitializeComponent();
        }

        public FCMF0214_OIL(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_SILP_OIL.Fill();
            IGR_SILP_OIL.Focus();

        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            string vMessageText = string.Empty;

            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {        
                    Search_DB(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_SILP_OIL.IsFocused)
                    {
                        IDA_SILP_OIL.AddOver();
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_SILP_OIL.IsFocused)
                    {
                        IDA_SILP_OIL.AddUnder();
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_SILP_OIL.IsFocused)
                    {
                        IDA_SILP_OIL.Update();
                    }   
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_SILP_OIL.IsFocused)
                    {
                        IDA_SILP_OIL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_SILP_OIL.IsFocused)
                    {
                        IDA_SILP_OIL.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Events -----

        private void FCMF0263_Load(object sender, EventArgs e)
        {
            IDA_SILP_OIL.FillSchema();
        }

        private void ILA_OIL_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OIL_TYPE");
        }

        private void ILA_GEARBOX_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "GEARBOX_TYPE");
        }

        private void ILA_DISTANCE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DISTANCE_TYPE");
        }

        private void IDA_SILP_OIL_UpdateCompleted(object pSender)
        {
            Search_DB();
        }

        #endregion;

        

    }
}