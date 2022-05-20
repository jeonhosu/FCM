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

namespace FCMF0804
{
    public partial class TAX_INVOICE_DETAIL : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mTAX_CODE;
        object mVAT_GUBUN;
        object mISSUE_DATE_FR;
        object mISSUE_DATE_TO;
        object mCUSTOMER_ID;

        #endregion;

        #region ----- Constructor -----

        public TAX_INVOICE_DETAIL(ISAppInterface pAppInterface, object pTAX_CODE, object pVAT_GUBUN
                                , object pISSUE_DATE_FR, object pISSUE_DATE_TO, object pCUSTOMER_ID)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mTAX_CODE = pTAX_CODE;
            mVAT_GUBUN = pVAT_GUBUN;
            mISSUE_DATE_FR = pISSUE_DATE_FR;
            mISSUE_DATE_TO = pISSUE_DATE_TO;
            mCUSTOMER_ID = pCUSTOMER_ID;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            idaTAX_INVOICE_DETAIL.SetSelectParamValue("W_TAX_CODE", mTAX_CODE);
            idaTAX_INVOICE_DETAIL.SetSelectParamValue("W_VAT_GUBUN", mVAT_GUBUN);
            idaTAX_INVOICE_DETAIL.SetSelectParamValue("W_ISSUE_DATE_FR", mISSUE_DATE_FR);
            idaTAX_INVOICE_DETAIL.SetSelectParamValue("W_ISSUE_DATE_TO", mISSUE_DATE_TO);
            idaTAX_INVOICE_DETAIL.SetSelectParamValue("W_CUSTOMER_ID", mCUSTOMER_ID);
            idaTAX_INVOICE_DETAIL.Fill();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {

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
        
        private void TAX_INVOICE_DETAIL_Shown(object sender, EventArgs e)
        {
            Search_DB();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }

        private void btnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }
        
        #endregion

        #region ------ Lookup Event ------

        #endregion

        #region ------ Adapter Event ------

        #endregion             

    }
}