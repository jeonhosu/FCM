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

namespace FCMF0818
{
    public partial class BILL_DETAIL : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mTAX_CODE;
        object mVAT_GUBUN;
        object mISSUE_DATE_FR;
        object mISSUE_DATE_TO;
        object mCUSTOMER_ID;
        object mBUSINESS_UNIT_TAX_YN;

        #endregion;

        #region ----- Constructor -----

        public BILL_DETAIL(ISAppInterface pAppInterface, object pTAX_CODE, object pVAT_GUBUN
                                , object pISSUE_DATE_FR, object pISSUE_DATE_TO, object pCUSTOMER_ID
                                , object pBUSINESS_UNIT_TAX_YN)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mTAX_CODE = pTAX_CODE;
            mVAT_GUBUN = pVAT_GUBUN;
            mISSUE_DATE_FR = pISSUE_DATE_FR;
            mISSUE_DATE_TO = pISSUE_DATE_TO;
            mCUSTOMER_ID = pCUSTOMER_ID;
            mBUSINESS_UNIT_TAX_YN = pBUSINESS_UNIT_TAX_YN;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_BILL_DETAIL.SetSelectParamValue("W_TAX_CODE", mTAX_CODE);
            IDA_BILL_DETAIL.SetSelectParamValue("W_VAT_GUBUN", mVAT_GUBUN);
            IDA_BILL_DETAIL.SetSelectParamValue("W_ISSUE_DATE_FR", mISSUE_DATE_FR);
            IDA_BILL_DETAIL.SetSelectParamValue("W_ISSUE_DATE_TO", mISSUE_DATE_TO);
            IDA_BILL_DETAIL.SetSelectParamValue("W_CUSTOMER_ID", mCUSTOMER_ID);
            IDA_BILL_DETAIL.SetSelectParamValue("W_BUSINESS_UNIT_TAX_YN", mBUSINESS_UNIT_TAX_YN);
            IDA_BILL_DETAIL.Fill();
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
            this.Close();
        }
        
        #endregion

        #region ------ Lookup Event ------

        #endregion

        #region ------ Adapter Event ------

        #endregion             

    }
}