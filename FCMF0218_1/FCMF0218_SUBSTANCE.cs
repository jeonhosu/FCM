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

namespace FCMF0218
{
    public partial class FCMF0218_SUBSTANCE : Office2007Form
    {
        #region ----- Variables -----

        object mHEADER_INTERFACE_ID = null;

        #endregion;

        #region ----- Constructor -----

        public FCMF0218_SUBSTANCE(ISAppInterface pAppInterface, object pHEADER_INTERFACE_ID)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mHEADER_INTERFACE_ID = pHEADER_INTERFACE_ID;
        }

        #endregion;

        #region ----- Private Methods ----



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

        private void FCMF0218_SUBSTANCE_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0218_SUBSTANCE_Shown(object sender, EventArgs e)
        {
            idaSUBSTANCE_IF.SetSelectParamValue("W_HEADER_INTERFACE_ID", mHEADER_INTERFACE_ID);
            idaSUBSTANCE_IF.Fill();
        }

        private void BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }
        
        #endregion

    }
}