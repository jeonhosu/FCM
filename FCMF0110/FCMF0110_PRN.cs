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

namespace FCMF0110
{
    public partial class FCMF0110_PRN : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        object vEnabled_YN = null;
        string vJob = null; 

        //계정코드 from.
        public object Get_Account_Code_Fr
        {
            get
            {
                return ACCOUNT_CODE_FR.EditValue;
            }
        }

        // 계정코드 To.
        public object Get_Account_Code_To
        {
            get
            {
                return ACCOUNT_CODE_TO.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public FCMF0110_PRN(ISAppInterface pAppInterface, string pOutChoice, object pEnabled_YN)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            vJob = pOutChoice;
            vEnabled_YN = pEnabled_YN;
        }

        #endregion;

        #region ----- Private Methods -----    
        
        private Boolean CheckData()
        {
            if (iString.ISNull(ACCOUNT_CODE_FR.EditValue) != string.Empty && iString.ISNull(ACCOUNT_CONTROL_ID_FR.EditValue) == string.Empty)
            {//계정코드 선택했으나 계정통제 ID값이 없는경우 --> 제대로 선택 안함.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_CODE_FR.Focus();
                return false;
            }
            if (iString.ISNull(ACCOUNT_CODE_TO.EditValue) != string.Empty && iString.ISNull(ACCOUNT_CONTROL_ID_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_CODE_TO.Focus();
                return false;
            }
            return true;
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
        
        #region ----- From Event -----

        private void FCMF0110_PRN_MASTER_Load(object sender, EventArgs e)
        {            
        }

        private void FCMF0110_PRN_Shown(object sender, EventArgs e)
        {
            if (vJob == "FILE")
            {
                vJob = " => Data Excel Export";
            }
            else
            {
                vJob = " => Data Printing...";
            }

            OUTCHOICE.PromptTextElement[0].Default = vJob;
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void btnPRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CheckData() == false)
            {
                return;
            }
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void btnCANCELE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CODE_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ENABLED_YN", vEnabled_YN);
        }

        private void ilaACCOUNT_CODE_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ACCOUNT_CODE_FR", ACCOUNT_CODE_FR.EditValue);
            ildACCOUNT_CONTROL_FR.SetLookupParamValue("W_ENABLED_YN", vEnabled_YN);
        }

        #endregion

    }
}