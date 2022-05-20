using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.IO;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0208
{
    public partial class FCMF0208 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();
        Object mSESSION_ID;

        #region ----- Variables -----


        #endregion;

        #region ----- Constructor -----

        public FCMF0208()
        {
            InitializeComponent();
        }

        public FCMF0208(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void Search_DB()
        { 
            IDA_SLIP_EXCEL.SetSelectParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_SLIP_EXCEL.Fill();
            IGR_SLIP_EXCEL.Focus();
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

        #region ----- Excel Upload -----
         
        private bool Set_Slip_Transfer()
        {
            DialogResult dlgResult;
            dlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10303"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlgResult == DialogResult.No)
            {
                return false;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            object vMESSAGE = string.Empty;

            IDC_SET_SLIP_TRANSFER.ExecuteNonQuery();
            vSTATUS = IDC_SET_SLIP_TRANSFER.GetCommandParamValue("O_STATUS").ToString();
            vMESSAGE = IDC_SET_SLIP_TRANSFER.GetCommandParamValue("O_MESSAGE");
            if (IDC_SET_SLIP_TRANSFER.ExcuteError == true || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (iString.ISNull(vMESSAGE) != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            return true;
        }

        private bool Cancel_Slip_Transfer()
        {
            DialogResult dlgResult;
            dlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlgResult == DialogResult.No)
            {
                return false;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            object vMESSAGE = String.Empty;

            IDC_CANCEL_SLIP_TRANSFER.ExecuteNonQuery();
            vSTATUS = IDC_CANCEL_SLIP_TRANSFER.GetCommandParamValue("O_STATUS").ToString();
            vMESSAGE = IDC_CANCEL_SLIP_TRANSFER.GetCommandParamValue("O_MESSAGE");
            if (IDC_CANCEL_SLIP_TRANSFER.ExcuteError == true || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (iString.ISNull(vMESSAGE) != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            return true;
        }

        #endregion

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


        private void FCMF0208_Load(object sender, EventArgs e)
        {
            IDA_SLIP_EXCEL.FillSchema(); 
        }

        private void FCMF0208_Shown(object sender, EventArgs e)
        {
            
        }
         
        private void BTN_EXPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            FCMF0208_EXPORT vFCMF0208_EXPORT = new FCMF0208_EXPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0208_EXPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vFCMF0208_EXPORT.ShowDialog();
            vFCMF0208_EXPORT.Dispose();
        }

        private void BTN_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            FCMF0208_IMPORT vFCMF0208_IMPORT = new FCMF0208_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSESSION_ID);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vFCMF0208_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vFCMF0208_IMPORT.ShowDialog();
            vFCMF0208_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void BTN_SLIP_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(Set_Slip_Transfer() == false)
            {
                return;
            }
            Search_DB();
        }

        private void BTN_SLIP_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(Cancel_Slip_Transfer() == false)
            {
                return;
            }
            Search_DB();
        }
        
        #endregion

        #region ----- Lookup Event -----


        #endregion

        #region ----- Adapeter Event -----
    
        #endregion


    }
}