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

namespace FCMF0813
{
    public partial class FCMF0813_CREATE : Office2007Form
    {

        #region ----- Public Class -----

        public object VAT_REPORT_TYPE
        {
            get
            {
                return V_VAT_REPORT_TYPE.EditValue;
            }
        }

        public object VAT_REPORT_TYPE_DESC
        {
            get
            {
                return V_VAT_REPORT_TYPE_DESC.EditValue;
            }
        }

        public object MODIFY_DESC
        {
            get
            {
                return V_MODIFY_DESC.EditValue;
            }
        }

        public object VAT_LEVIER_TYPE 
        {
            get
            {
                return V_VAT_LEVIER_TYPE.EditValue;
            }
        }

        public object WRITE_DATE
        {
            get
            {
                return V_WRITE_DATE.EditValue;
            }
        }

        #endregion
        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        // 입력 암호 리턴.
        public object Get_Encrypt_Password
        {
            get
            {
                return V_VAT_REPORT_TYPE_DESC.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public FCMF0813_CREATE(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


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
            try
            {
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
            }
            catch
            {
            }
            return mPrompt;
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

        private void FCMF0813_CREATE_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0813_CREATE_Shown(object sender, EventArgs e)
        {
            V_MODIFY_DESC.ReadOnly = true;

            idcGetDate.ExecuteNonQuery();
            V_WRITE_DATE.EditValue = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

            //부가세 과세구분//
            IDC_GET_VAT_LEVIER_TYPE_P.ExecuteNonQuery();
            V_VAT_LEVIER_TYPE.EditValue = IDC_GET_VAT_LEVIER_TYPE_P.GetCommandParamValue("O_VAT_LEVIER_TYPE");
            V_VAT_LEVIER_TYPE_NAME.EditValue = IDC_GET_VAT_LEVIER_TYPE_P.GetCommandParamValue("O_VAT_LEVIER_TYPE_NAME");

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void BTN_CREATE_REPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //과세자구분
            if (iString.ISNull(V_VAT_LEVIER_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_VAT_LEVIER_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_VAT_LEVIER_TYPE_NAME.Focus();
                return;
            }
            //신고구분
            if (iString.ISNull(V_VAT_REPORT_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_VAT_REPORT_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_VAT_REPORT_TYPE_DESC.Focus();
                return;
            }
            if (iString.ISNull(V_VAT_REPORT_TYPE.EditValue) == "02" && iString.ISNull(V_MODIFY_DESC.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10493"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_MODIFY_DESC.Focus();
                return;
            }
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        
        #region ------ Lookup Event ------

        private void ILA_VAT_LEVIER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "VAT_LEVIER");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VAT_REPORT_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "VAT_REPORT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_VAT_MAKE_GB_SelectedRowData(object pSender)
        {
            if (iString.ISNull(V_VAT_REPORT_TYPE.EditValue) == "02")
            {//수정신고//
                V_MODIFY_DESC.ReadOnly = false;
            }
            else
            {
                V_MODIFY_DESC.EditValue = string.Empty;
                V_MODIFY_DESC.ReadOnly = true;
            }
        }

        #endregion



        #region ------ Adapter Event ------

        #endregion

    }
}