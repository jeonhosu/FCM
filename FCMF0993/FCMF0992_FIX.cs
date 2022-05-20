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

namespace FCMF0992
{
    public partial class FCMF0992_FIX : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mSTATUS;
        object mTAX_BILL_ISSUE_NO;
        object mTAX_BILL_NO;
        object mHOMETAX_ISSUE_NO;

        #endregion;

        #region ----- Constructor -----

        public FCMF0992_FIX()
        {
            InitializeComponent();
        }

        public FCMF0992_FIX(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
          
        public FCMF0992_FIX(Form pMainForm, ISAppInterface pAppInterface, string pSTATUS,
                                object pTAX_BILL_ISSUE_NO, object pTAX_BILL_NO, object pHOMETAX_ISSUE_NO)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSTATUS = pSTATUS;
            mTAX_BILL_ISSUE_NO = pTAX_BILL_ISSUE_NO;
            mTAX_BILL_NO = pTAX_BILL_NO;
            mHOMETAX_ISSUE_NO = pHOMETAX_ISSUE_NO;
        }

        #endregion;

        #region ----- NEW ISSUE NO RETURN ------

        public string New_Tax_Bill_Issue_NO
        {
            get
            {
                return iConv.ISNull(V_NEW_TAX_BILL_ISSUE_NO.EditValue);
            } 
        }

        public string SRC_Tax_Bill_Issue_NO
        {
            get
            {
                return iConv.ISNull(SRC_TAX_BILL_ISSUE_NO.EditValue);
            }
        }

        public string SRC_Tax_Bill_NO
        {
            get
            {
                return iConv.ISNull(SRC_TAX_BILL_NO.EditValue);
            }
        }

        public string SRC_Hometax_Issue_NO
        {
            get
            {
                return iConv.ISNull(SRC_HOMETAX_ISSUE_NO.EditValue);
            }
        }

        #endregion

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

        #region ----- Form Evevnt -----

        private void FCMF0992_FIX_Load(object sender, EventArgs e)
        {
            SRC_TAX_BILL_ISSUE_NO.EditValue = mTAX_BILL_ISSUE_NO;
            SRC_TAX_BILL_NO.EditValue = mTAX_BILL_NO;
            SRC_HOMETAX_ISSUE_NO.EditValue = mHOMETAX_ISSUE_NO;
        }
             
        private void BTN_ISSUE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10182"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }            

            //현재화면 닫기//
            this.DialogResult = DialogResult.OK;
            this.Close(); 
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }
          
        #endregion

        #region ----- Lookup Event ------

        private void SetCommon(object pGROUP_CODE, object pENABLED_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pENABLED_YN);
        }
         
        private void ILA_TB_FIX_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_FIX_TYPE", "Y");   
        }

        #endregion

        #region ----- Adapter Event ------
         
        #endregion



    }
}