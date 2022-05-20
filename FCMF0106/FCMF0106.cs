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

namespace FCMF0106
{
    public partial class FCMF0106 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public FCMF0106(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property Method ------

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void SEARCH_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_ENTRY.TabIndex)
            {
                IDA_CREDIT_CARD_LENDING.Fill();
                IGR_CREDIT_CARD_LENDING.Focus();
            }
            else
            {
                if (iString.ISNull(W2_LENDING_DATE_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_LENDING_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                    W2_LENDING_DATE_FR.Focus();
                    return;
                }
                if (iString.ISNull(W2_LENDING_DATE_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_LENDING_DATE_TO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                    W2_LENDING_DATE_TO.Focus();
                    return;
                }

                IDA_CARD_LENDING_LIST.Fill();
                IGR_CARD_LENDING_LIST.Focus();
            }
        }

        private void Insert_Credit_Card()
        {
            LENDING_DATE_FR.EditValue = iDate.ISGetDate(DateTime.Today);
            LENDING_DATE_TO.EditValue = iDate.ISGetDate(DateTime.Today);

            CARD_NAME.Focus();
        }

        private void isSetCommonParameter(string pGroup_Code, string pEnabled_Flag)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag);
        }

        #endregion

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

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Button Click -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_CREDIT_CARD_LENDING.IsFocused)
                    {
                        IDA_CREDIT_CARD_LENDING.AddOver();
                        Insert_Credit_Card();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_CREDIT_CARD_LENDING.IsFocused)
                    {
                        IDA_CREDIT_CARD_LENDING.AddUnder();
                        Insert_Credit_Card();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                { 
                        IDA_CREDIT_CARD_LENDING.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CREDIT_CARD_LENDING.IsFocused)
                    {
                        IDA_CREDIT_CARD_LENDING.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CREDIT_CARD_LENDING.IsFocused)
                    {
                        IDA_CREDIT_CARD_LENDING.Delete();
                    }
                }
            }
        }

        #endregion
        
        #region ----- Form Event -----

        private void FCMF0106_Load(object sender, EventArgs e)
        {
            W1_LENDING_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today); 
            W1_LENDING_DATE_TO.EditValue = iDate.ISGetDate(DateTime.Today);

            W2_LENDING_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today); 
            W2_LENDING_DATE_TO.EditValue = iDate.ISGetDate(DateTime.Today);

            IDA_CREDIT_CARD_LENDING.FillSchema();
        }

        private void LENDING_DATE_FR_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            LENDING_DATE_TO.EditValue = LENDING_DATE_FR.EditValue;
        }
         
        #endregion

        #region ----- Lookup Code -----

        private void ILA_CREDIT_CARD_LENDING_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CREDIT_CARD_LENDING.SetLookupParamValue("W_RETURN_FLAG", "A");
        }

        private void ILA_CREDIT_CARD_LENDING_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CREDIT_CARD_LENDING.SetLookupParamValue("W_RETURN_FLAG", "Y");
        }

        private void ILA_LENDING_RETURN_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "LENDING_RETURN_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_CREDIT_CARD_LENDING_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["CARD_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Card Number(카드번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LENDING_PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Lending Person(대출자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LENDING_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("[Lending Date From] {0}", isMessageAdapter1.ReturnText("FCM_10010")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RETURN_TYPE"]) != string.Empty)
            {
                if (iString.ISNull(e.Row["RETURN_PERSON_ID"]) == string.Empty)
                {
                    MessageBoxAdv.Show(string.Format("[Return Person] {0}", isMessageAdapter1.ReturnText("FCM_10563")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                    e.Cancel = true;
                    return;
                }
            }
            if (iString.ISNull(e.Row["RETURN_PERSON_ID"]) != string.Empty)
            {
                if (iString.ISNull(e.Row["RETURN_TYPE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(string.Format("[Return Type] {0}", isMessageAdapter1.ReturnText("FCM_10563")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                    e.Cancel = true;
                    return;
                }
            }       
        }

        private void IDA_CREDIT_CARD_LENDING_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}