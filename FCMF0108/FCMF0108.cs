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

namespace FCMF0108
{
    public partial class FCMF0108 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public FCMF0108()
        {
            InitializeComponent();
        }

        public FCMF0108(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            idaACCOUNT_BOOK.Fill();
            ACCOUNT_BOOK_NAME.Focus();
        }

        private void Insert_Account_Book()
        {
            ACCOUNT_BOOK_CODE.Focus();
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
                    if (idaACCOUNT_BOOK.IsFocused)
                    {
                        idaACCOUNT_BOOK.AddOver();
                        Insert_Account_Book();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaACCOUNT_BOOK.IsFocused)
                    {
                        idaACCOUNT_BOOK.AddUnder();
                        Insert_Account_Book();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaACCOUNT_BOOK.IsFocused)
                    {
                        idaACCOUNT_BOOK.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaACCOUNT_BOOK.IsFocused)
                    {
                        idaACCOUNT_BOOK.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaACCOUNT_BOOK.IsFocused)
                    {
                        idaACCOUNT_BOOK.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0108_Load(object sender, EventArgs e)
        {
            BTN_PRE.BringToFront();
            BTN_NEXT.BringToFront();

            idaACCOUNT_BOOK.FillSchema();
        }

        private void BTN_PRE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaACCOUNT_BOOK.MovePrevious(ACCOUNT_BOOK_CODE.Name);
        }

        private void BTN_NEXT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaACCOUNT_BOOK.MoveNext(ACCOUNT_BOOK_CODE.Name);
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaACCOUNT_BOOK_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ACCOUNT_BOOK_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_BOOK_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_BOOK_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_BOOK_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["ACCOUNT_SET_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_SET_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["FISCAL_CALENDAR_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FISCAL_CALENDAR_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["FUTURE_OPEN_PERIOD_COUNT"] == DBNull.Value || Convert.ToInt32(e.Row["FUTURE_OPEN_PERIOD_COUNT"]) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FUTURE_OPEN_PERIOD_COUNT))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CURRENCY_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EXCHANGE_RATE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(EXCHANGE_RATE_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ESTIMATE_FORWARD_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ESTIMATE_FORWARD_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["OFFSET_EXCHANGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFSET_EXCHANGE_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ENABLED_FLAG"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10085"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(EFFECTIVE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaACCOUNT_BOOK_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaEXCHANGE_RATE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildEXCHANGE_RATE_TYPE.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ildEXCHANGE_RATE_TYPE.SetLookupParamValue("W_LOOKUP_TYPE", "EXCHANGE_RATE_TYPE");
            ildEXCHANGE_RATE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaEXCHANGE_RATE_APPLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildEXCHANGE_RATE_TYPE.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ildEXCHANGE_RATE_TYPE.SetLookupParamValue("W_LOOKUP_TYPE", "EXCHANGE_RATE_APPLY_TYPE");
            ildEXCHANGE_RATE_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY_CODE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VAT_LEVIER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "VAT_LEVIER");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBUDGET_CONTROL_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BUDGET_CONTROL_DEPT");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_GL_DATE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "GL_DATE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CHANGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_CHANGE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ESTIMATE_FORWARD_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ESTIMATE_FORWARD_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_OFFSET_EXCHANGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "OFFSET_EXCHANGE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

    }
}