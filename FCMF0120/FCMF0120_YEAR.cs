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

namespace FCMF0120
{
    public partial class FCMF0120_YEAR : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

         
        #endregion;

        #region ----- Constructor -----

        public FCMF0120_YEAR(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public FCMF0120_YEAR(ISAppInterface pAppInterface, object pFISCAL_CALENDAR_ID,
                                object pFISCAL_CALENDAR_CODE, object pFISCAL_CALENDAR_NAME)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            FISCAL_CALENDAR_ID.EditValue = pFISCAL_CALENDAR_ID;
            FISCAL_CALENDAR_CODE.EditValue = pFISCAL_CALENDAR_CODE;
            FISCAL_CALENDAR_NAME.EditValue = pFISCAL_CALENDAR_NAME;
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

        private void FCMF0120_YEAR_Load(object sender, EventArgs e)
        {
        }

        private void FCMF0120_YEAR_Shown(object sender, EventArgs e)
        {
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void FISCAL_YEAR_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            START_DATE.EditValue = iDate.ISGetDate(string.Format("{0}-01-01", FISCAL_YEAR.EditValue));
            END_DATE.EditValue = iDate.ISGetDate(string.Format("{0}-12-31", FISCAL_YEAR.EditValue));
        }

        private void BTN_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //회계년도 체크//
            if (iString.ISDecimaltoZero(FISCAL_COUNT.EditValue, 0) == 0)
            {// 회계기수
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10079"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(FISCAL_YEAR.EditValue) == string.Empty)
            {// 회계년도
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(FISCAL_YEAR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(START_DATE.EditValue) == string.Empty)
            {// 회계년도 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(START_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(END_DATE.EditValue) == string.Empty)
            {// 회계년도 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(END_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (Convert.ToDateTime(START_DATE.EditValue) > Convert.ToDateTime(END_DATE.EditValue))
            {// 회계년도 기간설정
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(YEAR_STATUS.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(YEAR_STATUS_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                return;
            }

            int mRecordCount = 0;
            string mStatus;
            string mMessage;
            idcRECORD_COUNT.ExecuteNonQuery();
            mRecordCount = iString.ISNumtoZero(idcRECORD_COUNT.GetCommandParamValue("O_RETURN_VALUE"));
            if (mRecordCount > 0)
            {
                if (DialogResult.Yes != MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                {
                    return;
                }
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CREATE_PERIOD.ExecuteNonQuery();
            mStatus = iString.ISNull(IDC_CREATE_PERIOD.GetCommandParamValue("O_STATUS"));
            mMessage = iString.ISNull(IDC_CREATE_PERIOD.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (mStatus == "F")
            {
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (mMessage != string.Empty)
            {
                MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        
        #region ------ Lookup Event ------

        #endregion

        #region ------ Adapter Event ------


        #endregion             

    }
}