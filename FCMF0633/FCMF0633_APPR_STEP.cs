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

namespace FCMF0633
{
    public partial class FCMF0633_APPR_STEP : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
         
        #endregion;

        #region ----- Constructor -----

        public FCMF0633_APPR_STEP(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public FCMF0633_APPR_STEP(ISAppInterface pAppInterface, object pBUDGET_MOVE_NUM, object pAPPROVAL_STEP_SEQ,
                                    object pBUDGET_PERIOD, object pBUDGET_MOVE_HEADER_ID,
                                    object pBUDGET_TYPE_NAME, object pBUDGET_TYPE, 
                                    object pBUDGET_DEPT_NAME, object pBUDGET_DEPT_CODE, object pBUDGET_DEPT_ID)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_BUDGET_MOVE_NUM.EditValue = pBUDGET_MOVE_NUM;
            V_APPROVAL_STEP_SEQ.EditValue = pAPPROVAL_STEP_SEQ;
            V_BUDGET_PERIOD.EditValue = pBUDGET_PERIOD;
            V_BUDGET_MOVE_HEADER_ID.EditValue = pBUDGET_MOVE_HEADER_ID;
            V_BUDGET_TYPE_NAME.EditValue = pBUDGET_TYPE_NAME;
            V_BUDGET_TYPE.EditValue = pBUDGET_TYPE;
            V_BUDGET_DEPT_NAME.EditValue = pBUDGET_DEPT_NAME;
            V_BUDGET_DEPT_CODE.EditValue = pBUDGET_DEPT_CODE;
            V_BUDGET_DEPT_ID.EditValue = pBUDGET_DEPT_ID;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            IDA_APPROVAL_PERSON.Fill();
        }

        private void Set_Grid_Item_Status(DataRow pDataRow)
        {
            int mIDX_APPR_FLAG = IGR_APPROVAL_PERSON.GetColumnToIndex("APPR_FLAG");
            int mIDX_PERSON_NAME = IGR_APPROVAL_PERSON.GetColumnToIndex("PERSON_NAME");
            int mIDX_EMAIL = IGR_APPROVAL_PERSON.GetColumnToIndex("EMAIL");
            int mIDX_ENABLED_FLAG = IGR_APPROVAL_PERSON.GetColumnToIndex("ENABLED_FLAG");

            if (pDataRow != null)
            {
                if (iString.ISNull(pDataRow["APPR_FLAG"]) == "Y".ToString())
                {
                    IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_PERSON_NAME].Updatable = 0;
                    IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_EMAIL].Updatable = 0;
                    IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_ENABLED_FLAG].Updatable = 0;
                }
                else
                {
                    IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_PERSON_NAME].Updatable = 1;
                    IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_EMAIL].Updatable = 1;
                    IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_ENABLED_FLAG].Updatable = 1;
                }
            }
            else
            {
                IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_PERSON_NAME].Updatable = 0;
                IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_EMAIL].Updatable = 0;
                IGR_APPROVAL_PERSON.GridAdvExColElement[mIDX_ENABLED_FLAG].Updatable = 0;
            }
            IGR_APPROVAL_PERSON.ResetDraw = true;
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

        private void FCMF0633_FILE_Load(object sender, EventArgs e)
        {
            IDA_APPROVAL_PERSON.FillSchema();
        }

        private void FCMF0633_FILE_Shown(object sender, EventArgs e)
        {
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;

            SearchDB();
        }

        private void BTN_INQUIRY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.Fill();
        }

        private void BTN_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.AddUnder();

            IGR_APPROVAL_PERSON.SetCellValue("BUDGET_TYPE", V_BUDGET_TYPE.EditValue);
            IGR_APPROVAL_PERSON.SetCellValue("BUDGET_HEADER_ID", V_BUDGET_MOVE_HEADER_ID.EditValue);
            IGR_APPROVAL_PERSON.SetCellValue("BUDGET_DEPT_ID", V_BUDGET_DEPT_ID.EditValue);
            IGR_APPROVAL_PERSON.Focus();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.Cancel();
        }

        private void BTN_UPDATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_APPROVAL_PERSON.Update();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        
        #region ------ Lookup Event ------

        private void ILA_APPROVAL_STEP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_PERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        #endregion

        private void IDA_APPROVAL_PERSON_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Item_Status(pBindingManager.DataRow);
        }

        private void IDA_APPROVAL_PERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Budget Type(예산구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BUDGET_HEADER_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Budget Req. Num(예산신청번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10019"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["APPROVAL_STEP_SEQ"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Step Seq(승인순번)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["APPROVAL_STEP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Approval Step(승인단계)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Person Name(사원)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        #region ------ Adapter Event ------


        #endregion             

    }
}