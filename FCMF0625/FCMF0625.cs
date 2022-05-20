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

namespace FCMF0625
{
    public partial class FCMF0625 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0625()
        {
            InitializeComponent();
        }

        public FCMF0625(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iConv.ISNull(W_BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_BUDGET_TYPE_NAME.Focus();
                return;
            }
            
            IDA_BUDGET_DEPT.Fill();
            IGR_BUDGET_DEPT.Focus();
        }

        private void SetCommonParameter(object pGroup_Code, object pENABLED_YN)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", pENABLED_YN);
        }

        private void SetCommonParameter_W(object pGroupCode, object pWhere, object pEnabled_YN)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }


        private void Insert_Approval_Line()
        {
            IGR_APPROVAL_LINE.SetCellValue("BUDGET_TYPE", W_BUDGET_TYPE.EditValue);
            IGR_APPROVAL_LINE.SetCellValue("BUDGET_TYPE_NAME", W_BUDGET_TYPE_NAME.EditValue);
            IGR_APPROVAL_LINE.SetCellValue("APPROVAL_STEP_SEQ", 0);
            IGR_APPROVAL_LINE.SetCellValue("ENABLED_FLAG", "Y");
            IGR_APPROVAL_LINE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
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
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.AddOver();
                        Insert_Approval_Line();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.AddUnder();
                        Insert_Approval_Line();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_BUDGET_DEPT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.Delete();
                    }
                }
            }
        }

        #endregion;  
        
        #region ----- Form Event -----

        private void FCMF0625_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_DEPT.FillSchema();
            IDA_APPROVAL_LINE.FillSchema();
        }

        private void FCMF0625_Shown(object sender, EventArgs e)
        {
            W_ENABLED_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_BUDGET_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BUDGET_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BUDGET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BUDGET_TYPE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BUDGET_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_BUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", DateTime.Today);
            ILD_BUDGET_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", DateTime.Today);            
        }

        private void ilaAPPROVE_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_APPROVAL_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["BUDGET_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10611"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10019"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["APPROVAL_STEP_SEQ"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Step Seq(승인순번)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["APPROVAL_STEP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Approval Step(승인단계)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Person Name(사원)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Effective Date From(적용시작일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
         
        private void IDA_APPROVAL_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

        private void B_CHANGE_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_BUDGET_TYPE.EditValue) == "")
            {
                MessageBox.Show("BUDGET_TYPE Is Empty!", "Warning");
                return;
            }
            if (iConv.ISNull(W_APPROVAL_STEP_CODE.EditValue) == "")
            {
                MessageBox.Show("APPROVAL_STEP Is Empty!", "Warning");
                return;
            }
            if (iConv.ISNumtoZero(W_PERSON_ID.EditValue) == 0)
            {
                MessageBox.Show("PERSON_ID_OLD Is Empty!", "Warning");
                return;
            }
            if (iConv.ISNumtoZero(W_PERSON_ID_NEW.EditValue) == 0)
            {
                MessageBox.Show("PERSON_ID_NEW Is Empty!", "Warning");
                return;
            }

            IDC_CHANGE_PERSON_ALL.ExecuteNonQuery();

            string L_RESULT;
            L_RESULT = iConv.ISNull(IDC_CHANGE_PERSON_ALL.GetCommandParamValue("W_RESULT"));
            MessageBox.Show(L_RESULT, "Information");

            // 사용자 조건 초기화
            W_PERSON_ID.EditValue = null;
            W_PERSON_NUM.EditValue = null;
            W_PERSON_NAME.EditValue = null;
            W_PERSON_ID_NEW.EditValue = null;
            W_PERSON_NUM_NEW.EditValue = null;
            W_PERSON_NAME_NEW.EditValue = null;

            // 재조회
            IDA_APPROVAL_LINE.Fill();
        }
    }
}