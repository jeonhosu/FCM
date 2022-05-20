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

namespace FCMF0092
{
    public partial class FCMF0092 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0092()
        {
            InitializeComponent();
        }

        public FCMF0092(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            IGR_APPROVAL_LINE.LastConfirmChanges();
            IDA_APPROVAL_LINE.OraSelectData.AcceptChanges();
            IDA_APPROVAL_LINE.Refillable = true;

            string vDEPT_CODE = iConv.ISNull(IGR_SLIP_DEPT.GetCellValue("DEPT_CODE"));
            int vIDX_DEPT_CODE = IGR_SLIP_DEPT.GetColumnToIndex("DEPT_CODE");

            IDA_SLIP_DEPT.Fill();
            for (int r = 0; r < IGR_SLIP_DEPT.RowCount; r++)
            {
                if (vDEPT_CODE == iConv.ISNull(IGR_SLIP_DEPT.GetCellValue(r, vIDX_DEPT_CODE)))
                {
                    IGR_SLIP_DEPT.CurrentCellMoveTo(r, vIDX_DEPT_CODE);
                    IGR_SLIP_DEPT.CurrentCellActivate(r, vIDX_DEPT_CODE);
                }
            }
            IGR_SLIP_DEPT.Focus();
        }

        private void SetCommonParameter(object pGroup_Code, object pENABLED_YN)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", pENABLED_YN);
        }

        private void Insert_Approval_Line()
        {
            IGR_APPROVAL_LINE.SetCellValue("APPROVAL_STEP_SEQ", 10);
            IGR_APPROVAL_LINE.SetCellValue("ENABLED_FLAG", "Y");
            IGR_APPROVAL_LINE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            IGR_APPROVAL_LINE.Focus();
        }


        private void Insert_Approval_Line_E()
        { 
            IGR_APPROVAL_LINE_E.SetCellValue("ENABLED_FLAG", "Y");
            IGR_APPROVAL_LINE_E.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            IGR_APPROVAL_LINE_E.Focus();
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
                    else if(IDA_APPROVAL_LINE_E.IsFocused)
                    {
                        IDA_APPROVAL_LINE_E.AddOver();
                        Insert_Approval_Line_E();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.AddUnder();
                        Insert_Approval_Line();
                    }
                    else if (IDA_APPROVAL_LINE_E.IsFocused)
                    {
                        IDA_APPROVAL_LINE_E.AddUnder();
                        Insert_Approval_Line_E();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_SLIP_DEPT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.Cancel();
                    }
                    else if (IDA_APPROVAL_LINE_E.IsFocused)
                    {
                        IDA_APPROVAL_LINE_E.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_APPROVAL_LINE.IsFocused)
                    {
                        IDA_APPROVAL_LINE.Delete();
                    }
                    else if(IDA_APPROVAL_LINE_E.IsFocused)
                    {
                        IDA_APPROVAL_LINE_E.Delete();
                    }
                }
            }
        }

        #endregion;  

        #region ----- Form Event -----

        private void FCMF0092_Load(object sender, EventArgs e)
        {
            IDA_SLIP_DEPT.FillSchema();
            IDA_APPROVAL_LINE.FillSchema();
        }

        private void FCMF0092_Shown(object sender, EventArgs e)
        {
            W_ENABLED_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        private void BTN_CHANGE_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        {

            FCMF0092_COPY vFCMF0092_COPY = new FCMF0092_COPY(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            DialogResult vResult = vFCMF0092_COPY.ShowDialog();
            if (vResult == DialogResult.OK)
            {
                SearchDB();
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaAPPROVE_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_APPROVAL_STEP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        private void ILA_SLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SLIP_TYPE.SetLookupParamValue("P_ENABLED_FLAG", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_APPROVAL_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["DEPT_ID"]) == string.Empty)
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

        private void IDA_APPROVAL_LINE_E_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_APPROVAL_LINE_E_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["SLIP_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Slip Type(전표유형)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        #endregion

    }
}