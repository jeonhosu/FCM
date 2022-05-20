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

namespace FCMF0113
{
    public partial class FCMF0113 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0113()
        {
            InitializeComponent();
        }

        public FCMF0113(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iString.ISNull(W_FS_FORM_TYPE_ID.EditValue) == string.Empty)               
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10156"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_FS_FORM_TYPE_NAME.Focus();
                return;
            }

            int vIDX_FORM_HEADER_ID = IGR_FORM_HEADER.GetColumnToIndex("FORM_HEADER_ID");
            int vIDX_FORM_ITEM_CODE = igrFORM.GetColumnToIndex("FORM_ITEM_CODE");
            int vIDX_ACCOUNT_CODE = igrMISS_ACCOUNT.GetColumnToIndex("ACCOUNT_CODE");

            int vFORM_HEADER_ID = iString.ISNumtoZero(FORM_HEADER_ID.EditValue);
            string vFORM_ITEM_CODE = iString.ISNull(igrFORM.GetCellValue("FORM_ITEM_CODE"));
            string vACCOUNT_CODE = iString.ISNull(igrMISS_ACCOUNT.GetCellValue("ACCOUNT_CODE"));
            
            IDA_FS_FORM_TYPE.Fill(); 
            IDA_FORM_LIST.Fill();
            IDA_MISS_ACCOUNT.Fill();

            // 기존 위치로 이동 //
            if (IGR_FORM_HEADER.RowCount > 0 && vFORM_HEADER_ID != 0)
            {
                for (int vRow = 0; vRow < IGR_FORM_HEADER.RowCount; vRow++)
                {
                    if (vFORM_HEADER_ID == iString.ISNumtoZero(IGR_FORM_HEADER.GetCellValue(vRow, vIDX_FORM_HEADER_ID)))
                    {
                        IGR_FORM_HEADER.CurrentCellMoveTo(vRow, vIDX_FORM_HEADER_ID);
                    }
                }
            }
            if (igrFORM.RowCount > 0 && vFORM_ITEM_CODE != string.Empty)
            {
                for (int vRow = 0; vRow < igrFORM.RowCount; vRow++)
                {
                    if(vFORM_ITEM_CODE == iString.ISNull(igrFORM.GetCellValue(vRow, vIDX_FORM_ITEM_CODE)))
                    {
                        igrFORM.CurrentCellMoveTo(vRow, vIDX_FORM_ITEM_CODE);
                    }
                }
            }

            if (igrMISS_ACCOUNT.RowCount > 0 && vACCOUNT_CODE != string.Empty)
            {
                for (int vRow = 0; vRow < igrMISS_ACCOUNT.RowCount; vRow++)
                {
                    if (vACCOUNT_CODE == iString.ISNull(igrMISS_ACCOUNT.GetCellValue(vRow, vIDX_ACCOUNT_CODE)))
                    {
                        igrMISS_ACCOUNT.CurrentCellMoveTo(vRow, vIDX_ACCOUNT_CODE);
                    }
                }
            }

            if (itbFORM.SelectedTab.TabIndex == 1)
            {   
                //Init_Form_Line();
                IGR_FORM_HEADER.Focus();
            }
            else if (itbFORM.SelectedTab.TabIndex == 2)
            {                
                igrFORM.Focus();
            }
            else if (itbFORM.SelectedTab.TabIndex == 3)
            {                
                igrMISS_ACCOUNT.Focus();
            }
        }

        private void Insert_Form_Type()
        {
            H_FS_FORM_LEVEL.EditValue = 0;
            H_EFFECTIVE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            H_ENABLED_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;

            H_FS_FORM_TYPE.Focus();
        }

        private void Insert_Header()
        {            
            DISPLAY_YN.CheckBoxValue = "Y";
            AMOUNT_PRINT_YN.CheckBoxValue = "Y";
            FORM_FRAME_YN.CheckBoxValue = "Y";
            ENABLED_FLAG.CheckBoxValue = "Y";
            EFFECTIVE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);

            FORM_ITEM_CODE.Focus();
        }

        private void Insert_Line()
        {
            if (iString.ISNull(ITEM_LEVEL_NAME.EditValue) == string.Empty)
            {
                IDA_FORM_LINE.Delete();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10160"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ITEM_LEVEL_NAME.Focus();
                return;
            }
            IGR_FORM_LINE.SetCellValue("ACCOUNT_DR_CR", ACCOUNT_DR_CR.EditValue);
            IGR_FORM_LINE.SetCellValue("ACCOUNT_DR_CR_NAME", ACCOUNT_DR_CR_NAME.EditValue);
            IGR_FORM_LINE.SetCellValue("ENABLED_FLAG", "Y");  
            IGR_FORM_LINE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
        }
      
        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Copy_FINANCIAL_STATEMENTS_FORM(object pFS_Form_Type_ID, object pNew_FS_Form_Type_ID)
        {
            if (iString.ISNull(pFS_Form_Type_ID) == string.Empty)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            decimal vRecord_Count = 0;
            IDC_FORM_COUNT.SetCommandParamValue("P_FS_FORM_TYPE_ID", pNew_FS_Form_Type_ID);
            IDC_FORM_COUNT.ExecuteNonQuery();
            vRecord_Count = iString.ISDecimaltoZero(IDC_FORM_COUNT.GetCommandParamValue("O_RECORD_COUNT"));
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (vRecord_Count != 0)
            {
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    return;
                }
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = String.Empty;
            string vMESSAGE = string.Empty;


            IDC_COPY_FORM.SetCommandParamValue("P_FS_FORM_TYPE_ID", pFS_Form_Type_ID);
            IDC_COPY_FORM.SetCommandParamValue("P_NEW_FS_FORM_TYPE_ID", pNew_FS_Form_Type_ID);
            IDC_COPY_FORM.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_COPY_FORM.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_COPY_FORM.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_COPY_FORM.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE == string.Empty)
                {
                    return;
                }
                MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion;

        #region ----- Item Level - Initialization -----

        //private void Set_Last_Level()
        //{
        //    int mForm_Type_Level = iString.ISNumtoZero(H_FS_FORM_LEVEL.EditValue);
        //    int mItem_Level = iString.ISNumtoZero(ITEM_LEVEL_NAME.EditValue);
        //    if (mForm_Type_Level == mItem_Level)
        //    {// 최종 레벨.
        //        LAST_LEVEL_YN.EditValue = "Y";
        //    }
        //    else
        //    {
        //        LAST_LEVEL_YN.EditValue = "N";
        //    }
        //    Init_Form_Line();
        //}

        //private void Init_Form_Line()
        //{
        //    if (iString.ISNull(LAST_LEVEL_YN.EditValue, "N") == "Y".ToString())
        //    {// 최종레벨 --> 계정과목 표시.
        //        IGR_FORM_LINE.GridAdvExColElement[IGR_FORM_LINE.GetColumnToIndex("JOIN_CONTROL_NAME")].LookupAdapter = ilaACCOUNT_CONTROL;
        //    }
        //    else
        //    {
        //        IGR_FORM_LINE.GridAdvExColElement[IGR_FORM_LINE.GetColumnToIndex("JOIN_CONTROL_NAME")].LookupAdapter = ILA_FORM_ITEM_LEVEL;
        //    }
        //}

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
                    IGR_FORM_HEADER.ResetDraw = true;
                    IGR_FORM_LINE.ResetDraw = true;

                    if (IDA_FS_FORM_TYPE.IsFocused)
                    {
                        IDA_FS_FORM_TYPE.AddOver();
                        Insert_Form_Type();
                    }
                    else if (IDA_FORM_LINE.IsFocused)
                    {
                        IDA_FORM_LINE.AddOver();
                        Insert_Line();
                    }
                    else if (IDA_FORM_HEADER.IsFocused)
                    {
                        IDA_FORM_HEADER.AddOver();
                        Insert_Header();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IGR_FORM_HEADER.ResetDraw = true;
                    IGR_FORM_LINE.ResetDraw = true;

                    if (IDA_FS_FORM_TYPE.IsFocused)
                    {
                        IDA_FS_FORM_TYPE.AddUnder();
                        Insert_Form_Type();
                    }
                    else if (IDA_FORM_LINE.IsFocused)
                    {
                        IDA_FORM_LINE.AddUnder();
                        Insert_Line();
                    }
                    else if (IDA_FORM_HEADER.IsFocused)
                    {
                        IDA_FORM_HEADER.AddUnder();
                        Insert_Header();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_FS_FORM_TYPE.Update();
                    }
                    catch(Exception Ex)
                    {
                        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_FS_FORM_TYPE.IsFocused)
                    {
                        IDA_FORM_LINE.Cancel();
                        IDA_FORM_HEADER.Cancel();
                        IDA_FS_FORM_TYPE.Cancel();
                    }
                    else if (IDA_FORM_LINE.IsFocused)
                    {
                        IDA_FORM_LINE.Cancel();
                    }
                    else if (IDA_FORM_HEADER.IsFocused)
                    {
                        IDA_FORM_LINE.Cancel();
                        IDA_FORM_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_FS_FORM_TYPE.IsFocused)
                    {
                        if (IDA_FS_FORM_TYPE.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_FS_FORM_TYPE.Delete();
                        }
                        else
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10307"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                    }
                    else if (IDA_FORM_LINE.IsFocused)
                    {
                        try
                        {
                            IDA_FORM_LINE.Delete();
                        }
                        catch
                        {

                        }
                    }
                    else if (IDA_FORM_HEADER.IsFocused)
                    {
                        try
                        {
                            IDA_FORM_HEADER.Delete();
                        }
                        catch 
                        {

                        }
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0113_Load(object sender, EventArgs e)
        {
            IDA_FS_FORM_TYPE.FillSchema();
            IDA_FORM_HEADER.FillSchema();
            IDA_FORM_LINE.FillSchema();
        }

        private void ITEM_LEVEL_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (iString.ISNumtoZero(H_FS_FORM_LEVEL.EditValue) < iString.ISNumtoZero(ITEM_LEVEL.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10133"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ITEM_LEVEL_NAME.Focus();
            }
        }

        private void ITEM_LEVEL_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {            
            //Set_Last_Level();
        }

        private void BTN_FS_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            FCMF0113_COPY vFCMF0113_COPY = new FCMF0113_COPY(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                            , W_FS_FORM_TYPE_ID.EditValue, W_FS_FORM_TYPE_NAME.EditValue);
            vdlgResult = vFCMF0113_COPY.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            {
                object vFS_Form_Type_ID = vFCMF0113_COPY.Get_FS_Form_Type_ID;
                object vNew_FS_Form_Type_ID = vFCMF0113_COPY.Get_New_FS_Form_Type_ID;
                
                Copy_FINANCIAL_STATEMENTS_FORM(vFS_Form_Type_ID, vNew_FS_Form_Type_ID);
            }
            vFCMF0113_COPY.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaFORM_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_FORM_TYPE", "Y");
        }

        private void ILA_FORM_ITEM_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FORM_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaFORM_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_FORM_TYPE", "Y");
        }
         
        private void ILA_ITEM_LEVEL_CODE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_ITEM_LEVEL_CODE.SetLookupParamValue("W_ITEM_LEVEL", H_FS_FORM_LEVEL.EditValue);
        }
         
        private void ILA_FS_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_TYPE", "Y");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ILA_PRT_POSITION_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("PRT_POSITION", "Y");
        }
         
        private void ilaFORM_ITEM_LEVEL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FORM_ITEM_LEVEL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_FS_ITEM_PROPERTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_PROPERTY", "Y");
        }

        private void ilaFORM_ITEM_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_ITEM_TYPE", "Y");
        }

        private void ilaFORM_ITEM_CLASS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_ITEM_CLASS", "Y");
        }

        private void ILA_REF_FORM_ITEM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_REF_FORM_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CALC_SIGN_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CALC_SIGN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_FORM_TYPE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["FS_FORM_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(H_FS_FORM_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FS_FORM_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(H_FS_FORM_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FS_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(H_FS_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FS_FORM_LEVEL"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(H_FS_FORM_LEVEL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaFORM_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["FS_FORM_TYPE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10156"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FORM_ITEM_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10157"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FORM_ITEM_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10158"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SORT_SEQ"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10159"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ITEM_LEVEL_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ITEM_LEVEL_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["ITEM_LEVEL"], 0) == Convert.ToInt32(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10160"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(H_FS_FORM_LEVEL.EditValue, 0) < iString.ISNumtoZero(e.Row["ITEM_LEVEL"], 0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10133"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                ITEM_LEVEL_NAME.Focus();
            }
            if (iString.ISNull(e.Row["PRT_POSITION_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10162"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LAST_LEVEL_YN"]) == "N" && iString.ISNull(e.Row["REF_FS_FORM_TYPE_ID"]) != string.Empty)
            {
                //최종자료가 아니면 입력 불가//
                MessageBoxAdv.Show(string.Format("Ref.F/S Form type : {0}", isMessageAdapter1.ReturnText("FCM_10326")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REF_FS_FORM_TYPE_ID"]) == string.Empty && iString.ISNull(e.Row["REF_FORM_HEADER_ID"]) != string.Empty)
            {
                //관련 재무제표 양식은 선택하지 않고 관련 항목은 선택하지 않음
                MessageBoxAdv.Show(string.Format("Ref.F/S Form type : {0}", isMessageAdapter1.ReturnText("FCM_10577")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REF_FS_FORM_TYPE_ID"]) != string.Empty && iString.ISNull(e.Row["REF_FORM_HEADER_ID"]) == string.Empty)
            {
                //관련 재무제표 양식은 선택했으나 관련 항목은 선택함
                MessageBoxAdv.Show(string.Format("Ref.F/S Form type : {0}", isMessageAdapter1.ReturnText("FCM_10578")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaFORM_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
        }

        private void IDA_FORM_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["JOIN_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10163"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CALC_SIGN_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10164"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
         
        #endregion        

 
    }
}