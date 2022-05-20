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

namespace FCMF0526
{
    public partial class FCMF0526_SET : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0526_SET(ISAppInterface pAppInterface, object pDUE_DATE_FR, object pDUE_DATE_TO, object pGL_DATE)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_DUE_DATE_FR.EditValue = pDUE_DATE_FR;
            V_DUE_DATE_TO.EditValue = pDUE_DATE_TO;
            V_GL_DATE.EditValue = pGL_DATE;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch
            {
                vDateTime = DateTime.Today;
            }
            return vDateTime;
        } 
        
        private Boolean CheckData()
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_FR.Focus();
                return false;
            }
            return true;
        }

        private void SEARCH_DB()
        {
            if (iString.ISNull(V_DUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_DUE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DUE_DATE_FR.Focus();
                return;
            }
            CHECK_YN.CheckBoxValue = "N";
            IDA_BILL_ACCOUNT.Fill();
            IGR_BILL_ACCOUNT.Focus();
        }

        private void Set_Grid_Control(object pCELL_STATUS)
        {
            int vIDX_CHECK = IGR_BILL_ACCOUNT.GetColumnToIndex("CHECK_YN");
            IGR_BILL_ACCOUNT.GridAdvExColElement[vIDX_CHECK].Insertable = pCELL_STATUS;
            IGR_BILL_ACCOUNT.GridAdvExColElement[vIDX_CHECK].Updatable = pCELL_STATUS;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
               pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }

            IGR_BILL_ACCOUNT.LastConfirmChanges();
            IDA_BILL_ACCOUNT.OraSelectData.AcceptChanges();
            IDA_BILL_ACCOUNT.Refillable = true;
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
        
        #region ----- From Event -----

        private void FCMF0526_SET_Load(object sender, EventArgs e)
        {
            IDA_BILL_ACCOUNT.FillSchema();
        }

        private void FCMF0526_SET_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        private void IGR_BALANCE_STATEMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BILL_ACCOUNT.GetColumnToIndex("CHECK_YN"))
            {
                IGR_BILL_ACCOUNT.LastConfirmChanges();
                IDA_BILL_ACCOUNT.OraSelectData.AcceptChanges();
                IDA_BILL_ACCOUNT.Refillable = true;
            }
        }
        
        private void isbtnSEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CheckData() == false)
            {
                return;
            }

            DialogResult vdlgResult;
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }
            
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int mError_Count = 0;

            int mIDX_CHECK_YN = IGR_BILL_ACCOUNT.GetColumnToIndex("CHECK_YN");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BILL_ACCOUNT.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_BILL_TYPE = IGR_BILL_ACCOUNT.GetColumnToIndex("BILL_TYPE");
            int mIDX_ERROR_YN = IGR_BILL_ACCOUNT.GetColumnToIndex("ERROR_YN");
            int mIDX_MESSAGE = IGR_BILL_ACCOUNT.GetColumnToIndex("MESSAGE");

            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_BILL_ACCOUNT.RowCount; c++)
            {
                if (iString.ISNull(IGR_BILL_ACCOUNT.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_BILL_ACCOUNT.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BILL_ACCOUNT.CurrentCellActivate(c, mIDX_CHECK_YN);

                    IDC_SET_BATCH_BILL_SELECTED.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BILL_ACCOUNT.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_SET_BATCH_BILL_SELECTED.SetCommandParamValue("P_BILL_TYPE", IGR_BILL_ACCOUNT.GetCellValue(c, mIDX_BILL_TYPE));
                    
                    IDC_SET_BATCH_BILL_SELECTED.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_SET_BATCH_BILL_SELECTED.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_SET_BATCH_BILL_SELECTED.GetCommandParamValue("O_MESSAGE"));

                    if (IDC_SET_BATCH_BILL_SELECTED.ExcuteError || mSTATUS == "F")
                    {
                        mSTATUS = "Y";
                        mError_Count = mError_Count + 1;
                    }
                    else
                    {
                        IGR_BILL_ACCOUNT.SetCellValue(c, mIDX_CHECK_YN, "N");
                        mSTATUS = "N";
                    }
                    IGR_BILL_ACCOUNT.SetCellValue(c, mIDX_ERROR_YN, mSTATUS);
                    IGR_BILL_ACCOUNT.SetCellValue(c, mIDX_MESSAGE, mMESSAGE);
                }
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_BILL_ACCOUNT.LastConfirmChanges();
            IDA_BILL_ACCOUNT.OraSelectData.AcceptChanges();
            IDA_BILL_ACCOUNT.Refillable = true;
            if (mError_Count > 0)
            {
                return;
            }
            else
            {
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(IGR_BILL_ACCOUNT, CHECK_YN.CheckBoxValue);
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("P_STD_DATE", GetDate());
        }
        
        private void ilaACCOUNT_CONTROL_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        #endregion

        #region ----- Adapter Event -----


        #endregion


    }
}