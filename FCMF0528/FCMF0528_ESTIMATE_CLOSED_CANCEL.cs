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

namespace FCMF0528
{
    public partial class FCMF0528_ESTIMATE_CLOSED_CANCEL : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0528_ESTIMATE_CLOSED_CANCEL(ISAppInterface pAppInterface, object pGL_DATE)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_BALANCE_DATE.EditValue = pGL_DATE;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private Boolean CheckData()
        {
            if (iConv.ISNull(V_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_BALANCE_DATE.Focus();
                return false;
            }
            return true;
        }

        private void SEARCH_DB()
        {
            if (iConv.ISNull(V_BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_BALANCE_DATE.Focus();
                return;
            }
            CHECK_YN.CheckBoxValue = "N";
            IDA_ACC_CURR_ESTIMATE.Fill();
            IGR_ACC_CURR_ESTIMATE.Focus();
        }

        private void Set_Grid_Control(object pCELL_STATUS)
        {
            int vIDX_CHECK = IGR_ACC_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN");
            IGR_ACC_CURR_ESTIMATE.GridAdvExColElement[vIDX_CHECK].Insertable = pCELL_STATUS;
            IGR_ACC_CURR_ESTIMATE.GridAdvExColElement[vIDX_CHECK].Updatable = pCELL_STATUS;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }

            IGR_ACC_CURR_ESTIMATE.LastConfirmChanges();
            IDA_ACC_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_ACC_CURR_ESTIMATE.Refillable = true;
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

        private void FCMF0528_ESTIMATE_CLOSED_CANCEL_Load(object sender, EventArgs e)
        {
            IDA_ACC_CURR_ESTIMATE.FillSchema();
        }

        private void FCMF0528_ESTIMATE_CLOSED_CANCEL_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        private void IGR_BALANCE_STATEMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_ACC_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN"))
            {
                IGR_ACC_CURR_ESTIMATE.LastConfirmChanges();
                IDA_ACC_CURR_ESTIMATE.OraSelectData.AcceptChanges();
                IDA_ACC_CURR_ESTIMATE.Refillable = true;
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

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int mError_Count = 0;

            int mIDX_CHECK_YN = IGR_ACC_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_ACC_CURR_ESTIMATE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_ERROR_YN = IGR_ACC_CURR_ESTIMATE.GetColumnToIndex("ERROR_YN");
            int mIDX_MESSAGE = IGR_ACC_CURR_ESTIMATE.GetColumnToIndex("MESSAGE");

            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_ACC_CURR_ESTIMATE.RowCount; c++)
            {
                if (iConv.ISNull(IGR_ACC_CURR_ESTIMATE.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_ACC_CURR_ESTIMATE.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_ACC_CURR_ESTIMATE.CurrentCellActivate(c, mIDX_CHECK_YN);

                    isDataTransaction1.BeginTran();
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_ACC_CURR_ESTIMATE.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.ExecuteNonQuery();
                    mSTATUS = iConv.ISNull(IDC_CANCEL_CLOSED_CURR_ESTIMATE.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iConv.ISNull(IDC_CANCEL_CLOSED_CURR_ESTIMATE.GetCommandParamValue("O_MESSAGE"));

                    if (IDC_CANCEL_CLOSED_CURR_ESTIMATE.ExcuteError || mSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        mSTATUS = "Y";
                        mError_Count = mError_Count + 1;
                    }
                    else
                    {
                        IGR_ACC_CURR_ESTIMATE.SetCellValue(c, mIDX_CHECK_YN, "N");
                        mSTATUS = "N";
                        isDataTransaction1.Commit();
                    }
                                        
                    IGR_ACC_CURR_ESTIMATE.SetCellValue(c, mIDX_ERROR_YN, mSTATUS);
                    IGR_ACC_CURR_ESTIMATE.SetCellValue(c, mIDX_MESSAGE, mMESSAGE);                                        
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_ACC_CURR_ESTIMATE.LastConfirmChanges();
            IDA_ACC_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_ACC_CURR_ESTIMATE.Refillable = true;
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
            Select_Check_YN(IGR_ACC_CURR_ESTIMATE, CHECK_YN.CheckBoxValue);
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        #endregion

        #region ----- Adapter Event -----
        
        private void IDA_BALANCE_STATEMENT_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            object mCELL_STATUS;
            mCELL_STATUS = "1";            

            Set_Grid_Control(mCELL_STATUS);
            IGR_ACC_CURR_ESTIMATE.LastConfirmChanges();
            IDA_ACC_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_ACC_CURR_ESTIMATE.Refillable = true;
        }
       
        #endregion


    }
}