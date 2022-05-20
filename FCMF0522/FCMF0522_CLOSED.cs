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

namespace FCMF0522
{
    public partial class FCMF0522_CLOSED : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0522_CLOSED(ISAppInterface pAppInterface, object pGL_DATE)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            BALANCE_DATE.EditValue = pGL_DATE;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private Boolean CheckData()
        {
            //if (iString.ISNull(LAST_CLOSED_DATE.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    GL_DATE.Focus();
            //    return false;
            //}
            if (iString.ISNull(BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BALANCE_DATE.Focus();
                return false;
            }
            //if (Convert.ToDateTime(LAST_CLOSED_DATE.EditValue) > Convert.ToDateTime(GL_DATE.EditValue))
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    GL_DATE.Focus();
            //    return false;
            //}
            return true;
        }

        private void SEARCH_DB()
        {
            if (iString.ISNull(BALANCE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BALANCE_DATE.Focus();
                return;
            }
            CHECK_YN.CheckBoxValue = "N";
            IDA_BALANCE_STATEMENT.Fill();
            IGR_BALANCE_STATEMENT.Focus();
        }

        private void Set_Grid_Control(object pCELL_STATUS)
        {
            int vIDX_CHECK = IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN");
            IGR_BALANCE_STATEMENT.GridAdvExColElement[vIDX_CHECK].Insertable = pCELL_STATUS;
            IGR_BALANCE_STATEMENT.GridAdvExColElement[vIDX_CHECK].Updatable = pCELL_STATUS;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            int vIDX_CLOSED_DATE = pGrid.GetColumnToIndex("CLOSED_DATE");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                if (iDate.ISGetDate(BALANCE_DATE.EditValue) <= iDate.ISGetDate(pGrid.GetCellValue(i, vIDX_CLOSED_DATE)))
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, "N");
                }
                else
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
                }
            }

            IGR_BALANCE_STATEMENT.LastConfirmChanges();
            IDA_BALANCE_STATEMENT.OraSelectData.AcceptChanges();
            IDA_BALANCE_STATEMENT.Refillable = true;
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

        private void FCMF0522_CLOSED_Load(object sender, EventArgs e)
        {
            IDA_BALANCE_STATEMENT.FillSchema();
        }

        private void FCMF0522_CLOSED_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        private void btnSEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
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
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10383"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.DoEvents();
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            int mError_Count = 0;

            int mIDX_CHECK_YN = IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_ERROR_YN = IGR_BALANCE_STATEMENT.GetColumnToIndex("ERROR_YN");
            int mIDX_MESSAGE = IGR_BALANCE_STATEMENT.GetColumnToIndex("MESSAGE");

            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_BALANCE_STATEMENT.RowCount; c++)
            {
                if (iString.ISNull(IGR_BALANCE_STATEMENT.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_BALANCE_STATEMENT.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BALANCE_STATEMENT.CurrentCellActivate(c, mIDX_CHECK_YN);

                    isDataTransaction1.BeginTran();
                    idcBALANCE_CLOSED.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BALANCE_STATEMENT.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    idcBALANCE_CLOSED.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(idcBALANCE_CLOSED.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(idcBALANCE_CLOSED.GetCommandParamValue("O_MESSAGE"));

                    if (idcBALANCE_CLOSED.ExcuteError || mSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        mSTATUS = "Y";
                        mError_Count = mError_Count + 1;
                    }
                    else
                    {
                        IGR_BALANCE_STATEMENT.SetCellValue(c, mIDX_CHECK_YN, "N");
                        mSTATUS = "N";
                        isDataTransaction1.Commit();
                    }

                    IGR_BALANCE_STATEMENT.SetCellValue(c, mIDX_ERROR_YN, mSTATUS);
                    IGR_BALANCE_STATEMENT.SetCellValue(c, mIDX_MESSAGE, mMESSAGE);
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_BALANCE_STATEMENT.LastConfirmChanges();
            IDA_BALANCE_STATEMENT.OraSelectData.AcceptChanges();
            IDA_BALANCE_STATEMENT.Refillable = true;
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

        private void IGR_BALANCE_STATEMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN"))
            {
                IGR_BALANCE_STATEMENT.LastConfirmChanges();
                IDA_BALANCE_STATEMENT.OraSelectData.AcceptChanges();
                IDA_BALANCE_STATEMENT.Refillable = true;
            }
        }

        private void CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(IGR_BALANCE_STATEMENT, CHECK_YN.CheckBoxValue);
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
            if (iDate.ISGetDate(BALANCE_DATE.EditValue) <= iDate.ISGetDate(pBindingManager.DataRow["CLOSED_DATE"]))
            {
                IGR_BALANCE_STATEMENT.SetCellValue("CHECK_YN", "N");
                mCELL_STATUS = "0";
            }
            else
            {
                mCELL_STATUS = "1";
            }

            Set_Grid_Control(mCELL_STATUS);
            IGR_BALANCE_STATEMENT.LastConfirmChanges();
            IDA_BALANCE_STATEMENT.OraSelectData.AcceptChanges();
            IDA_BALANCE_STATEMENT.Refillable = true;
        }

        #endregion

    }
}