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

namespace FCMF0310
{
    public partial class FCMF0310 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0310()
        {
            InitializeComponent();
        }

        public FCMF0310(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void SearchDB()
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10300"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_DPR_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10221"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_DPR_TYPE_NAME.Focus();
                return;
            }
                        

            if (TB_MAIN.SelectedTab.TabIndex == TP_CREATE_SLIP.TabIndex)
            {
                if (IGR_DPR_SLIP_DETAIL.RowCount > 0)
                {
                    IGR_DEPRECIATION_HISTORY.LastConfirmChanges();
                    IDA_DPR_HISTORY.OraSelectData.AcceptChanges();
                    IDA_DPR_HISTORY.Refillable = true;
                }

                IDA_DPR_HISTORY.Fill();
                IGR_DPR_SLIP_DETAIL.Focus();

                ///////////////////////////////////////////////////////////
                //자산 배부율 미등록 자료 체크 //
                IDC_CHECK_DPR_ACCOUNT_P.ExecuteNonQuery();
                string mSTATUS = iString.ISNull(IDC_CHECK_DPR_ACCOUNT_P.GetCommandParamValue("O_STATUS"));
                string mMESSAGE = iString.ISNull(IDC_CHECK_DPR_ACCOUNT_P.GetCommandParamValue("O_MESSAGE"));
                if (mSTATUS == "F")
                {
                    if (mMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (IDC_CHECK_DPR_ACCOUNT_P.ExcuteError)
                {
                    MessageBoxAdv.Show(IDC_CHECK_DPR_ACCOUNT_P.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SLIP_LIST.TabIndex)
            {
                IDA_DPR_SLIP_SUM.Fill();
                IGR_DPR_SLIP_SUM.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_MINUS_DPR_AMOUNT.TabIndex)
            {
                IDA_DPR_AMOUNT_MINUS.Fill();
                IGR_DPR_AMOUNT_MINUS.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_CHK_DIST_RATE.TabIndex)
            {
                IDA_DPR_DIST_ACCOUNT.Fill();
                IGR_DPR_DIST_ACCOUNT.Focus();
            }
        }

        private void SetCommonParameter_W(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_WHERE", "VALUE1 = 'Y'");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Depreciation_History_Check_YN()
        {
            string vCHECK_FLAG = CHECK_YN_1.CheckBoxString;
            int vIDX_CHECK = IGR_DEPRECIATION_HISTORY.GetColumnToIndex("SELECT_YN");
            for (int i = 0; i < IGR_DEPRECIATION_HISTORY.RowCount; i++)
            {
                IGR_DEPRECIATION_HISTORY.SetCellValue(i, vIDX_CHECK, vCHECK_FLAG);
            }

            IGR_DEPRECIATION_HISTORY.LastConfirmChanges();
            IDA_DPR_HISTORY.OraSelectData.AcceptChanges();
            IDA_DPR_HISTORY.Refillable = true;
        }

        private void SET_GRIDE_CELL_STATE(DataRow pDataRow)
        {
            //bool mREAD_ONLY_YN = true;
            //int mINSERT_YN = 0;
            //int mUPDATE_YN = 0;
            //int mIDX_COL = igrDPR_SLIP.GetColumnToIndex("CHECK_YN");

            //if (pDataRow == null || iString.ISNull(pDataRow["SLIP_YN"]) == "N")
            //{
            //    mREAD_ONLY_YN = false;
            //    mINSERT_YN = 1;
            //    mUPDATE_YN = 1;
            //}
            //else
            //{
            //    mREAD_ONLY_YN = true;
            //    mINSERT_YN = 0;
            //    mUPDATE_YN = 0;
            //}
            //igrDPR_SLIP.GridAdvExColElement[mIDX_COL].ReadOnly = mREAD_ONLY_YN;
            //igrDPR_SLIP.GridAdvExColElement[mIDX_COL].Insertable = mINSERT_YN;
            //igrDPR_SLIP.GridAdvExColElement[mIDX_COL].Updatable = mUPDATE_YN;

            //igrDPR_SLIP.ResetDraw = true;
        }

        private void Show_Slip_Detail()
        {
            //int mSLIP_HEADER_ID = iString.ISNumtoZero(igrDPR_SLIP.GetCellValue("SLIP_HEADER_ID"));
            //if (mSLIP_HEADER_ID != Convert.ToInt32(0))
            //{
            //    Application.UseWaitCursor = true;
            //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            //    FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
            //    vFCMF0204.Show();

            //    this.Cursor = System.Windows.Forms.Cursors.Default;
            //    Application.UseWaitCursor = false;
            //} 
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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {                 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DPR_HISTORY.IsFocused)
                    {
                        IDA_DPR_HISTORY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DPR_HISTORY.IsFocused)
                    {
                        IDA_DPR_HISTORY.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0310_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0310_Shown(object sender, EventArgs e)
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);

            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "DPR_TYPE");
            idcDV_COMMON.ExecuteNonQuery();
            W_DPR_TYPE.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");
            W_DPR_TYPE_NAME.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");

        }

        private void CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Depreciation_History_Check_YN();
        }

        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10300"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_DPR_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10221"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_DPR_TYPE_NAME.Focus();
                return;
            }
            
            //전표생성여부 묻기.
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10303"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            isDataTransaction1.BeginTran();

            // 처리대상선택//
            int mCHECK_COUNT = 0;
            int mIDX_SELECT_YN = IGR_DEPRECIATION_HISTORY.GetColumnToIndex("SELECT_YN");
            int mIDX_ASSET_ID = IGR_DEPRECIATION_HISTORY.GetColumnToIndex("ASSET_ID");
            int mIDX_PERIOD_NAME = IGR_DEPRECIATION_HISTORY.GetColumnToIndex("PERIOD_NAME");
            int mIDX_DPR_TYPE = IGR_DEPRECIATION_HISTORY.GetColumnToIndex("DPR_TYPE");

            string mSTATUS = string.Empty;
            string mMESSAGE = string.Empty;

            for (int nRow = 0; nRow < IGR_DEPRECIATION_HISTORY.RowCount; nRow++)
            {
                if (iString.ISNull(IGR_DEPRECIATION_HISTORY.GetCellValue(nRow, mIDX_SELECT_YN)) == "Y")
                {
                    mCHECK_COUNT = mCHECK_COUNT + 1;

                    // 대상 UPDATE //
                    IDC_UPDATE_DPR_HISTORY_LIST.SetCommandParamValue("W_ASSET_ID", IGR_DEPRECIATION_HISTORY.GetCellValue(nRow, mIDX_ASSET_ID));
                    IDC_UPDATE_DPR_HISTORY_LIST.SetCommandParamValue("W_PERIOD_NAME", IGR_DEPRECIATION_HISTORY.GetCellValue(nRow, mIDX_PERIOD_NAME));
                    IDC_UPDATE_DPR_HISTORY_LIST.SetCommandParamValue("W_DPR_TYPE", IGR_DEPRECIATION_HISTORY.GetCellValue(nRow, mIDX_DPR_TYPE));
                    IDC_UPDATE_DPR_HISTORY_LIST.ExecuteNonQuery();
                    mSTATUS = IDC_UPDATE_DPR_HISTORY_LIST.GetCommandParamValue("O_STATUS").ToString();
                    if (IDC_UPDATE_DPR_HISTORY_LIST.ExcuteError || mSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        mMESSAGE = iString.ISNull(IDC_UPDATE_DPR_HISTORY_LIST.GetCommandParamValue("O_MESSAGE"));
                        if (mMESSAGE != string.Empty)
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();

                            MessageBoxAdv.Show(mMESSAGE, "1.Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }

            //선택된 처리대상에 대해 전표 전송//
            IDC_CREATE_DPR_SLIP.ExecuteNonQuery();
            mSTATUS = IDC_CREATE_DPR_SLIP.GetCommandParamValue("O_STATUS").ToString();
            if (IDC_CREATE_DPR_SLIP.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                mMESSAGE = iString.ISNull(IDC_CREATE_DPR_SLIP.GetCommandParamValue("O_MESSAGE"));
                if (mMESSAGE != string.Empty)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(mMESSAGE, "2.Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            isDataTransaction1.Commit();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            SearchDB();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(IGR_DPR_SLIP_SUM.GetCellValue("PERIOD_NAME")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10300"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);               
                return;
            }
            if (iString.ISNull(IGR_DPR_SLIP_SUM.GetCellValue("DPR_TYPE")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10221"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(IGR_DPR_SLIP_SUM.GetCellValue("AST_CATEGORY_ID")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10095"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(IGR_DPR_SLIP_SUM.GetCellValue("HEADER_INTERFACE_ID")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10335"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 

            //전표삭제여부 묻기.
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10333"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = string.Empty;
            string mMESSAGE = string.Empty;

            //선택된 처리대상에 대해 전표 전송//
            IDA_CANCEL_SLIP.ExecuteNonQuery();
            mSTATUS = IDA_CANCEL_SLIP.GetCommandParamValue("O_STATUS").ToString();
            if (IDA_CANCEL_SLIP.ExcuteError || mSTATUS == "F")
            {
                mMESSAGE = iString.ISNull(IDA_CANCEL_SLIP.GetCommandParamValue("O_MESSAGE"));
                if (mMESSAGE != string.Empty)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(mMESSAGE, "2.Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            SearchDB();
        }

        private void IGR_DEPRECIATION_HISTORY_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_SELECT_YN = IGR_DEPRECIATION_HISTORY.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_SELECT_YN)
            {
                IGR_DEPRECIATION_HISTORY.LastConfirmChanges();
                IDA_DPR_HISTORY.OraSelectData.AcceptChanges();
                IDA_DPR_HISTORY.Refillable = true;
            }
        }

        private void igrDPR_SLIP_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        #endregion

        #region ----- Lookup Event ------

        private void ilaPERIOD_NAME_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD.SetLookupParamValue("W_START_YYYYMM", null);
            ildPERIOD.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today,1));
        }

        private void ilaASSET_CATEGORY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDPR_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("DPR_TYPE", "N");
        }

        #endregion

        #region ----- Adapter : DPR_SLIP ------

        private void idaDPR_SLIP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["CHECK_YN"]) != "Y")
            {
                return;
            }
            if (iString.ISNull(e.Row["ASSET_CATEGORY_ID"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10095"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DPR_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10097"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EXPENSE_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10220"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["COST_CENTER_ID"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10302"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDPR_SLIP_UpdateCompleted(object pSender)
        {
            Application.DoEvents();
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            string mMESSAGE;
            IDC_CREATE_DPR_SLIP.ExecuteNonQuery();
            mMESSAGE = iString.ISNull(IDC_CREATE_DPR_SLIP.GetCommandParamValue("O_MESSAGE"));
            Application.DoEvents();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;

            if (mMESSAGE != String.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            // RE-QUERY.
            SearchDB();
        }

        private void idaDPR_SLIP_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            SET_GRIDE_CELL_STATE(pBindingManager.DataRow);
        }

        #endregion

    }
}