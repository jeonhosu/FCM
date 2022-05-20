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

namespace FCMF0628
{
    public partial class FCMF0628_SLIP : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mSession_ID;
        InfoSummit.Win.ControlAdv.ISGridAdvEx mGrid;

        #endregion;

        #region ----- Constructor -----

        public FCMF0628_SLIP(ISAppInterface pAppInterface, object pSession_ID, 
                            object pBUDGET_DEPT_ID, object pBUDGET_DEPT_CODE, object pBUDGET_DEPT_NAME,  
                            object pBUDGET_PERIOD,
                            object pBUDGET_APPLY_HEADER_ID,
                            InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            mSession_ID = pSession_ID;

            V_BUDGET_APPLY_HEADER_ID.EditValue = pBUDGET_APPLY_HEADER_ID;
            V_BUDGET_DEPT_ID.EditValue = pBUDGET_DEPT_ID;
            V_BUDGET_DEPT_CODE.EditValue = pBUDGET_DEPT_CODE;
            V_BUDGET_DEPT_NAME.EditValue = pBUDGET_DEPT_NAME;
            V_BUDGET_PERIOD.EditValue = pBUDGET_PERIOD;
            V_DATE_FR.EditValue = iDate.ISMonth_1st(pBUDGET_PERIOD);
            V_DATE_TO.EditValue = iDate.ISMonth_Last(pBUDGET_PERIOD);

            mGrid = pGrid;
        }

        #endregion;

        #region ----- Private Methods -----

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

        private void SEARCH_DB()
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            INIT_MANAGEMENT_COLUMN();
            Application.DoEvents();
             
            CHECK_YN.CheckBoxValue = "N";
            IGR_BUDGET_APPLY_SLIP_LIST.LastConfirmChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.Refillable = true;
            
            IDA_BUDGET_APPLY_SLIP_LIST.Fill();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IGR_BUDGET_APPLY_SLIP_LIST.Focus();             
        }

        private void Set_Grid_Control(object pCELL_STATUS)
        {
            int vIDX_CHECK = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CHECK_YN");
            IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[vIDX_CHECK].Insertable = pCELL_STATUS;
            IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[vIDX_CHECK].Updatable = pCELL_STATUS;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }
            IGR_BUDGET_APPLY_SLIP_LIST.LastConfirmChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.Refillable = true;

            Set_Selected_Total_Amount();
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            //idaITEM_PROMPT.Fill();
            //if (idaITEM_PROMPT.CurrentRows.Count == 0)
            //{
            //    return;
            //}

            //int mStart_Column = 13;
            //int mIDX_Column;            // 시작 COLUMN.            
            //int mMax_Column = 10;       // 종료 COLUMN.
            //int mENABLED_COLUMN;        // 사용여부 COLUMN.

            //object mENABLED_FLAG;       // 사용(표시)여부.
            //object mCOLUMN_DESC;        // 헤더 프롬프트.

            //for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            //{
            //    mENABLED_COLUMN = mMax_Column + mIDX_Column;
            //    mENABLED_FLAG = idaITEM_PROMPT.CurrentRow[mENABLED_COLUMN];
            //    mCOLUMN_DESC = idaITEM_PROMPT.CurrentRow[mIDX_Column];
            //    if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            //    {
            //        IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
            //    }
            //    else
            //    {
            //        IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
            //        IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
            //        IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
            //    }
            //}

            //// 전표일자 표시
            //mIDX_Column = 0;
            //mIDX_Column = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("GL_DATE");
            //mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            //if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            //{
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            //}
            //else
            //{
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            //}

            //// 적요.
            //mIDX_Column = 0;
            //mIDX_Column = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("SLIP_REMARK");
            //mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["REMARK_YN"]);
            //if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            //{
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            //}
            //else
            //{
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            //}

            //// 외화금액 - 통화관리 하는 경우 적용.
            //mIDX_Column = 0;
            //mIDX_Column = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT");
            //int mIDX_EXCHANGE = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("EXCHANGE_RATE");
            //mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["CURR_CONTROL_YN"]);
            //if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            //{
            //    //외화금액
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Insertable = 0;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Updatable = 0;

            //    //환율
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_EXCHANGE].Visible = 0;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_EXCHANGE].Insertable = 0;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_EXCHANGE].Updatable = 0;
            //}
            //else
            //{
            //    //외화금액
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Insertable = 1;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_Column].Updatable = 1;

            //    //환율
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_EXCHANGE].Visible = 1;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_EXCHANGE].Insertable = 1;
            //    IGR_BUDGET_APPLY_SLIP_LIST.GridAdvExColElement[mIDX_EXCHANGE].Updatable = 1;
            //}
            //IGR_BUDGET_APPLY_SLIP_LIST.ResetDraw = true;
        }

        private void Set_Selected_Total_Amount()
        {
            decimal mTotal_Curr_Amount = 0;
            decimal mTotal_Amount = 0;
            int mIDX_CHECK_YN = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CHECK_YN");
            int mIDX_GL_CURRENCY_AMOUNT = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("GL_CURRENCY_AMOUNT");
            int mIDX_GL_AMOUNT = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("GL_AMOUNT");

            for (int i = 0; i < IGR_BUDGET_APPLY_SLIP_LIST.RowCount; i++)
            {
                if ("Y" == iString.ISNull(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue(i, mIDX_CHECK_YN)))
                {
                    mTotal_Curr_Amount = iString.ISDecimaltoZero(mTotal_Curr_Amount, 0) +
                                        iString.ISDecimaltoZero(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue(i, mIDX_GL_CURRENCY_AMOUNT), 0);

                    mTotal_Amount = iString.ISDecimaltoZero(mTotal_Amount, 0) +
                                    iString.ISDecimaltoZero(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue(i, mIDX_GL_AMOUNT), 0);
                }
            }
            TOTAL_CURR_AMOUNT.EditValue = mTotal_Curr_Amount;
            TOTAL_AMOUNT.EditValue = mTotal_Amount;
        }

        private void Set_Selected_Completed()
        {
            IGR_BUDGET_APPLY_SLIP_LIST.LastConfirmChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.Refillable = true;

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
              
            int mIDX_CHECK_YN = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CHECK_YN");
            int mIDX_HEADER_ID = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("HEADER_ID");
            
            //int mIDX_LINE_ID = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("LINE_ID");
            //int mIDX_SLIP_DATE = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("SLIP_DATE");
            //int mIDX_SLIP_NUM = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("SLIP_NUM");
            //int mIDX_CURRENCY_CODE = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CURRENCY_CODE");
            //int mIDX_GL_CURRENCY_AMOUNT = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("GL_CURRENCY_AMOUNT");
            //int mIDX_GL_AMOUNT = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("GL_AMOUNT");
            //int mIDX_REMARK = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("REMARK");
            //int mIDX_VENDOR_CODE = IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("VENDOR_CODE"); 

            //int mIDX_SLIP_LINE_ID = mGrid.GetColumnToIndex("LINE_ID");
            int mIDX_SLIP_HEADER_ID = mGrid.GetColumnToIndex("HEADER_ID"); 
            bool vEXISTS_FLAG = false;
            decimal vHEADER_ID;

            //isDataTransaction1.BeginTran();
            for (int c = 0; c < IGR_BUDGET_APPLY_SLIP_LIST.RowCount; c++)
            {
                if (iString.ISNull(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    IGR_BUDGET_APPLY_SLIP_LIST.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BUDGET_APPLY_SLIP_LIST.CurrentCellActivate(c, mIDX_CHECK_YN);

                    vHEADER_ID = iString.ISDecimaltoZero(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue(c, mIDX_HEADER_ID), -1);
                    
                    //기존 선택된 그리드에 값이 있으면 skip//
                    vEXISTS_FLAG = false;
                    for(int r =0; r < mGrid.RowCount;r++)
                    {
                        if (vHEADER_ID == iString.ISDecimaltoZero(mGrid.GetCellValue(r, mIDX_SLIP_HEADER_ID), 0))
                        {
                            vEXISTS_FLAG = true;
                        } 
                    }

                    if(vEXISTS_FLAG == true)
                    {
                        //
                    }
                    else
                    {
                        //선택한 내용 저장// 
                        IDC_SAVE_BUDGET_APPLY_SLIP.SetCommandParamValue("P_SESSION_ID", mSession_ID);
                        IDC_SAVE_BUDGET_APPLY_SLIP.SetCommandParamValue("P_HEADER_ID", IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue(c, mIDX_HEADER_ID)); 
                        IDC_SAVE_BUDGET_APPLY_SLIP.ExecuteNonQuery();
                        mSTATUS = iString.ISNull(IDC_SAVE_BUDGET_APPLY_SLIP.GetCommandParamValue("O_STATUS"));
                        mMESSAGE = iString.ISNull(IDC_SAVE_BUDGET_APPLY_SLIP.GetCommandParamValue("O_MESSAGE"));
                        if (mSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            Application.DoEvents();

                            if (mMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                    }
                }
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            this.DialogResult = System.Windows.Forms.DialogResult.OK; 
        }

        private void Set_Cancel_Closed()
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        #endregion;

        #region ---- 에디터 프롬프트 리턴 -----

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

        #endregion

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

        private void FCMF0628_SLIP_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_APPLY_SLIP_LIST.FillSchema();
        }

        private void FCMF0628_SLIP_Shown(object sender, EventArgs e)
        {
            //R_GL_DATE.CheckedState = ISUtil.Enum.CheckedState.Checked;
            //V_DATE_TYPE.EditValue = R_GL_DATE.RadioCheckedString;

            Application.DoEvents(); 
        }

        private void V_ACCOUNT_CODE_KeyUp(object pSender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.F9)
            {
                SEARCH_DB();
            }
        }

        private void V_GL_DATE_Click(object sender, EventArgs e)
        {
            //if (R_GL_DATE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            //{
            //    V_DATE_TYPE.EditValue = R_GL_DATE.RadioCheckedString;
            //}
        }

        private void V_DUE_DATE_Click(object sender, EventArgs e)
        {
            //if (R_DUE_DATE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            //{
            //    V_DATE_TYPE.EditValue = R_DUE_DATE.RadioCheckedString;
            //}
        }

        private void V_DATE_FR_KeyUp(object pSender, KeyEventArgs e)
        {

        }

        private void V_DATE_TO_KeyUp(object pSender, KeyEventArgs e)
        {

        }         

        private void IGR_BALANCE_REMAIN_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CHECK_YN"))
            {
                Set_Selected_Total_Amount();
            }
            else if (e.ColIndex == IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("REMAIN_AMOUNT"))
            {
                decimal mREMAIN_CURR_AMOUNT = 0;
                decimal mREMAIN_AMOUNT = iString.ISDecimaltoZero(e.NewValue);
                decimal mEXCHANGE_RATE = iString.ISDecimaltoZero(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue("EXCHANGE_RATE"));
                if (iString.ISNull(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue("CONTROL_CURRENCY_YN")) == "Y" && mEXCHANGE_RATE != 0)
                {
                    mREMAIN_CURR_AMOUNT = Math.Round(mREMAIN_AMOUNT / mEXCHANGE_RATE, 2);
                    IGR_BUDGET_APPLY_SLIP_LIST.SetCellValue("CURR_REMAIN_AMOUNT", mREMAIN_CURR_AMOUNT);
                }
            }
            else if (e.ColIndex == IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT"))
            {
                decimal mREMAIN_AMOUNT = 0;
                decimal mREMAIN_CURR_AMOUNT = iString.ISDecimaltoZero(e.NewValue);
                decimal mEXCHANGE_RATE = iString.ISDecimaltoZero(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue("EXCHANGE_RATE"));
                if (iString.ISNull(IGR_BUDGET_APPLY_SLIP_LIST.GetCellValue("CONTROL_CURRENCY_YN")) == "Y" && mEXCHANGE_RATE != 0)
                {
                    mREMAIN_AMOUNT = Math.Round(mREMAIN_CURR_AMOUNT * mEXCHANGE_RATE, 2);
                    IGR_BUDGET_APPLY_SLIP_LIST.SetCellValue("REMAIN_AMOUNT", mREMAIN_AMOUNT);
                }
            }
        }

        private void IGR_BALANCE_REMAIN_LIST_CellKeyUp(object pSender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.S)
            {
                Set_Selected_Completed();
            }
            else if (e.Control == true && e.KeyCode == Keys.Q)
            {
                Set_Cancel_Closed();
            }
        }

        private void IGR_BALANCE_REMAIN_LIST_CurrentCellEditingComplete(object pSender, ISGridAdvExCellEditingEventArgs e)
        {
            if (e.ColIndex == IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("REMAIN_AMOUNT"))
            {
                Set_Selected_Total_Amount();
            }
            else if (e.ColIndex == IGR_BUDGET_APPLY_SLIP_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT"))
            {
                Set_Selected_Total_Amount();
            }
            IGR_BUDGET_APPLY_SLIP_LIST.LastConfirmChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.Refillable = true;
        }

        private void isbtnSEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Selected_Completed();
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Cancel_Closed();
        }

        private void CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(IGR_BUDGET_APPLY_SLIP_LIST, CHECK_YN.CheckBoxValue);
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ILA_BUDGET_ACCOUNT_CONTROL_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BUDGET_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
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
            IGR_BUDGET_APPLY_SLIP_LIST.LastConfirmChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_SLIP_LIST.Refillable = true;
        }

        private void IDA_BALANCE_REMAIN_LIST_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Set_Selected_Total_Amount();
        }

        #endregion

    }
}