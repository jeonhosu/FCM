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

namespace FCMF0203
{
    public partial class FCMF0203_SET : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mSession_ID;
        object mCURRENCY_CODE;

        #endregion;

        #region ----- Constructor -----

        public FCMF0203_SET(ISAppInterface pAppInterface, object pSession_ID, object pCURRENCY_CODE, object pBATCH_TYPE, object pACCOUNT_CONTROL_ID, object pACCOUNT_CODE, object pACCOUNT_DESC, object pGL_DATE)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            mSession_ID = pSession_ID;
            mCURRENCY_CODE = pCURRENCY_CODE;

            V_BATCH_TYPE.EditValue = pBATCH_TYPE;
            V_ACCOUNT_CONTROL_ID.EditValue = pACCOUNT_CONTROL_ID;
            V_ACCOUNT_CODE.EditValue = pACCOUNT_CODE;
            V_ACCOUNT_DESC.EditValue = pACCOUNT_DESC;
            V_GL_DATE.EditValue = pGL_DATE;
            V_DATE_FR.EditValue = DBNull.Value;
            V_DATE_TO.EditValue = DBNull.Value;
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
            if(iString.ISNull(V_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(V_ACCOUNT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_ACCOUNT_CODE.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            INIT_MANAGEMENT_COLUMN();
            Application.DoEvents();
             
            CHECK_YN.CheckBoxValue = "N";
            IGR_BALANCE_REMAIN_LIST.LastConfirmChanges();
            IDA_BALANCE_REMAIN_LIST.OraSelectData.AcceptChanges();
            IDA_BALANCE_REMAIN_LIST.Refillable = true;
            
            IDA_BALANCE_REMAIN_LIST.Fill();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IGR_BALANCE_REMAIN_LIST.Focus();             
        }

        private void Set_Grid_Control(object pCELL_STATUS)
        {
            int vIDX_CHECK = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CHECK_YN");
            IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[vIDX_CHECK].Insertable = pCELL_STATUS;
            IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[vIDX_CHECK].Updatable = pCELL_STATUS;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }
            IGR_BALANCE_REMAIN_LIST.LastConfirmChanges();
            IDA_BALANCE_REMAIN_LIST.OraSelectData.AcceptChanges();
            IDA_BALANCE_REMAIN_LIST.Refillable = true;

            Set_Selected_Total_Amount();
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            idaITEM_PROMPT.Fill();
            if (idaITEM_PROMPT.CurrentRows.Count == 0)
            {
                return;
            }

            int mStart_Column = 13;
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = idaITEM_PROMPT.CurrentRow[mENABLED_COLUMN];
                mCOLUMN_DESC = idaITEM_PROMPT.CurrentRow[mIDX_Column];
                if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                    IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }

            // 전표일자 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("GL_DATE");
            mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 적요.
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("SLIP_REMARK");
            mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 외화금액 - 통화관리 하는 경우 적용.
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT");
            int mIDX_EXCHANGE = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("EXCHANGE_RATE");
            mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["CURR_CONTROL_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                //외화금액
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Updatable = 0;

                //환율
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_EXCHANGE].Visible = 0;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_EXCHANGE].Insertable = 0;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_EXCHANGE].Updatable = 0;
            }
            else
            {
                //외화금액
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Insertable = 1;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_Column].Updatable = 1;

                //환율
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_EXCHANGE].Visible = 1;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_EXCHANGE].Insertable = 1;
                IGR_BALANCE_REMAIN_LIST.GridAdvExColElement[mIDX_EXCHANGE].Updatable = 1;
            }
            IGR_BALANCE_REMAIN_LIST.ResetDraw = true;
        }

        private void Set_Selected_Total_Amount()
        {
            decimal mTotal_Curr_Amount = 0;
            decimal mTotal_Amount = 0;
            int mIDX_CHECK_YN = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CHECK_YN");
            int mIDX_REMAIN_CURR_AMOUNT = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT");
            int mIDX_REMAIN_AMOUNT = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("REMAIN_AMOUNT");

            for (int i = 0; i < IGR_BALANCE_REMAIN_LIST.RowCount; i++)
            {
                if ("Y" == iString.ISNull(IGR_BALANCE_REMAIN_LIST.GetCellValue(i, mIDX_CHECK_YN)))
                {
                    mTotal_Curr_Amount = iString.ISDecimaltoZero(mTotal_Curr_Amount, 0) + 
                                        iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue(i, mIDX_REMAIN_CURR_AMOUNT), 0);

                    mTotal_Amount = iString.ISDecimaltoZero(mTotal_Amount, 0) +
                                    iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue(i, mIDX_REMAIN_AMOUNT), 0);
                }
            }
            TOTAL_CURR_AMOUNT.EditValue = mTotal_Curr_Amount;
            TOTAL_AMOUNT.EditValue = mTotal_Amount;
        }

        private void Set_Selected_Completed()
        {
            IGR_BALANCE_REMAIN_LIST.LastConfirmChanges();
            IDA_BALANCE_REMAIN_LIST.OraSelectData.AcceptChanges();
            IDA_BALANCE_REMAIN_LIST.Refillable = true;

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
            int mIDX_CHECK_YN = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CHECK_YN");

            int mIDX_REMAIN_AMOUNT = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("REMAIN_AMOUNT");
            int mIDX_CURR_REMAIN_AMOUNT = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT");

            int mIDX_OLD_REMAIN_AMOUNT = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("OLD_REMAIN_AMOUNT");
            int mIDX_OLD_CURR_REMAIN_AMOUNT = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("OLD_CURR_REMAIN_AMOUNT");

            int mIDX_BALANCE_DATE = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("BALANCE_DATE");
            int mIDX_GL_DATE = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("GL_DATE");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_CURRENCY_CODE = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CURRENCY_CODE");
            int mIDX_ITEM_GROUP_ID = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("ITEM_GROUP_ID");
            int mIDX_BALANCE_STATEMENT_ID = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("BALANCE_STATEMENT_ID");

            int mIDX_EXCHANGE_RATE = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("EXCHANGE_RATE");
            int mIDX_REMARK = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("REMARK");
            int mIDX_VENDOR_ID = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("VENDOR_ID");

            //금액 검증//
            decimal vRemain_AMT = 0;
            decimal vCharge_AMT = 0;
            decimal vGap_AMT = 0;
            decimal vCurr_Remain_AMT = 0;
            decimal vCurr_Charge_AMT = 0;
            decimal vCurr_Gap_AMT = 0;
            //isDataTransaction1.BeginTran();
            for (int c = 0; c < IGR_BALANCE_REMAIN_LIST.RowCount; c++)
            {
                if (iString.ISNull(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    IGR_BALANCE_REMAIN_LIST.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BALANCE_REMAIN_LIST.CurrentCellActivate(c, mIDX_CHECK_YN);

                    // 데이터 저장전 검증.
                    //1.1. 금액검증(원래 잔액보다 수정금액이 클수 없음) 
                    vRemain_AMT = 0;
                    vCharge_AMT = 0;
                    vGap_AMT = 0;

                    vRemain_AMT = iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_OLD_REMAIN_AMOUNT)); //잔액//
                    vCharge_AMT = iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_REMAIN_AMOUNT));     //선택한금액//
                    vGap_AMT = vRemain_AMT - vCharge_AMT;
                    if (vRemain_AMT < 0)
                    {//음수 잔액인 경우 처리 후 잔액이 양수이면 오류.
                        vGap_AMT = vGap_AMT * -1;
                    }
                     
                    //1.2. 외화//
                    vCurr_Remain_AMT = 0;
                    vCurr_Charge_AMT = 0;
                    vCurr_Gap_AMT = 0; 
                    vCurr_Remain_AMT = iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_OLD_CURR_REMAIN_AMOUNT)); //잔액//
                    vCurr_Charge_AMT = iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_CURR_REMAIN_AMOUNT));     //선택한금액//
                    vCurr_Gap_AMT = vCurr_Remain_AMT - vCurr_Charge_AMT;
                    if (vCurr_Remain_AMT < 0)
                    {//음수 잔액인 경우 처리 후 잔액이 양수이면 오류.
                        vCurr_Gap_AMT = vCurr_Gap_AMT * -1;
                    } 

                    if (vGap_AMT < 0 && vCurr_Gap_AMT < 0)
                    {//그외 처리 후 잔액이 음수이면 오류.
                        //isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10185"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    } 

                    ////2. 잔액일자, 발생일자, 계정, 통화, 잔액 그룹id, 잔액 헤더 id 값 검증.
                    //mIDX_COL1 = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("BALANCE_DATE");
                    //if (iString.ISNull(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_COL1)) == string.Empty)
                    //{
                    //    isDataTransaction1.RollBack();
                    //    Application.UseWaitCursor = false;
                    //    this.Cursor = System.Windows.Forms.Cursors.Default;
                    //    Application.DoEvents();
                    //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10444"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}
                    ////3. 발생일자, 계정, 통화, 잔액 그룹id, 잔액 헤더 id 값 검증.
                    //mIDX_COL1 = IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("GL_DATE");
                    //if (iString.ISNull(IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_COL1)) == string.Empty)
                    //{
                    //    isDataTransaction1.RollBack();
                    //    Application.UseWaitCursor = false;
                    //    this.Cursor = System.Windows.Forms.Cursors.Default;
                    //    Application.DoEvents();
                    //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10444"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}

                    //선택한 내용 저장//
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_SESSION_ID", mSession_ID);
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_BALANCE_DATE", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_BALANCE_DATE));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_GL_DATE", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_GL_DATE));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_CURRENCY_CODE", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_CURRENCY_CODE));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_ITEM_GROUP_ID));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_BALANCE_STATEMENT_ID));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_GL_AMOUNT", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_REMAIN_AMOUNT));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_GL_CURR_AMOUNT", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_CURR_REMAIN_AMOUNT));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_EXCHANGE_RATE", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_EXCHANGE_RATE));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_REMARK", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_REMARK));
                    IDC_SAVE_BALANCE_REMAIN_TP.SetCommandParamValue("P_VENDOR_ID", IGR_BALANCE_REMAIN_LIST.GetCellValue(c, mIDX_VENDOR_ID));
                    IDC_SAVE_BALANCE_REMAIN_TP.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_SAVE_BALANCE_REMAIN_TP.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_SAVE_BALANCE_REMAIN_TP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SAVE_BALANCE_REMAIN_TP.ExcuteError || mSTATUS == "F")
                    {
                        //isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            //isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
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

        private void FCMF0203_SET_Load(object sender, EventArgs e)
        {
            IDA_BALANCE_REMAIN_LIST.FillSchema();
        }

        private void FCMF0203_SET_Shown(object sender, EventArgs e)
        {
            R_GL_DATE.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_DATE_TYPE.EditValue = R_GL_DATE.RadioCheckedString;

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
            if (R_GL_DATE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_DATE_TYPE.EditValue = R_GL_DATE.RadioCheckedString;
            }
        }

        private void V_DUE_DATE_Click(object sender, EventArgs e)
        {
            if (R_DUE_DATE.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_DATE_TYPE.EditValue = R_DUE_DATE.RadioCheckedString;
            }
        }

        private void V_DATE_FR_KeyUp(object pSender, KeyEventArgs e)
        {

        }

        private void V_DATE_TO_KeyUp(object pSender, KeyEventArgs e)
        {

        }         

        private void IGR_BALANCE_REMAIN_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CHECK_YN"))
            {
                Set_Selected_Total_Amount();
            }
            else if (e.ColIndex == IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("REMAIN_AMOUNT"))
            {
                decimal mREMAIN_CURR_AMOUNT = 0;
                decimal mREMAIN_AMOUNT = iString.ISDecimaltoZero(e.NewValue);
                decimal mEXCHANGE_RATE = iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue("EXCHANGE_RATE"));
                object vCURRENCY = IGR_BALANCE_REMAIN_LIST.GetCellValue("CURRENCY_CODE");
                if (iString.ISNull(IGR_BALANCE_REMAIN_LIST.GetCellValue("CONTROL_CURRENCY_YN")) == "Y" && mEXCHANGE_RATE != 0)
                {
                    try
                    {
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_CURRENCY_TO", vCURRENCY);
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_CURR_AMOUNT", mREMAIN_AMOUNT);
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_EXCHANGE_RATE", mEXCHANGE_RATE);
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_EXCH_TYPE", "SLIP_BALANCESTATEMENT");
                        IDC_BASE_AMOUNT_P.ExecuteNonQuery();
                        mREMAIN_CURR_AMOUNT = iString.ISDecimaltoZero(IDC_BASE_AMOUNT_P.GetCommandParamValue("O_AMOUNT"));
                    }
                    catch
                    {
                        mREMAIN_CURR_AMOUNT = Math.Round(mREMAIN_AMOUNT / mEXCHANGE_RATE, 2);
                    }
                    IGR_BALANCE_REMAIN_LIST.SetCellValue("CURR_REMAIN_AMOUNT", mREMAIN_CURR_AMOUNT);
                }
            }
            else if (e.ColIndex == IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT"))
            {
                decimal mREMAIN_AMOUNT = 0;
                decimal mREMAIN_CURR_AMOUNT = iString.ISDecimaltoZero(e.NewValue);
                decimal mEXCHANGE_RATE = iString.ISDecimaltoZero(IGR_BALANCE_REMAIN_LIST.GetCellValue("EXCHANGE_RATE"));
                if (iString.ISNull(IGR_BALANCE_REMAIN_LIST.GetCellValue("CONTROL_CURRENCY_YN")) == "Y" && mEXCHANGE_RATE != 0)
                {
                    try
                    {
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_CURRENCY_TO", mCURRENCY_CODE);
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_CURR_AMOUNT", mREMAIN_CURR_AMOUNT);
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_EXCHANGE_RATE", mEXCHANGE_RATE);
                        IDC_BASE_AMOUNT_P.SetCommandParamValue("W_EXCH_TYPE", "SLIP_BALANCESTATEMENT");
                        IDC_BASE_AMOUNT_P.ExecuteNonQuery();
                        mREMAIN_AMOUNT = iString.ISDecimaltoZero(IDC_BASE_AMOUNT_P.GetCommandParamValue("O_AMOUNT"));
                    }
                    catch
                    {
                        mREMAIN_AMOUNT = Math.Round(mREMAIN_CURR_AMOUNT * mEXCHANGE_RATE);
                    }
                    IGR_BALANCE_REMAIN_LIST.SetCellValue("REMAIN_AMOUNT", mREMAIN_AMOUNT);
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
            if (e.ColIndex == IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("REMAIN_AMOUNT"))
            {
                Set_Selected_Total_Amount();
            }
            else if (e.ColIndex == IGR_BALANCE_REMAIN_LIST.GetColumnToIndex("CURR_REMAIN_AMOUNT"))
            {
                Set_Selected_Total_Amount();
            }
            IGR_BALANCE_REMAIN_LIST.LastConfirmChanges();
            IDA_BALANCE_REMAIN_LIST.OraSelectData.AcceptChanges();
            IDA_BALANCE_REMAIN_LIST.Refillable = true;
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
            Select_Check_YN(IGR_BALANCE_REMAIN_LIST, CHECK_YN.CheckBoxValue);
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaVENDOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaVENDOR_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        private void ilaACCOUNT_CONTROL_0_SelectedRowData(object pSender)
        {
            INIT_MANAGEMENT_COLUMN();
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
            IGR_BALANCE_REMAIN_LIST.LastConfirmChanges();
            IDA_BALANCE_REMAIN_LIST.OraSelectData.AcceptChanges();
            IDA_BALANCE_REMAIN_LIST.Refillable = true;
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