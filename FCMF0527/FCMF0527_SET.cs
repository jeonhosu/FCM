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

namespace FCMF0527
{
    public partial class FCMF0527_SET : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public object Get_Due_Date
        {
            get
            {
                return V_GL_DATE.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public FCMF0527_SET(ISAppInterface pAppInterface, object pDUE_DATE)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_GL_DATE.EditValue = pDUE_DATE;
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
            if(iString.ISNull(V_GL_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(V_GL_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_GL_DATE.Focus();
                return;
            }
            //if (iString.ISNull(V_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(V_ACCOUNT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    V_ACCOUNT_CODE.Focus();
            //    return;
            //}            

            CHECK_YN.CheckBoxValue = "N";
            IGR_BILL_SLIP_LIST.LastConfirmChanges();
            IDA_BILL_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BILL_SLIP_LIST.Refillable = true;
            
            IDA_BILL_SLIP_LIST.Fill();
            IGR_BILL_SLIP_LIST.Focus();
        }

        private void Set_Grid_Control(object pCELL_STATUS)
        {
            int vIDX_CHECK = IGR_BILL_SLIP_LIST.GetColumnToIndex("CHECK_YN");
            IGR_BILL_SLIP_LIST.GridAdvExColElement[vIDX_CHECK].Insertable = pCELL_STATUS;
            IGR_BILL_SLIP_LIST.GridAdvExColElement[vIDX_CHECK].Updatable = pCELL_STATUS;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }
            IGR_BILL_SLIP_LIST.LastConfirmChanges();
            IDA_BILL_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BILL_SLIP_LIST.Refillable = true;
        }

        private void INIT_MANAGEMENT_COLUMN()
        {
            idaITEM_PROMPT.Fill();
            if (idaITEM_PROMPT.CurrentRows.Count == 0)
            {
                return;
            }

            int mStart_Column = 6;
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
                    IGR_BILL_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BILL_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                    IGR_BILL_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    IGR_BILL_SLIP_LIST.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }

            // 전표일자 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_BILL_SLIP_LIST.GetColumnToIndex("GL_DATE");
            mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 적요.
            mIDX_Column = 0;
            mIDX_Column = IGR_BILL_SLIP_LIST.GetColumnToIndex("REMARK");
            mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
            }

            // 외화금액 - 통화관리 하는 경우 적용.
            mIDX_Column = 0;
            mIDX_Column = IGR_BILL_SLIP_LIST.GetColumnToIndex("GL_CURR_AMOUNT");
            mENABLED_FLAG = iString.ISNull(idaITEM_PROMPT.CurrentRow["CONTROL_CURRENCY_YN"]);
            if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 0;
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Insertable = 0;
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Updatable = 0;
            }
            else
            {
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Visible = 1;
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Insertable = 1;
                IGR_BILL_SLIP_LIST.GridAdvExColElement[mIDX_Column].Updatable = 1;
            }
            IGR_BILL_SLIP_LIST.ResetDraw = true;
        }

        private void Set_Selected_Total_Amount()
        {
            decimal mTotal_Curr_Amount = 0;
            decimal mTotal_Amount = 0;
            int mIDX_CHECK_YN = IGR_BILL_SLIP_LIST.GetColumnToIndex("CHECK_YN");
            int mIDX_REMAIN_CURR_AMOUNT = IGR_BILL_SLIP_LIST.GetColumnToIndex("GL_CURR_AMOUNT");
            int mIDX_REMAIN_AMOUNT = IGR_BILL_SLIP_LIST.GetColumnToIndex("GL_AMOUNT");

            for (int i = 0; i < IGR_BILL_SLIP_LIST.RowCount; i++)
            {
                if ("Y" == iString.ISNull(IGR_BILL_SLIP_LIST.GetCellValue(i, mIDX_CHECK_YN)))
                {
                    mTotal_Curr_Amount = iString.ISDecimaltoZero(mTotal_Curr_Amount, 0) + 
                                        iString.ISDecimaltoZero(IGR_BILL_SLIP_LIST.GetCellValue(i, mIDX_REMAIN_CURR_AMOUNT), 0);

                    mTotal_Amount = iString.ISDecimaltoZero(mTotal_Amount, 0) +
                                    iString.ISDecimaltoZero(IGR_BILL_SLIP_LIST.GetCellValue(i, mIDX_REMAIN_AMOUNT), 0);
                }
            }
            TOTAL_CURR_AMOUNT.EditValue = mTotal_Curr_Amount;
            TOTAL_AMOUNT.EditValue = mTotal_Amount;
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

        private void FCMF0527_SET_Load(object sender, EventArgs e)
        {
            IDA_BILL_SLIP_LIST.FillSchema();
        }

        private void FCMF0527_SET_Shown(object sender, EventArgs e)
        {
            
        }

        private void IGR_BALANCE_REMAIN_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BILL_SLIP_LIST.GetColumnToIndex("CHECK_YN"))
            {
                Set_Selected_Total_Amount();
            }
        }

        private void IGR_BALANCE_REMAIN_LIST_CurrentCellEditingComplete(object pSender, ISGridAdvExCellEditingEventArgs e)
        {
            if (e.ColIndex == IGR_BILL_SLIP_LIST.GetColumnToIndex("GL_AMOUNT"))
            {
                Set_Selected_Total_Amount();
            }
            else if (e.ColIndex == IGR_BILL_SLIP_LIST.GetColumnToIndex("GL_CURR_AMOUNT"))
            {
                Set_Selected_Total_Amount();
            }
            IGR_BILL_SLIP_LIST.LastConfirmChanges();
            IDA_BILL_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BILL_SLIP_LIST.Refillable = true;
        }

        private void isbtnSEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IGR_BILL_SLIP_LIST.LastConfirmChanges();
            IDA_BILL_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BILL_SLIP_LIST.Refillable = true;
            
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();


            DialogResult dlgResult;
            //기존에 처리중인 데이터 존재 여부 체크.
            IDC_PROC_BILL_SLIP_COUNT.ExecuteNonQuery();
            int vRECORD_COUNT = iString.ISNumtoZero(IDC_PROC_BILL_SLIP_COUNT.GetCommandParamValue("O_COUNT"));
            if (vRECORD_COUNT > 0)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                dlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10070"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgResult == DialogResult.No)
                {
                    return;
                }
                else
                {
                    //기존 처리중인 데이터 초기화.
                    IDC_INIT_BILL_SLIP.ExecuteNonQuery();
                }
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
            int mIDX_CHECK_YN = IGR_BILL_SLIP_LIST.GetColumnToIndex("CHECK_YN");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BILL_SLIP_LIST.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_BILL_TYPE = IGR_BILL_SLIP_LIST.GetColumnToIndex("BILL_TYPE");
            int mIDX_BILL_NUM = IGR_BILL_SLIP_LIST.GetColumnToIndex("BILL_NUM");
            int mIDX_REMARK = IGR_BILL_SLIP_LIST.GetColumnToIndex("REMARK");
            isDataTransaction1.BeginTran();
            for (int c = 0; c < IGR_BILL_SLIP_LIST.RowCount; c++)
            {
                if (iString.ISNull(IGR_BILL_SLIP_LIST.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_BILL_SLIP_LIST.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BILL_SLIP_LIST.CurrentCellActivate(c, mIDX_CHECK_YN);

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
 
                    IDC_SAVE_BILL_SLIP_LIST.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BILL_SLIP_LIST.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_SAVE_BILL_SLIP_LIST.SetCommandParamValue("W_BILL_TYPE", IGR_BILL_SLIP_LIST.GetCellValue(c, mIDX_BILL_TYPE));
                    IDC_SAVE_BILL_SLIP_LIST.SetCommandParamValue("W_BILL_NUM", IGR_BILL_SLIP_LIST.GetCellValue(c, mIDX_BILL_NUM));
                    IDC_SAVE_BILL_SLIP_LIST.SetCommandParamValue("P_REMARK", IGR_BILL_SLIP_LIST.GetCellValue(c, mIDX_REMARK));
                    IDC_SAVE_BILL_SLIP_LIST.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_SAVE_BILL_SLIP_LIST.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_SAVE_BILL_SLIP_LIST.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SAVE_BILL_SLIP_LIST.ExcuteError || mSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void CHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(IGR_BILL_SLIP_LIST, CHECK_YN.CheckBoxValue);
            Set_Selected_Total_Amount();
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("P_DUE_DATE_TO", V_GL_DATE.EditValue);
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
            //INIT_MANAGEMENT_COLUMN();
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
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
            IGR_BILL_SLIP_LIST.LastConfirmChanges();
            IDA_BILL_SLIP_LIST.OraSelectData.AcceptChanges();
            IDA_BILL_SLIP_LIST.Refillable = true;
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