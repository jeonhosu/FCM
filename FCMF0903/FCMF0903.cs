using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;

using InfoSummit.Win.ControlAdv;

namespace FCMF0903
{
    public partial class FCMF0903 : Office2007Form
    {
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();
        private ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0903()
        {
            InitializeComponent();
        }

        public FCMF0903(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;
        
        #region ----- Methods ----

        private bool Check_Inquiry_Condition()
        {
            if (iConv.ISNull(W_CLOSED_YEAR.EditValue) == string.Empty)
            {
                //년도는 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CLOSED_YEAR.Focus();
                return false;
            }

            return true;
        }

        private void Search_DB()
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            IDA_BALANCE_ACCOUNT.Fill();
            IGR_BALANCE_ACCOUNT.Focus(); 
        }

        private void Search_Balance_Statement(object pACCOUNT_CONTROL_ID)
        {
            if (iConv.ISNull(pACCOUNT_CONTROL_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            IGR_BALANCE_STATEMENT.LastConfirmChanges();
            IDA_BALANCE_STATEMENT_FW.OraSelectData.AcceptChanges();
            IDA_BALANCE_STATEMENT_FW.Refillable = true;

            try
            {
                IDA_BALANCE_STATEMENT_FW.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", -1);
                IDA_BALANCE_STATEMENT_FW.Fill();
            }
            catch
            {
                //
            }

            INIT_MANAGEMENT_COLUMN(pACCOUNT_CONTROL_ID);
            Application.DoEvents();

            IDA_BALANCE_STATEMENT_FW.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_BALANCE_STATEMENT_FW.Fill(); 
        }

        private void Search_Detail()
        {   
            IDA_BALANCE_STATEMENT_DTL.SetSelectParamValue("W_CONFIRM_FLAG", "A");
            IDA_BALANCE_STATEMENT_DTL.Fill();
            IGR_BALANCE_STATEMENT_DTL.Focus(); 
        }

        private void INIT_MANAGEMENT_COLUMN(object pACCOUNT_CONTROL_ID)
        {
            IDA_ITEM_PROMPT.SetSelectParamValue("W_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_ITEM_PROMPT.Fill();

            int mStart_Column = 3;
            int mIDX_Column;            // 시작 COLUMN.    
            int mMax_Column = 10;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            if (IDA_ITEM_PROMPT.OraSelectData.Rows.Count == 0)
            {
                for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mENABLED_COLUMN = mMax_Column + mIDX_Column;
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0; 
                }

                // 전표일자 표시
                mIDX_Column = 0;
                mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE");
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                 
                // 적요.
                mIDX_Column = 0;
                mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("SLIP_REMARK");
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                 
                // 환율 - 통화관리 하는 경우 적용.
                mIDX_Column = 0;
                mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("EXCHANGE_RATE");
                //환율
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                //외화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 0;
                 
                // 환산환율 적용 - 환산환율 관리 하는 경우 적용.
                mIDX_Column = 0;
                mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("NEW_EXCHANGE_RATE");
                // 환산환율.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                //환산원화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 0;
                //환산손익.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 2].Visible = 0;
               
                IGR_BALANCE_STATEMENT.ResetDraw = true;
                return;
            }

            
            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = IDA_ITEM_PROMPT.CurrentRow[mENABLED_COLUMN];

                if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                }
            }

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = IDA_ITEM_PROMPT.CurrentRow[mIDX_Column];
                if (iConv.ISNull(mCOLUMN_DESC) != string.Empty)
                {
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = mCOLUMN_DESC.ToString();
                    IGR_BALANCE_STATEMENT.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = mCOLUMN_DESC.ToString();
                }
            }

            // 전표일자 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE");
            mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0; 
            }
            else
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1; 
            }

            // 적요.
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("SLIP_REMARK");
            mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT.CurrentRow["REMARK_YN"]);
            if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0; 
            }
            else
            {
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1; 
            }

            // 환율 - 통화관리 하는 경우 적용.
            mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT.CurrentRow["CURR_CONTROL_YN"]);            
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("EXCHANGE_RATE"); 
            if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                //환율
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                //외화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 0;
            }
            else
            {
                //환율
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                //외화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 1;
            } 

            // 환산환율 적용 - 환산환율 관리 하는 경우 적용.
            mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT.CurrentRow["CURR_ESTIMATE_YN"]);
            mIDX_Column = 0;
            mIDX_Column = IGR_BALANCE_STATEMENT.GetColumnToIndex("NEW_EXCHANGE_RATE");
            if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                // 환산환율.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 0;
                //환산원화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 0;
                //환산손익.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 2].Visible = 0;
            }
            else
            {
                // 환산환율.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column].Visible = 1;
                //환산원화.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 1].Visible = 1;
                //환산손익.
                IGR_BALANCE_STATEMENT.GridAdvExColElement[mIDX_Column + 2].Visible = 1;
            } 
            IGR_BALANCE_STATEMENT.ResetDraw = true;
        }

        private void Select_Check_YN(object pCHECK_FLAG)
        {
            int vIDX_CHECK = IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < IGR_BALANCE_STATEMENT.RowCount; i++)
            {
                IGR_BALANCE_STATEMENT.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }
            IGR_BALANCE_STATEMENT.LastConfirmChanges();
            IDA_BALANCE_STATEMENT_FW.OraSelectData.AcceptChanges();
            IDA_BALANCE_STATEMENT_FW.Refillable = true; 
        }

        #endregion;

        #region ----- MDi ToolBar Button Event -----

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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BALANCE_STATEMENT_FW.IsFocused)
                    {
                        IDA_BALANCE_STATEMENT_FW.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0903_Load(object sender, EventArgs e)
        {
            W_CLOSED_YEAR.EditValue = iDate.ISYear(System.DateTime.Today);

            W_BALANCE_DATE_FR.EditValue = iDate.ISGetDate(string.Format("{0}-01-01", W_CLOSED_YEAR.EditValue));
            W_BALANCE_DATE_TO.EditValue = iDate.ISGetDate(string.Format("{0}-12-31", W_CLOSED_YEAR.EditValue));
        }

        private void FCMF0903_Shown(object sender, EventArgs e)
        {
            V_RB_NO_CARRIED_FORWARD.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_CARRIED_FORWORD_FLAG.EditValue = V_RB_NO_CARRIED_FORWARD.RadioCheckedString;

            BTN_CARRIED_FORWARD_OK.Enabled = true;
            BTN_CARRIED_FORWARD_CANCEL.Enabled = false;

            IDA_BALANCE_STATEMENT_FW.FillSchema();
            IDA_BALANCE_ACCOUNT.FillSchema();        
        }

        private void IGR_BALANCE_STATEMENT_CellDoubleClick(object pSender)
        {
            if (IGR_BALANCE_STATEMENT.RowCount > 0)
            {
                TB_BALANCE_STATEMENT.SelectedIndex = 1; 
                TB_BALANCE_STATEMENT.Focus();

                Search_Detail();
            }
        }

        private void IGR_BALANCE_STATEMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN"))
            {
                IGR_BALANCE_STATEMENT.LastConfirmChanges();
                IDA_BALANCE_STATEMENT_FW.OraSelectData.AcceptChanges();
                IDA_BALANCE_STATEMENT_FW.Refillable = true; 
            }
        }

        private void CB_SELECT_ALL_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(CB_SELECT_ALL.CheckBoxValue);
        }

        private void BTN_BSD_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Search_Detail();
        }

        private void V_RB_NO_CARRIED_FORWARD_Click(object sender, EventArgs e)
        {
            if (V_RB_NO_CARRIED_FORWARD.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_CARRIED_FORWORD_FLAG.EditValue = V_RB_NO_CARRIED_FORWARD.RadioCheckedString;
                BTN_CARRIED_FORWARD_OK.Enabled = true;
                BTN_CARRIED_FORWARD_CANCEL.Enabled = false;
            }
        }

        private void V_RB_CARRIED_FORWARD_Click(object sender, EventArgs e)
        {
            if (V_RB_CARRIED_FORWARD.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_CARRIED_FORWORD_FLAG.EditValue = V_RB_CARRIED_FORWARD.RadioCheckedString;
                BTN_CARRIED_FORWARD_OK.Enabled = false;
                BTN_CARRIED_FORWARD_CANCEL.Enabled = true;
            }
        }

        #endregion

        #region ----- Button Event -----
         
        private void BTN_CARRIED_FORWARD_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            string vSTATUS = null;
            string vMESSAGE = null;
             
            //처리여부 묻기//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_CHECK_YN = IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN");
            int vIDX_BALANCE_STATEMENT_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("BALANCE_STATEMENT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int vIDX_CURRENCY_CODE = IGR_BALANCE_STATEMENT.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_ITEM_GROUP_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("ITEM_GROUP_ID");
            int vIDX_GL_DATE_FR = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE_FR");
            int vIDX_GL_DATE_TO = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE_TO");
            int vIDX_NEW_BALANCE_DATE = IGR_BALANCE_STATEMENT.GetColumnToIndex("NEW_BALANCE_DATE");

            object vGL_DATE_YN = IGR_BALANCE_ACCOUNT.GetCellValue("GL_DATE_YN");
            
            try
            {
                for (int r = 0; r < IGR_BALANCE_STATEMENT.RowCount; r++)
                {
                    if (iConv.ISNull(IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                    {
                        IGR_BALANCE_STATEMENT.CurrentCellActivate(r, vIDX_CHECK_YN);
                        IGR_BALANCE_STATEMENT.CurrentCellMoveTo(r, vIDX_CHECK_YN);

                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_BALANCE_STATEMENT_ID));
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_ACCOUNT_CONTROL_ID));
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_CURRENCY_CODE", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_CURRENCY_CODE));
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_ITEM_GROUP_ID));
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_GL_DATE_YN", vGL_DATE_YN);
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_GL_DATE_FR", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_GL_DATE_FR));
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_GL_DATE_TO", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_GL_DATE_TO));
                        IDC_SET_BALANCE_STATEMENT.SetCommandParamValue("P_NEW_BALANCE_DATE", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_NEW_BALANCE_DATE));
                        IDC_SET_BALANCE_STATEMENT.ExecuteNonQuery();
                        vSTATUS = iConv.ISNull(IDC_SET_BALANCE_STATEMENT.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iConv.ISNull(IDC_SET_BALANCE_STATEMENT.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            return;
                        }

                        IGR_BALANCE_STATEMENT.SetCellValue(r, vIDX_CHECK_YN, "N");
                    }
                }

                IGR_BALANCE_STATEMENT.LastConfirmChanges();
                IDA_BALANCE_STATEMENT_FW.OraSelectData.AcceptChanges();
                IDA_BALANCE_STATEMENT_FW.Refillable = true;
            }
            catch (System.Exception ex)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);     
        }

        private void BTN_CARRIED_FORWARD_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            string vSTATUS = null;
            string vMESSAGE = null;
             
            //처리여부 묻기//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10436"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_CHECK_YN = IGR_BALANCE_STATEMENT.GetColumnToIndex("CHECK_YN");
            int vIDX_BALANCE_STATEMENT_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("BALANCE_STATEMENT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int vIDX_CURRENCY_CODE = IGR_BALANCE_STATEMENT.GetColumnToIndex("CURRENCY_CODE");
            int vIDX_ITEM_GROUP_ID = IGR_BALANCE_STATEMENT.GetColumnToIndex("ITEM_GROUP_ID"); 
            int vIDX_GL_DATE_FR = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE_FR");
            int vIDX_GL_DATE_TO = IGR_BALANCE_STATEMENT.GetColumnToIndex("GL_DATE_TO");
            int vIDX_NEW_BALANCE_DATE = IGR_BALANCE_STATEMENT.GetColumnToIndex("NEW_BALANCE_DATE");
            
            object vGL_DATE_YN = IGR_BALANCE_ACCOUNT.GetCellValue("GL_DATE_YN");

            try
            {
                for (int r = 0; r < IGR_BALANCE_STATEMENT.RowCount; r++)
                {
                    if (iConv.ISNull(IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                    {
                        IGR_BALANCE_STATEMENT.CurrentCellActivate(r, vIDX_CHECK_YN);
                        IGR_BALANCE_STATEMENT.CurrentCellMoveTo(r, vIDX_CHECK_YN);

                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_BALANCE_STATEMENT_ID));
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_ACCOUNT_CONTROL_ID));
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_CURRENCY_CODE", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_CURRENCY_CODE));
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_ITEM_GROUP_ID));
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_GL_DATE_YN", vGL_DATE_YN);
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_GL_DATE_FR", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_GL_DATE_FR));
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_GL_DATE_TO", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_GL_DATE_TO));
                        IDC_CANCEL_BALANCE_STATEMENT.SetCommandParamValue("P_NEW_BALANCE_DATE", IGR_BALANCE_STATEMENT.GetCellValue(r, vIDX_NEW_BALANCE_DATE));
                        IDC_CANCEL_BALANCE_STATEMENT.ExecuteNonQuery();
                        vSTATUS = iConv.ISNull(IDC_CANCEL_BALANCE_STATEMENT.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iConv.ISNull(IDC_CANCEL_BALANCE_STATEMENT.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            return;
                        }
                        
                        IGR_BALANCE_STATEMENT.SetCellValue(r, vIDX_CHECK_YN, "N");
                    }
                }

                IGR_BALANCE_STATEMENT.LastConfirmChanges();
                IDA_BALANCE_STATEMENT_FW.OraSelectData.AcceptChanges();
                IDA_BALANCE_STATEMENT_FW.Refillable = true;
            }
            catch (System.Exception ex)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);   
        }
         
        #endregion;

        #region ----- Lookup Event -----
        
        private void ilaACCOUNT_CONTROL_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaPERIOD_YEAR_W_SelectedRowData(object pSender)
        {
            W_BALANCE_DATE_FR.EditValue = iDate.ISGetDate(string.Format("{0}-01-01", W_CLOSED_YEAR.EditValue));
            W_BALANCE_DATE_TO.EditValue = iDate.ISGetDate(string.Format("{0}-12-31", W_CLOSED_YEAR.EditValue));
        }

        #endregion;

        #region ----- Adapter Event -----

        private void IDA_BALANCE_ACCOUNT_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Search_Balance_Statement(-1);
                return;
            }

            Search_Balance_Statement(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
        }

        #endregion

    }
}