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
    public partial class FCMF0528_ESTIMATE_SET : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime(); 

        object mACCOUNT_ALL = "N";
        #endregion;

        #region ----- Constructor -----

        public FCMF0528_ESTIMATE_SET(ISAppInterface pAppInterface, object pGL_DATE, string pCURR_ESTIMATE_STATUS, object pACCOUNT_ALL)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;
      
            V_BALANCE_DATE.EditValue = pGL_DATE;
            V_CURR_ESTIMATE_STATUS.EditValue = pCURR_ESTIMATE_STATUS;
            mACCOUNT_ALL = pACCOUNT_ALL;

            if (pCURR_ESTIMATE_STATUS == "CANCEL_ESTIMATE")
            {
                Point BTN_LOCATION = new Point(678, 6);                
                BTN_ESTIMATE_CANCEL.Location = BTN_LOCATION; 
                BTN_ESTIMATE_OK.Visible = false;
                BTN_ESTIMATE_CLOSED.Visible = false;
                BTN_ESTIMATE_CANCEL_CLOSED.Visible = false;
                BTN_ESTIMATE_CANCEL.Visible = true; 
            }
            else if (pCURR_ESTIMATE_STATUS == "CLOSED_ESTIMATE")
            {
                Point BTN_LOCATION = new Point(678, 6);        
                BTN_ESTIMATE_CLOSED.Location = BTN_LOCATION;
                BTN_ESTIMATE_OK.Visible = false;
                BTN_ESTIMATE_CLOSED.Visible = true;
                BTN_ESTIMATE_CANCEL_CLOSED.Visible = false;
                BTN_ESTIMATE_CANCEL.Visible = false;
            }
            else if (pCURR_ESTIMATE_STATUS == "CANCEL_CLOSED_ESTIMATE")
            {
                Point BTN_LOCATION = new Point(678, 6);        
                BTN_ESTIMATE_CANCEL_CLOSED.Location = BTN_LOCATION;
                BTN_ESTIMATE_OK.Visible = false;
                BTN_ESTIMATE_CLOSED.Visible = false;
                BTN_ESTIMATE_CANCEL_CLOSED.Visible = true;
                BTN_ESTIMATE_CANCEL.Visible = false;
            }
            else
            {
                Point BTN_LOCATION = new Point(678, 6);        
                BTN_ESTIMATE_OK.Location = BTN_LOCATION;
                BTN_ESTIMATE_OK.Visible = true;
                BTN_ESTIMATE_CLOSED.Visible = false;
                BTN_ESTIMATE_CANCEL_CLOSED.Visible = false;
                BTN_ESTIMATE_CANCEL.Visible = false; 
            }
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
            if (iConv.ISNull(V_CURR_ESTIMATE_STATUS.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10032"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_CURR_ESTIMATE_STATUS.Focus();
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
            CB_SELECT_ALL.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            IDA_BALANCE_ACCOUNT.SetSelectParamValue("W_ACCOUNT_ALL", mACCOUNT_ALL);
            IDA_BALANCE_ACCOUNT.SetSelectParamValue("W_ACCOUNT_CODE_TO", V_ACCOUNT_CODE.EditValue);
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
            IDA_BAL_CURR_ESTIMATE.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", -1);
            IDA_BAL_CURR_ESTIMATE.Fill();

            INIT_MANAGEMENT_COLUMN(pACCOUNT_CONTROL_ID);
            Application.DoEvents();

            IDA_BAL_CURR_ESTIMATE.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_BAL_CURR_ESTIMATE.Fill();
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
            }

            pGrid.LastConfirmChanges();
            IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_BAL_CURR_ESTIMATE.Refillable = true;
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

                    IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0; 
                }

                // 전표일자 표시
                mIDX_Column = 0;
                mIDX_Column = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("GL_DATE");
                IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mIDX_Column].Visible = 0;
                
                IGR_BAL_CURR_ESTIMATE.ResetDraw = true;
                return;
            }
            
            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = IDA_ITEM_PROMPT.CurrentRow[mENABLED_COLUMN];

                if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].Visible = 1;
                }
            }

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mCOLUMN_DESC = IDA_ITEM_PROMPT.CurrentRow[mIDX_Column];
                if (iConv.ISNull(mCOLUMN_DESC) != string.Empty)
                {
                    IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].Default = mCOLUMN_DESC.ToString();
                    IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mStart_Column + mIDX_Column].HeaderElement[0].TL1_KR = mCOLUMN_DESC.ToString();
                }
            }

            // 전표일자 표시
            mIDX_Column = 0;
            mIDX_Column = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("GL_DATE");
            mENABLED_FLAG = iConv.ISNull(IDA_ITEM_PROMPT.CurrentRow["GL_DATE_YN"]);
            if (iConv.ISNull(mENABLED_FLAG, "N") == "N".ToString())
            {
                IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mIDX_Column].Visible = 0;
            }
            else
            {
                IGR_BAL_CURR_ESTIMATE.GridAdvExColElement[mIDX_Column].Visible = 1;
            }             
            IGR_BAL_CURR_ESTIMATE.ResetDraw = true;
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

        private void FCMF0528_ESTIMATE_SET_Load(object sender, EventArgs e)
        {
            IDA_BAL_CURR_ESTIMATE.FillSchema();
        }

        private void FCMF0528_ESTIMATE_SET_Shown(object sender, EventArgs e)
        {
            CB_SELECT_ALL.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
        }

        private void IGR_BAL_CURR_ESTIMATE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN"))
            {
                IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
                IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
                IDA_BAL_CURR_ESTIMATE.Refillable = true;
            }
        }
        
        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void BTN_ESTIMATE_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CheckData() == false)
            {
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            object mGL_DATE_YN = IGR_BALANCE_ACCOUNT.GetCellValue("GL_DATE_YN");

            int mIDX_CHECK_YN = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN");
            int mIDX_BALANCE_STATEMENT_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("BALANCE_STATEMENT_ID");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_CURRENCY_CODE = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CURRENCY_CODE");
            int mIDX_ITEM_GROUP_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ITEM_GROUP_ID");
            int mIDX_GL_DATE_FR = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("GL_DATE_FR");
            int mIDX_GL_DATE_TO = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("GL_DATE_TO"); 

            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_BAL_CURR_ESTIMATE.RowCount; c++)
            {
                if (iConv.ISNull(IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_BAL_CURR_ESTIMATE.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BAL_CURR_ESTIMATE.CurrentCellActivate(c, mIDX_CHECK_YN);

                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_DATE", V_BALANCE_DATE.EditValue);
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_BALANCE_STATEMENT_ID));
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_CURRENCY_CODE", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CURRENCY_CODE));
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ITEM_GROUP_ID));
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_GL_DATE_YN", mGL_DATE_YN);
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_GL_DATE_FR", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_GL_DATE_FR));
                    IDC_SET_CURR_ESTIMATE.SetCommandParamValue("P_GL_DATE_TO", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_GL_DATE_TO));
                    IDC_SET_CURR_ESTIMATE.ExecuteNonQuery();
                    mSTATUS = iConv.ISNull(IDC_SET_CURR_ESTIMATE.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iConv.ISNull(IDC_SET_CURR_ESTIMATE.GetCommandParamValue("O_MESSAGE"));

                    if (mSTATUS == "F")
                    {
                        IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
                        IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
                        IDA_BAL_CURR_ESTIMATE.Refillable = true;

                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                    IGR_BAL_CURR_ESTIMATE.SetCellValue(c, mIDX_CHECK_YN, "N"); 
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
            IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_BAL_CURR_ESTIMATE.Refillable = true;

            IDA_BAL_CURR_ESTIMATE.Fill();
        }

        private void BTN_ESTIMATE_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CheckData() == false)
            {
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
             
            int mIDX_CHECK_YN = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN");
            int mIDX_BALANCE_STATEMENT_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("BALANCE_STATEMENT_ID");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_CURRENCY_CODE = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CURRENCY_CODE");
            int mIDX_ITEM_GROUP_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ITEM_GROUP_ID"); 

            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_BAL_CURR_ESTIMATE.RowCount; c++)
            {
                if (iConv.ISNull(IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();
                    
                    IGR_BAL_CURR_ESTIMATE.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BAL_CURR_ESTIMATE.CurrentCellActivate(c, mIDX_CHECK_YN);

                    IDC_CANCEL_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_DATE", V_BALANCE_DATE.EditValue);
                    IDC_CANCEL_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_BALANCE_STATEMENT_ID));
                    IDC_CANCEL_CURR_ESTIMATE.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_CANCEL_CURR_ESTIMATE.SetCommandParamValue("P_CURRENCY_CODE", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CURRENCY_CODE));
                    IDC_CANCEL_CURR_ESTIMATE.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ITEM_GROUP_ID));
                    IDC_CANCEL_CURR_ESTIMATE.ExecuteNonQuery();
                    mSTATUS = iConv.ISNull(IDC_CANCEL_CURR_ESTIMATE.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iConv.ISNull(IDC_CANCEL_CURR_ESTIMATE.GetCommandParamValue("O_MESSAGE"));

                    if (mSTATUS == "F")
                    {
                        IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
                        IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
                        IDA_BAL_CURR_ESTIMATE.Refillable = true;

                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                    IGR_BAL_CURR_ESTIMATE.SetCellValue(c, mIDX_CHECK_YN, "N");
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
            IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_BAL_CURR_ESTIMATE.Refillable = true;

            IDA_BAL_CURR_ESTIMATE.Fill();
        }

        private void BTN_ESTIMATE_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CheckData() == false)
            {
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int mIDX_CHECK_YN = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN");
            int mIDX_BALANCE_STATEMENT_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("BALANCE_STATEMENT_ID");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_CURRENCY_CODE = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CURRENCY_CODE");
            int mIDX_ITEM_GROUP_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ITEM_GROUP_ID");

            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_BAL_CURR_ESTIMATE.RowCount; c++)
            {
                if (iConv.ISNull(IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_BAL_CURR_ESTIMATE.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BAL_CURR_ESTIMATE.CurrentCellActivate(c, mIDX_CHECK_YN);

                    IDC_SET_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_DATE", V_BALANCE_DATE.EditValue);
                    IDC_SET_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_BALANCE_STATEMENT_ID));
                    IDC_SET_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_SET_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_CURRENCY_CODE", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CURRENCY_CODE));
                    IDC_SET_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ITEM_GROUP_ID));
                    IDC_SET_CLOSED_CURR_ESTIMATE.ExecuteNonQuery();
                    mSTATUS = iConv.ISNull(IDC_SET_CLOSED_CURR_ESTIMATE.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iConv.ISNull(IDC_SET_CLOSED_CURR_ESTIMATE.GetCommandParamValue("O_MESSAGE"));

                    if (mSTATUS == "F")
                    {
                        IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
                        IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
                        IDA_BAL_CURR_ESTIMATE.Refillable = true;

                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                    IGR_BAL_CURR_ESTIMATE.SetCellValue(c, mIDX_CHECK_YN, "N");
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
            IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_BAL_CURR_ESTIMATE.Refillable = true;

            IDA_BAL_CURR_ESTIMATE.Fill();
        }

        private void BTN_ESTIMATE_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CheckData() == false)
            {
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int mIDX_CHECK_YN = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CHECK_YN");
            int mIDX_BALANCE_STATEMENT_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("BALANCE_STATEMENT_ID");
            int mIDX_ACCOUNT_CONTROL_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            int mIDX_CURRENCY_CODE = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("CURRENCY_CODE");
            int mIDX_ITEM_GROUP_ID = IGR_BAL_CURR_ESTIMATE.GetColumnToIndex("ITEM_GROUP_ID");
            
            string mSTATUS = "F";
            string mMESSAGE = null;

            for (int c = 0; c < IGR_BAL_CURR_ESTIMATE.RowCount; c++)
            {
                if (iConv.ISNull(IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CHECK_YN)) == "Y")
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    IGR_BAL_CURR_ESTIMATE.CurrentCellMoveTo(c, mIDX_CHECK_YN);
                    IGR_BAL_CURR_ESTIMATE.CurrentCellActivate(c, mIDX_CHECK_YN);

                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_DATE", V_BALANCE_DATE.EditValue);
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_BALANCE_STATEMENT_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_BALANCE_STATEMENT_ID));
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_ACCOUNT_CONTROL_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ACCOUNT_CONTROL_ID));
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_CURRENCY_CODE", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_CURRENCY_CODE));
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.SetCommandParamValue("P_ITEM_GROUP_ID", IGR_BAL_CURR_ESTIMATE.GetCellValue(c, mIDX_ITEM_GROUP_ID));
                    IDC_CANCEL_CLOSED_CURR_ESTIMATE.ExecuteNonQuery();
                    mSTATUS = iConv.ISNull(IDC_CANCEL_CLOSED_CURR_ESTIMATE.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iConv.ISNull(IDC_CANCEL_CLOSED_CURR_ESTIMATE.GetCommandParamValue("O_MESSAGE"));

                    if (mSTATUS == "F")
                    {
                        IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
                        IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
                        IDA_BAL_CURR_ESTIMATE.Refillable = true;

                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                    IGR_BAL_CURR_ESTIMATE.SetCellValue(c, mIDX_CHECK_YN, "N");
                }
            }
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IGR_BAL_CURR_ESTIMATE.LastConfirmChanges();
            IDA_BAL_CURR_ESTIMATE.OraSelectData.AcceptChanges();
            IDA_BAL_CURR_ESTIMATE.Refillable = true;

            IDA_BAL_CURR_ESTIMATE.Fill();
        }

        private void BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }
         
        private void CB_SELECT_ALL_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Select_Check_YN(IGR_BAL_CURR_ESTIMATE, CB_SELECT_ALL.CheckBoxValue);
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaVENDOR_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----
         
        private void IDA_BALANCE_ACCOUNT_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Search_Balance_Statement(-1);
                return;
            }
            CB_SELECT_ALL.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            Search_Balance_Statement(pBindingManager.DataRow["ACCOUNT_CONTROL_ID"]);
        }

        #endregion

    }
}