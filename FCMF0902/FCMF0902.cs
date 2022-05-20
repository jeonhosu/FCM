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

namespace FCMF0902
{
    public partial class FCMF0902 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ----- Constructor -----

        public FCMF0902()
        {
            InitializeComponent();
        }

        public FCMF0902(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----         
        
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

        private void SEARCH_DB()
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            IDA_YEAR_REPLACE_ACC.Fill();
            IDA_SELECT_YEAR_REPLACE_SUM.Fill();

            IGR_YEAR_REPLACE_ACC.Focus();
        }

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            //ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            //ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Init_Sub_Panel(bool pShow_Flag, string pSub_Panel)
        {
            if (mSUB_SHOW_FLAG == true && pShow_Flag == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pShow_Flag == true)
            {
                try
                {
                    if (pSub_Panel == "CREATE_SLIP")
                    {
                        V_ERR_MSG_1.EditValue = string.Empty;
                        V_ERR_MSG_2.EditValue = string.Empty;

                        GB_CREATE_SLIP.Left = 93;
                        GB_CREATE_SLIP.Top = 200;
                         
                        GB_CREATE_SLIP.Width = 790;
                        GB_CREATE_SLIP.Height = 180; 

                        GB_CREATE_SLIP.Border3DStyle = Border3DStyle.Bump;
                        GB_CREATE_SLIP.BorderStyle = BorderStyle.Fixed3D;

                        GB_CREATE_SLIP.BringToFront();
                        GB_CREATE_SLIP.Visible = true;
                    } 
                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                
                GB_INQURIY_CONDITION.Enabled = false;
                GB_BUTTON.Enabled = false;
                GB_GRID.Enabled = false;
                GB_SUM.Enabled = false; 
            }
            else
            {
                try
                {
                    if (pSub_Panel == "ALL")
                    {
                        GB_CREATE_SLIP.Visible = false; 
                    }
                    else if (pSub_Panel == "CREATE_SLIP")
                    {
                        GB_CREATE_SLIP.Visible = false;
                    }  
                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }

                GB_INQURIY_CONDITION.Enabled = true;
                GB_BUTTON.Enabled = true;
                GB_GRID.Enabled = true;
                GB_SUM.Enabled = true; 
            }
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
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

        #region ----- Form Event -----
        
        private void FCMF0902_Load(object sender, EventArgs e)
        {
            
        }

        private void FCMF0902_Shown(object sender, EventArgs e)
        {
            Init_Sub_Panel(false, "ALL");
            W_CLOSED_YEAR.EditValue = iDate.ISYear(System.DateTime.Today); 
        }
          
        #endregion

        #region ----- Button Event -----

        private void BTN_CREATE_BALANCE_AMT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            string vSTATUS = null;
            string vMESSAGE = null;

            //전표생성 여부 체크//
            IDC_GET_SLIP_INTERFACE_FLAG_P.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_GET_SLIP_INTERFACE_FLAG_P.GetCommandParamValue("O_INTERFACE_FLAG"));
            if (vSTATUS == "Y")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10452"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //처리여부 묻기//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            isDataTransaction1.BeginTran();

            try
            {
                IDC_SET_YEAR_REPLACE.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_SET_YEAR_REPLACE.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_YEAR_REPLACE.GetCommandParamValue("O_MESSAGE"));
                if (vSTATUS == "F")
                {
                    isDataTransaction1.RollBack();
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    return;
                }
            }
            catch (System.Exception ex)
            {
                isDataTransaction1.RollBack();
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
            isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void C_BTN_EXECUTE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            int vERR_CNT = 0;
            string vSTATUS = null;
            string vMESSAGE = null; 

            //전표생성 여부 체크//
            IDC_GET_SLIP_INTERFACE_FLAG_P.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_GET_SLIP_INTERFACE_FLAG_P.GetCommandParamValue("O_INTERFACE_FLAG"));
            if (vSTATUS == "Y")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10452"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //처리여부 묻기//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
             
            isDataTransaction1.BeginTran(); 
            if (CB_CF_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                //1.손익잔액대체.
                try
                {
                    IDC_SET_YEAR_REPLACE.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_YEAR_REPLACE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_YEAR_REPLACE.GetCommandParamValue("O_MESSAGE"));
                    if (vSTATUS == "F")
                    {
                        isDataTransaction1.RollBack(); 
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();
                        V_ERR_MSG_CR.EditValue = vMESSAGE;
                        return;
                    }
                    IDC_SET_YEAR_REPLACE.DataTransaction = null;
                }
                catch (System.Exception ex)
                {
                    vERR_CNT++;
                    isDataTransaction1.RollBack(); 
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    V_ERR_MSG_CR.EditValue = ex.Message;
                    System.Windows.Forms.Application.DoEvents();
                }
                CB_CF_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                V_ERR_MSG_CR.EditValue = "OK";
            }

            if (CB_YR_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                //1.손익잔액대체.
                try
                {
                    IDC_SET_SLIP_YEAR_REPLACE.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_SLIP_YEAR_REPLACE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_SLIP_YEAR_REPLACE.GetCommandParamValue("O_MESSAGE"));
                    if (vSTATUS == "F")
                    {
                        vERR_CNT++;
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        V_ERR_MSG_1.EditValue = vMESSAGE;
                        return;
                    }
                }
                catch (System.Exception ex)
                {
                    vERR_CNT++;
                    isDataTransaction1.RollBack();
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    V_ERR_MSG_1.EditValue = ex.Message;
                    System.Windows.Forms.Application.DoEvents();
                }
                CB_YR_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                V_ERR_MSG_1.EditValue = "OK";
            }

            if (CB_ES_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                //2.미처분이익잉여금 차기년도 이월.
                try
                {
                    IDC_SET_SLIP_EARNED_SURPLUS.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_SLIP_EARNED_SURPLUS.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_SLIP_EARNED_SURPLUS.GetCommandParamValue("O_MESSAGE"));
                    if (vSTATUS == "F")
                    {
                        vERR_CNT++;
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();
                        V_ERR_MSG_2.EditValue = vMESSAGE;
                        return;
                    }
                }
                catch (System.Exception ex)
                {
                    vERR_CNT++;
                    isDataTransaction1.RollBack();
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    V_ERR_MSG_2.EditValue = ex.Message;
                    System.Windows.Forms.Application.DoEvents();
                }
                CB_ES_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                V_ERR_MSG_2.EditValue = "OK";
            }
            isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (vERR_CNT == 0)
            {
                Init_Sub_Panel(false, "CREATE_SLIP");
            }
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void C_BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Init_Sub_Panel(false, "CREATE_SLIP");
        }

        private void BTN_CREATE_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            CB_YR_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
            CB_ES_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
            
            Init_Sub_Panel(true, "CREATE_SLIP");
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Check_Inquiry_Condition() == false)
            {
                return;
            }

            string vSTATUS = null;
            string vMESSAGE = null;

            //전표생성 여부 체크//
            IDC_GET_SLIP_INTERFACE_FLAG_P.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_GET_SLIP_INTERFACE_FLAG_P.GetCommandParamValue("O_INTERFACE_FLAG"));
            if (vSTATUS == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10426"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //처리여부 묻기//
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10436"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            try
            {
                IDC_CANCEL_SLIP_YEAR_REPLACE.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_CANCEL_SLIP_YEAR_REPLACE.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_CANCEL_SLIP_YEAR_REPLACE.GetCommandParamValue("O_MESSAGE"));
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

        private void ILA_W_FORM_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FS_TYPE", "Y");
        }

        #endregion


        #region ----- Adapter Event -----
         
        
        #endregion


    }
}