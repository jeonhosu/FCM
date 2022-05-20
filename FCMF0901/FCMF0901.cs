using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;

using InfoSummit.Win.ControlAdv;

namespace FCMF0901
{
    public partial class FCMF0901 : Office2007Form
    {
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();
        private ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0901()
        {
            InitializeComponent();
        }

        public FCMF0901(Form pMainForm, ISAppInterface pAppInterface)
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

            if (TB_MAIN.SelectedTab.TabIndex == TP_ACC.TabIndex)
            {
                IDA_BALANCE_FORWARD_ACC.Fill();
                IGR_BALANCE_FORWARD_ACC.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_MGT.TabIndex)
            {
                IDA_BALANCE_FORWARD_MGT.Fill();
                IGR_BALANCE_FORWARD_MGT.Focus();
            }
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

        private void FCMF0901_Load(object sender, EventArgs e)
        {
            W_CLOSED_YEAR.EditValue = iDate.ISYear(System.DateTime.Today);
        }

        private void FCMF0901_Shown(object sender, EventArgs e)
        {

        }

        #endregion

        #region ----- Button Event -----

        private void BTN_CREATE_FORWARD_AMT_ButtonClick(object pSender, EventArgs pEventArgs)
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

            try
            {               
                IDC_SET_BALANCE_FORWARD.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_SET_BALANCE_FORWARD.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_BALANCE_FORWARD.GetCommandParamValue("O_MESSAGE"));
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

        private void BTN_CREATE_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
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

            try
            {
                IDC_SET_SLIP_BALANCE_FORWARD.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_SET_SLIP_BALANCE_FORWARD.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_SLIP_BALANCE_FORWARD.GetCommandParamValue("O_MESSAGE"));
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
                IDC_CANCEL_SLIP_BALANCE_FORWARD.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_CANCEL_SLIP_BALANCE_FORWARD.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_CANCEL_SLIP_BALANCE_FORWARD.GetCommandParamValue("O_MESSAGE"));
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
        
        private void ilaACCOUNT_CONTROL_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion;

    }
}