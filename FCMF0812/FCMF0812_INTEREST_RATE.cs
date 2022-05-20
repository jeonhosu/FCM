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

namespace FCMF0812
{
    public partial class FCMF0812_INTEREST_RATE : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0812_INTEREST_RATE(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IDA_INTEREST_RATE.Fill();
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

        #region ----- Form Event -----

        private void FCMF0812_INTEREST_RATE_Load(object sender, EventArgs e)
        {
            IDA_INTEREST_RATE.FillSchema();
        }

        private void FCMF0812_INTEREST_RATE_Shown(object sender, EventArgs e)
        {
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;

            Search_DB();
        }

        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Search_DB();
        }

        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(STD_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_DATE.Focus();
                return;
            }

            try
            {
                IDC_INTEREST_RATE.ExecuteNonQuery();
                string vSTATUS = iString.ISNull(IDC_INTEREST_RATE.GetCommandParamValue("O_STATUS"));
                string vMESSAGE = iString.ISNull(IDC_INTEREST_RATE.GetCommandParamValue("O_MESSAGE"));

                if (IDC_INTEREST_RATE.ExcuteError)
                {
                    MessageBoxAdv.Show(IDC_INTEREST_RATE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (vSTATUS == "F")
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //입력부 초기화//
            STD_DATE.EditValue = null;
            INTEREST_RATE.EditValue = null;
            DESCRIPTION.EditValue = string.Empty;

            //재조회//
            Search_DB();
        }

        private void BTN_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            try
            {
                IDC_DELETE_INTEREST_RATE.ExecuteNonQuery();
                string vSTATUS = iString.ISNull(IDC_DELETE_INTEREST_RATE.GetCommandParamValue("O_STATUS"));
                string vMESSAGE = iString.ISNull(IDC_DELETE_INTEREST_RATE.GetCommandParamValue("O_MESSAGE"));

                if (IDC_DELETE_INTEREST_RATE.ExcuteError)
                {
                    MessageBoxAdv.Show(IDC_DELETE_INTEREST_RATE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (vSTATUS == "F")
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
              
            //재조회//
            Search_DB();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_INTEREST_RATE.Cancel();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }
        
        #endregion
        
        #region ------ Lookup Event ------

        #endregion

        #region ------ Adapter Event ------

        private void idaINTEREST_RATE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["INTEREST_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10291"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaINTEREST_RATE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion             

    }
}