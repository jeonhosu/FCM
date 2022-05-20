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
    public partial class FCMF0203_ITEM_DEAL : Office2007Form
    {                
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        // 전표에서 전달받는 값.
        object mISSUE_NUM = null;
        object mVENDOR_CODE = null;
        object mBANK_CODE = null;
        object mISSUE_DATE = null;
        object mCURRENCY_CODE = null;
        
        public object Get_VENDOR_CODE
        {// 거래처코드
            get
            {
                return VENDOR_CODE.EditValue;
            }
        }

        public object Get_VENDOR_NAME
        {//거래처명
            get
            {
                return VENDOR_NAME.EditValue;
            }
        }

        public object Get_BANK_CODE
        {//발행은행코드.
            get
            {
                return BANK_CODE.EditValue;
            }
        }

        public object Get_BANK_NAME
        {//발행은행명
            get
            {
                return BANK_NAME.EditValue;
            }
        }
        
        public object Get_ISSUE_DATE
        {//발행일자.
            get
            {
                return ISSUE_DATE.EditValue;
            }
        }

        public object Get_CURRENCY_CODE
        {//통화
            get
            {
                return CURRENCY_CODE.EditValue;
            }
        }

        public object Get_ISSUE_NUM
        {//발급번호.
            get
            {
                return ISSUE_NUM.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public FCMF0203_ITEM_DEAL(ISAppInterface pAppInterface, object pISSUE_NUM, object pCURRENCY_CODE
                                    , object pVENDOR_CODE, object pBANK_CODE, object pISSUE_DATE)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mISSUE_NUM = pISSUE_NUM;
            mCURRENCY_CODE = pCURRENCY_CODE;
            mVENDOR_CODE = pVENDOR_CODE;
            mBANK_CODE = pBANK_CODE;
            mISSUE_DATE = pISSUE_DATE;            
        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            idaDEAL_CONFIRMATION.Fill();
            if (idaDEAL_CONFIRMATION.CurrentRows.Count == 0)
            {
                ISSUE_NUM_0.Focus();
            }
            else
            {
                ISSUE_NUM.Focus();
            }
        }

        private void Set_Insert()
        {            
            //거래처
            idcVENDOR.SetCommandParamValue("W_VENDOR_CODE", mVENDOR_CODE);
            idcVENDOR.ExecuteNonQuery();
            VENDOR_NAME.EditValue = idcVENDOR.GetCommandParamValue("O_VENDOR_NAME");
            VENDOR_ID.EditValue = idcVENDOR.GetCommandParamValue("O_VENDOR_ID");
            TAX_REG_NO.EditValue = idcVENDOR.GetCommandParamValue("O_TAX_REG_NO");

            //발행은행.
            idcBANK.SetCommandParamValue("W_BANK_CODE", mBANK_CODE);
            idcBANK.ExecuteNonQuery();
            BANK_NAME.EditValue = idcBANK.GetCommandParamValue("O_BANK_NAME");
            BANK_ID.EditValue = idcBANK.GetCommandParamValue("O_BANK_ID");

            VENDOR_CODE.EditValue = mVENDOR_CODE;
            BANK_CODE.EditValue = mBANK_CODE;
            ISSUE_DATE.EditValue = mISSUE_DATE;
            CURRENCY_CODE.EditValue = mCURRENCY_CODE;
            
            ISSUE_NUM.Focus();
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

        private void FCMF0203_ITEM_DEAL_Load(object sender, EventArgs e)
        {
            idaDEAL_CONFIRMATION.FillSchema();
        }

        private void FCMF0203_ITEM_DEAL_Shown(object sender, EventArgs e)
        {
            ISSUE_NUM_0.EditValue = mISSUE_NUM;
            SEARCH_DB();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void ADD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDEAL_CONFIRMATION.AddUnder();
            Set_Insert();
        }

        private void CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDEAL_CONFIRMATION.Cancel();
        }

        private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDEAL_CONFIRMATION.Update();
            if (idaDEAL_CONFIRMATION.CurrentRow.RowState == DataRowState.Unchanged)
            {                
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaISSUE_NUM_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        private void ilaVENDOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaVENDOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBANK_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        
        #endregion

        #region ----- Adapter Event -----

        private void idaDEAL_CONFIRMATION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ISSUE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10361"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VENDOR_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ISSUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10362"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BANK_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10200"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

        
    }
}