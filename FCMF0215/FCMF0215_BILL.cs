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

namespace FCMF0215
{
    public partial class FCMF0215_BILL : Office2007Form
    {                
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        // 전표에서 전달받는 값.
        object mDEPT_ID = null;
        object mDEPT_NAME = null;
        object mBILL_CLASS = null; 
        object mBILL_NUM = null;        
        object mBILL_AMOUNT = null;
        object mVENDOR_CODE = null;
        object mBANK_CODE = null;
        object mVAT_ISSUE_DATE = null;
        object mISSUE_DATE = null;
        object mDUE_DATE = null;
        object mMANAGEMENT_PERSON_ID = null;
        object mMANAGEMENT_PERSON_NAME = null;

        public object Get_BILL_AMOUNT
        {// 어음금액.
            get
            {
                return BILL_AMOUNT.EditValue;
            }
        }

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
        public object Get_VAT_ISSUE_DATE
        {//세금계산서발행일.
            get
            {
                return VAT_ISSUE_DATE.EditValue;
            }
        }

        public object Get_ISSUE_DATE
        {//발행일자.
            get
            {
                return ISSUE_DATE.EditValue;
            }
        }

        public object Get_DUE_DATE
        {//만기일자.
            get
            {
                return DUE_DATE.EditValue;
            }
        }

        public object Get_BILL_NUM
        {//어음번호.
            get
            {
                return BILL_NUM.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public FCMF0215_BILL(ISAppInterface pAppInterface, object pDEPT_ID, object pDEPT_NAME
                                    , object pBILL_CLASS, object pBILL_NUM, object pBILL_AMOUNT
                                    , object pVENDOR_CODE, object pBANK_CODE
                                    , object pVAT_ISSUE_DATE, object pISSUE_DATE, object pDUE_DATE
                                    , object pMANAGEMENT_PERSON_ID, object pMANAGEMENT_PERSON_NAME)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mDEPT_ID = pDEPT_ID;
            mDEPT_NAME = pDEPT_NAME;
            mBILL_CLASS = pBILL_CLASS;
            mBILL_NUM = pBILL_NUM;
            mBILL_AMOUNT = pBILL_AMOUNT;
            mVENDOR_CODE = pVENDOR_CODE;
            mBANK_CODE = pBANK_CODE;
            mVAT_ISSUE_DATE = pVAT_ISSUE_DATE;
            mISSUE_DATE = pISSUE_DATE;
            mDUE_DATE = pDUE_DATE;
            mMANAGEMENT_PERSON_ID = pMANAGEMENT_PERSON_ID;
            mMANAGEMENT_PERSON_NAME = pMANAGEMENT_PERSON_NAME;
        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            BTN_GET_BILL_NUM.Enabled = false;

            idaBILL_MASTER.Fill();
            BILL_NUM_0.Focus();
        }

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Insert()
        {
            // 어음상태.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_STATUS");
            idcDV_COMMON.ExecuteNonQuery();
            BILL_STATUS_NAME.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            BILL_STATUS.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");

            //자타구분.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_MODE");
            idcDV_COMMON.ExecuteNonQuery();
            BILL_MODE_DESC.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            BILL_MODE.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");
            
            ////어음구분
            //idcCOMMON_CODE_NAME.SetCommandParamValue("W_GROUP_CODE", "BILL_TYPE");
            //idcCOMMON_CODE_NAME.SetCommandParamValue("W_CODE", mBILL_CLASS);
            //idcCOMMON_CODE_NAME.ExecuteNonQuery();
            //BILL_TYPE_NAME.EditValue = idcCOMMON_CODE_NAME.GetCommandParamValue("O_RETURN_VALUE");
            //if (iString.ISNull(BILL_TYPE_NAME.EditValue) == string.Empty)
            //{
            //    //DEFAULT VALUE 설정 : 어음구분.
            //    idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_TYPE");
            //    idcDV_COMMON.ExecuteNonQuery();
            //    BILL_TYPE_NAME.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            //    BILL_TYPE.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");
            //}

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

            MANAGEMENT_PERSON_ID.EditValue = mMANAGEMENT_PERSON_ID;
            MANAGEMENT_PERSON_NAME.EditValue = mMANAGEMENT_PERSON_NAME;

            //전표값 설정.
            KEEP_DEPT_NAME.EditValue = mDEPT_NAME;
            KEEP_DEPT_ID.EditValue = mDEPT_ID;

            RECEIPT_DEPT_NAME.EditValue = mDEPT_NAME;
            RECEIPT_DEPT_ID.EditValue = mDEPT_ID;
 
            BILL_NUM.EditValue = mBILL_NUM;
            BILL_AMOUNT.EditValue = mBILL_AMOUNT;
            VENDOR_CODE.EditValue = mVENDOR_CODE;
            BANK_CODE.EditValue = mBANK_CODE;

            VAT_ISSUE_DATE.EditValue = mVAT_ISSUE_DATE;
            ISSUE_DATE.EditValue = mISSUE_DATE;
            DUE_DATE.EditValue = mDUE_DATE;

            BILL_AMOUNT.EditValue = 0;

            BTN_GET_BILL_NUM.Enabled = true;
            BILL_NUM.Focus();
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

        private void FCMF0215_BILL_Load(object sender, EventArgs e)
        {
            idaBILL_MASTER.FillSchema();
        }

        private void FCMF0215_BILL_Shown(object sender, EventArgs e)
        {
            V_BILL_CLASS.EditValue = mBILL_CLASS;
            BILL_NUM_0.EditValue = mBILL_NUM;

            BTN_GET_BILL_NUM.Enabled = false;
            SEARCH_DB();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void ADD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaBILL_MASTER.AddUnder();
            Set_Insert();
        }

        private void CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            BTN_GET_BILL_NUM.Enabled = false;
            idaBILL_MASTER.Cancel();
        }

        private void ibtnSAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            BTN_GET_BILL_NUM.Enabled = false;
            idaBILL_MASTER.Update();
            BILL_NUM.Focus();
        }

        private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaBILL_MASTER.Update();
            if (idaBILL_MASTER.CurrentRow.RowState == DataRowState.Unchanged)
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
        
        private void BILL_NUM_0_KeyUp(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && e.Shift == true)
            {
                idaBILL_MASTER.Update();
                if (idaBILL_MASTER.CurrentRow.RowState == DataRowState.Unchanged)
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
            }
        }

        private void BTN_GET_BILL_NUM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (idaBILL_MASTER.CurrentRow.RowState == DataRowState.Added)
            {
                IDC_GET_BILL_NUM_P.ExecuteNonQuery();
            }
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaBILL_NUM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBILL_NUM.SetLookupParamValue("W_BILL_NUM", BILL_NUM_0.EditValue);
            ildBILL_NUM.SetLookupParamValue("W_BILL_CLASS", mBILL_CLASS);
            ildBILL_NUM.SetLookupParamValue("W_BILL_STATUS", "10");
        }

        private void ilaBILL_NUM_0_SelectedRowData(object pSender)
        {
            SEARCH_DB();
        }

        private void ilaBILL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON_W.SetLookupParamValue("W_GROUP_CODE", "BILL_TYPE");
            ILD_COMMON_W.SetLookupParamValue("W_WHERE", string.Format("VALUE1 = {0}", V_BILL_CLASS.EditValue));
            ILD_COMMON_W.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ilaVENDOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBILL_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_STATUS", "Y");
        }

        private void ilaRECEIPT_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaKEEP_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBILL_MODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_MODE", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_START_DATE", ISSUE_DATE.EditValue);
            ildPERSON.SetLookupParamValue("W_END_DATE", ISSUE_DATE.EditValue);
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBILL_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BILL_NUM"]) == string.Empty)
            {// 어음번호
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10142"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BILL_NUM.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BILL_TYPE"]) == string.Empty)
            {// 어음종류
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10143"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BILL_TYPE_NAME.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VENDOR_ID"]) == string.Empty)
            {// 고객정보
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10135"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                VENDOR_NAME.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ISSUE_DATE"]) == string.Empty)
            {// 발행일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10144"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ISSUE_DATE.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DUE_DATE"]) == string.Empty)
            {// 만기일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10145"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ISSUE_DATE.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["BILL_AMOUNT"]) == 0)
            {// 어음금액
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10146"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BILL_AMOUNT.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BILL_MODE"]) == string.Empty)
            {// 자타구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10353"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BILL_MODE_DESC.Focus();
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}