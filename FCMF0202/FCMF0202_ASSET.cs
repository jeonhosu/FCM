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

namespace FCMF0202
{    
    public partial class FCMF0202_ASSET : Office2007Form
    {                
        #region ----- Variables -----
     
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        object mSLIP_LINE_ID;

        #endregion;

        public object Get_SUP_AMOUNT
        {
            get
            {
                return SUM_SUPPLY_AMOUNT.EditValue;
            }
        }

        public object Get_VAT_AMOUNT
        {
            get
            {
                return SUM_VAT_AMOUNT.EditValue;
            }
        }

        public object Get_ASSET_COUNT
        {
            get
            {
                return SUM_COUNT.EditValue;
            }
        }

        public InfoSummit.Win.ControlAdv.ISDataAdapter IDA_ADAPTER
        {
            get
            {
                return idaDPR_ASSET;
            }
        }

        #region ----- Constructor -----

        public FCMF0202_ASSET(ISAppInterface pAppInterface, object pSLIP_LINE_ID, object pSUPPLY_AMOUNT, object pVAT_AMOUNT)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSLIP_LINE_ID = pSLIP_LINE_ID;
            SUPPLY_AMOUNT.EditValue = pSUPPLY_AMOUNT;
            VAT_AMOUNT.EditValue = pVAT_AMOUNT;
        }

        public FCMF0202_ASSET(ISAppInterface pAppInterface, ref InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSUPPLY_AMOUNT, object pVAT_AMOUNT)
        {
            InitializeComponent();            
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            igrDPR_SPEC.DataAdapter = pAdapter;

            SUPPLY_AMOUNT.EditValue = pSUPPLY_AMOUNT;
            VAT_AMOUNT.EditValue = pVAT_AMOUNT;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            //idaDPR_ASSET.SetSelectParamValue("W_SLIP_LINE_ID", 1);//mSLIP_LINE_ID);
            idaDPR_ASSET.Fill();

            igrDPR_SPEC.CurrentCellMoveTo(1);
            igrDPR_SPEC.CurrentCellActivate(1);
            igrDPR_SPEC.Focus();
        }

        private void Init_SUM_AMOUNT()
        {
            decimal mSUPPLY_AMOUNT = 0;
            decimal mVAT_AMOUNT = 0;
            decimal mCOUNT = 0;
            int mIDX_SUPPLY_AMOUNT = igrDPR_SPEC.GetColumnToIndex("SUPPLY_AMOUNT");
            int mIDX_VAT_AMOUNT = igrDPR_SPEC.GetColumnToIndex("VAT_AMOUNT");
            int mIDX_COUNT = igrDPR_SPEC.GetColumnToIndex("ASSET_COUNT");
            for (int r = 0; r <= igrDPR_SPEC.RowCount; r++)
            {
                mSUPPLY_AMOUNT = mSUPPLY_AMOUNT + iString.ISDecimaltoZero(igrDPR_SPEC.GetCellValue(r, mIDX_SUPPLY_AMOUNT));
                mVAT_AMOUNT = mVAT_AMOUNT + iString.ISDecimaltoZero(igrDPR_SPEC.GetCellValue(r, mIDX_VAT_AMOUNT));
                mCOUNT = mCOUNT + iString.ISDecimaltoZero(igrDPR_SPEC.GetCellValue(r, mIDX_COUNT));
            }
            SUM_SUPPLY_AMOUNT.EditValue = mSUPPLY_AMOUNT;
            SUM_VAT_AMOUNT.EditValue = mVAT_AMOUNT;
            SUM_COUNT.EditValue = mCOUNT;
        }

        //private void CheckData()
        //{
        //    if (iString.ISNull(BILL_NUM.EditValue) == string.Empty)
        //    {// 어음번호
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10142"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        BILL_NUM.Focus();
        //        return;
        //    }
        //    if (iString.ISNull(BILL_TYPE.EditValue) == string.Empty)
        //    {// 어음종류
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10143"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        BILL_TYPE_NAME.Focus();
        //        return;
        //    }
        //    if (iString.ISNull(CUSTOMER_ID.EditValue) == string.Empty)
        //    {// 고객정보
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10135"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        CUSTOMER_NAME.Focus();                
        //        return;
        //    }
        //    if (iString.ISNull(ISSUE_DATE.EditValue) == string.Empty)
        //    {// 발행일자
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10144"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        ISSUE_DATE.Focus();
        //        return;
        //    }
        //    if (iString.ISNull(DUE_DATE.EditValue) == string.Empty)
        //    {// 만기일자
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10145"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        ISSUE_DATE.Focus();
        //        return;
        //    }
        //    if (iString.ISNumtoZero(BILL_AMOUNT.EditValue) == 0)
        //    {// 어음금액
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10146"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        BILL_AMOUNT.Focus();
        //        return;
        //    }
        //    if (iString.ISNull(RECEIPT_DATE.EditValue) == string.Empty)
        //    {// 입금일자
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10147"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        ISSUE_DATE.Focus();
        //        return;
        //    }
        //}

        //private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        //{
        //    ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
        //    ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        //}

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

        #region ----- From Event -----
        
        private void FCMF0202_ASSET_Load(object sender, EventArgs e)
        {
            idaDPR_ASSET.FillSchema();
        }

        private void FCMF0202_ASSET_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }

        private void FCMF0202_ASSET_FormClosed(object sender, FormClosedEventArgs e)
        {
           this.DialogResult = DialogResult.OK;
        }

        private void ibtnSEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void ibtnSAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SUPPLY_AMOUNT.Focus();
            Init_SUM_AMOUNT();
           // idaDPR_ASSET.Update();
        }

        private void ibtnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDPR_ASSET.Cancel();
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----
        
        #endregion

        #region ----- Adapter Event -----
        
        private void idaDPR_SPEC_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //if (iString.ISNull(e.Row["SLIP_LINE_ID"]) == string.Empty)
            //{// 전표라인
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10271"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["DPR_ASSET_GB_ID"]) == string.Empty)
            {// 자산구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Asset Type(자산구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDPR_SPEC_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaDPR_SPEC_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_SUM_AMOUNT();
        }

        private void idaDPR_SPEC_UpdateCompleted(object pSender)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        #endregion

        private void igrDPR_SPEC_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {

        }

    }
}