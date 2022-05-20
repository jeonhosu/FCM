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

namespace FCMF0992
{
    public partial class FCMF0992 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        #endregion;

        #region ----- Constructor -----

        public FCMF0992()
        {
            InitializeComponent();
        }

        public FCMF0992(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Set_Default_Values()
        {
            IDC_GET_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TB_ISSUE_STATUS");
            IDC_GET_DEFAULT_VALUE.ExecuteNonQuery();
            W_TB_ISSUE_STATUS_NAME.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_TB_ISSUE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
        }

        private void Search_DB()
        {
            IGR_TAX_BILL_ISSUE.LastConfirmChanges();
            IDA_TAX_BILL_ISSUE.OraSelectData.AcceptChanges();
            IDA_TAX_BILL_ISSUE.Refillable = true;

            IDA_TAX_BILL_ISSUE.Fill();
            IGR_TAX_BILL_ISSUE.Focus();
        }

        private void Init_Insert()
        {
            IGR_TAX_BILL_ISSUE.SetCellValue("ENABLED_FLAG", "Y");
            IGR_TAX_BILL_ISSUE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));

            IGR_TAX_BILL_ISSUE.Focus();
        }

        private decimal Sync_VAT_Amount(object pVAT_TAX_TYPE, object pSUPPLY_AMOUNT)
        {
            decimal vVAT_AMOUNT =0;

            IDC_VAT_AMT_P.SetCommandParamValue("W_VAT_TAX_TYPE", pVAT_TAX_TYPE);
            IDC_VAT_AMT_P.SetCommandParamValue("W_SUPPLY_AMT", pSUPPLY_AMOUNT);
            IDC_VAT_AMT_P.ExecuteNonQuery();
            vVAT_AMOUNT = iConv.ISDecimaltoZero(IDC_VAT_AMT_P.GetCommandParamValue("O_VAT_AMT"));
            return vVAT_AMOUNT;
        }

        private bool Show_Detail(object pTAX_BILL_ISSUE_NO)
        {
            string vStatus = "UPDATE";
            if(iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue("HOMETAX_ISSUE_STATUS")) == "2") 
            {
                vStatus = "INQUIRY";
            }

            FCMF0992_DETAIL vFCMF0992_DETAIL = new FCMF0992_DETAIL(this.MdiParent, isAppInterfaceAdv1.AppInterface, vStatus, pTAX_BILL_ISSUE_NO);
            vFCMF0992_DETAIL.ShowDialog();
            vFCMF0992_DETAIL.Dispose();
            return true;
        }

        private bool Show_Fix(object pTAX_BILL_ISSUE_NO, object pTAX_BILL_NO, object pHOMETAX_ISSUE_NO)
        {
            if (iConv.ISNull(pHOMETAX_ISSUE_NO) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10176"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string vTax_Bill_Issue_NO = string.Empty;
            string vSRC_Tax_Bill_Issue_No = string.Empty; 
            DialogResult vdlgResult = DialogResult.None;
            FCMF0992_FIX vFCMF0992_FIX = new FCMF0992_FIX(this.MdiParent, isAppInterfaceAdv1.AppInterface, "FIX",
                                                            pTAX_BILL_ISSUE_NO, pTAX_BILL_NO, pHOMETAX_ISSUE_NO);
            vdlgResult = vFCMF0992_FIX.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            {
                vTax_Bill_Issue_NO = vFCMF0992_FIX.New_Tax_Bill_Issue_NO;
                vSRC_Tax_Bill_Issue_No = vFCMF0992_FIX.SRC_Tax_Bill_Issue_NO; 
            }
            vFCMF0992_FIX.Dispose();
            if (vdlgResult == DialogResult.No)
            {
                return false; 
            }

            vdlgResult = DialogResult.None;
            if (vTax_Bill_Issue_NO != string.Empty)
            {
                //수정세금계산서 조회//
                FCMF0992_DETAIL vFCMF0992_DETAIL = new FCMF0992_DETAIL(this.MdiParent, isAppInterfaceAdv1.AppInterface, "FIX", vTax_Bill_Issue_NO);
                vdlgResult = vFCMF0992_DETAIL.ShowDialog();
                vFCMF0992_DETAIL.Dispose();

                if (vdlgResult == DialogResult.OK)
                {
                    Search_DB();
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void Init_BTN_EXEC_FIX(bool vStatus)
        {
            if (vStatus == false)
            {
                BTN_EXEC_FIX.Enabled = false;
            }
            else
            {
                BTN_EXEC_FIX.Enabled = true;
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
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE.AddOver();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE.AddUnder();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_TAX_BILL_ISSUE.Update();
                    }
                    catch (Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.Cancel();
                        IDA_TAX_BILL_ISSUE.Cancel();
                    }
                    else if (IDA_TAX_BILL_ISSUE_LINE.IsFocused)
                    {
                        IDA_TAX_BILL_ISSUE_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_TAX_BILL_ISSUE.IsFocused)
                    {
                        object vTAX_BILL_ISSUE_NO = IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_NO");
                        if (MessageBoxAdv.Show(string.Format("[Tax Bill Issue No :: {0}] {1}", vTAX_BILL_ISSUE_NO, isMessageAdapter1.ReturnText("FCM_10525")), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }

                        IDC_DELETE_TAX_BILL.SetCommandParamValue("P_TAX_BILL_ISSUE_NO", vTAX_BILL_ISSUE_NO);
                        IDC_DELETE_TAX_BILL.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_TAX_BILL.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_TAX_BILL.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        Search_DB();
                    } 
                }
            }
        }

        #endregion;

        #region ----- Form Evevnt -----

        private void FCMF0992_Load(object sender, EventArgs e)
        {
            W_ALL_VIEW_FLAG.BringToFront();
            W_SYNC_TAX_BILL_MASTER.BringToFront();
            W_SYNC_BILL365_STATUS.BringToFront();

            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_START_DATE.EditValue = iDate.ISMonth_1st(W_PERIOD_NAME.EditValue);
            W_END_DATE.EditValue = iDate.ISMonth_Last(W_PERIOD_NAME.EditValue);

            //기본값.
            IDC_GET_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TB_ISSUE_STATUS");
            IDC_GET_DEFAULT_VALUE.ExecuteNonQuery();

            W_TB_ISSUE_STATUS_NAME.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_TB_ISSUE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE");

            //기본값.
            IDC_GET_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TB_HT_ISSUE_STATUS");
            IDC_GET_DEFAULT_VALUE.ExecuteNonQuery();

            W_TB_HT_ISSUE_STATUS_NAME.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_TB_HT_ISSUE_STATUS.EditValue = IDC_GET_DEFAULT_VALUE.GetCommandParamValue("O_CODE"); 

            IDA_TAX_BILL_ISSUE.FillSchema();
            IDA_TAX_BILL_ISSUE_LINE.FillSchema();

            //버튼 상태//
            Init_BTN_EXEC_FIX(false);  //수정세금계산서 발행.
        }

        private void BTN_ISSUE_TAX_BILL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            FCMF0992_DETAIL vFCMF0992_DETAIL = new FCMF0992_DETAIL(this.MdiParent, isAppInterfaceAdv1.AppInterface, "INSERT", string.Empty);
            vFCMF0992_DETAIL.ShowDialog();
            vFCMF0992_DETAIL.Dispose();

        }

        private void BTN_RESEND_EMAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //최종상태가 확인 또는 국세청 전송된 경우 가능.
            if (iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_NO")) == string.Empty)
            {
                return;
            }

            //최종상태가 확인 또는 국세청 전송된 경우 가능.
            if (iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue("HOMETAX_ISSUE_STATUS")) == "0" ||
                iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_STATUS")) == "0" ||
                iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_STATUS")) == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10189"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            FCMF0992_EMAIL vFCMF0992_EMAIL = new FCMF0992_EMAIL(this.MdiParent, isAppInterfaceAdv1.AppInterface,
                                                                IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_NO"),
                                                                IGR_TAX_BILL_ISSUE.GetCellValue("SELL_USER_EMAIL"),
                                                                IGR_TAX_BILL_ISSUE.GetCellValue("BUY_USER_EMAIL"),
                                                                IGR_TAX_BILL_ISSUE.GetCellValue("BUY_USER2_EMAIL"));
            vFCMF0992_EMAIL.ShowDialog();
            vFCMF0992_EMAIL.Dispose();
        }

        private void BTN_EXEC_BILL365_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Questioin", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vCURR_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE")).Date;
            DateTime vISSUE_DATE;

            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            object vDELAY_ISSUE_YN = "N";

            int vIDX_CHECK_YN = IGR_TAX_BILL_ISSUE.GetColumnToIndex("CHECK_YN");
            int vIDX_ISSUE_DATE = IGR_TAX_BILL_ISSUE.GetColumnToIndex("ISSUE_DATE");
            int vIDX_TAX_BILL_ISSUE_NO = IGR_TAX_BILL_ISSUE.GetColumnToIndex("TAX_BILL_ISSUE_NO");
            int vIDX_DELAY_ISSUE_YN = IGR_TAX_BILL_ISSUE.GetColumnToIndex("DELAY_ISSUE_YN");

            for (int r = 0; r < IGR_TAX_BILL_ISSUE.RowCount; r++)
            {
                if (iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                {
                    vDELAY_ISSUE_YN = IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_DELAY_ISSUE_YN);
                    vISSUE_DATE = iDate.ISGetDate(IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_ISSUE_DATE)).Date;
                    if (iDate.ISYearMonth(vCURR_DATE) == iDate.ISYearMonth(vISSUE_DATE))
                    {
                        vDELAY_ISSUE_YN = "N";
                    }
                    else
                    {
                        if (vCURR_DATE <= iDate.ISGetDate(string.Format("{0}-10", iDate.ISYearMonth(vCURR_DATE))) && iDate.ISGetDate(iDate.ISMonth_1st(iDate.ISDate_Month_Add(vCURR_DATE, -1))) <= vISSUE_DATE)
                        {
                            vDELAY_ISSUE_YN = "N";
                        }
                        else
                        {
                            if (iConv.ISNull(vDELAY_ISSUE_YN) == "Y")
                            {
                                //이미 적용되어 있음/
                            }
                            else if (MessageBoxAdv.Show("지연교부 발행대상입니다. 발행하시겠습니까? \r\n 발행시 지연교부세가 발생합니다.", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                vDELAY_ISSUE_YN = "Y";
                            }
                            else
                            {
                                vDELAY_ISSUE_YN = "N";
                            }
                        }
                    }

                    IDC_SET_TRANSFER_BILL365.SetCommandParamValue("W_ISSUE_DATE", vISSUE_DATE);
                    IDC_SET_TRANSFER_BILL365.SetCommandParamValue("W_TAX_BILL_ISSUE_NO", IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_TAX_BILL_ISSUE_NO));
                    IDC_SET_TRANSFER_BILL365.SetCommandParamValue("W_DELAY_ISSUE_YN", vDELAY_ISSUE_YN);
                    IDC_SET_TRANSFER_BILL365.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_TRANSFER_BILL365.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_TRANSFER_BILL365.GetCommandParamValue("O_MESSAGE"));
                    if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BTN_CANCEL_BILL365_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Questioin", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_GET_DATE.ExecuteNonQuery();
            DateTime vCURR_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE")).Date;
            DateTime vISSUE_DATE;

            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty; 

            int vIDX_CHECK_YN = IGR_TAX_BILL_ISSUE.GetColumnToIndex("CHECK_YN");
            int vIDX_ISSUE_DATE = IGR_TAX_BILL_ISSUE.GetColumnToIndex("ISSUE_DATE");
            int vIDX_TAX_BILL_ISSUE_NO = IGR_TAX_BILL_ISSUE.GetColumnToIndex("TAX_BILL_ISSUE_NO"); 

            for (int r = 0; r < IGR_TAX_BILL_ISSUE.RowCount; r++)
            {
                if (iConv.ISNull(IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_CHECK_YN)) == "Y")
                {
                    vISSUE_DATE = iDate.ISGetDate(IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_ISSUE_DATE)).Date;

                    IDC_CANCEL_TRANSFER_BILL365.SetCommandParamValue("W_ISSUE_DATE", vISSUE_DATE);
                    IDC_CANCEL_TRANSFER_BILL365.SetCommandParamValue("W_TAX_BILL_ISSUE_NO", IGR_TAX_BILL_ISSUE.GetCellValue(r, vIDX_TAX_BILL_ISSUE_NO));
                    IDC_CANCEL_TRANSFER_BILL365.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_CANCEL_TRANSFER_BILL365.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_CANCEL_TRANSFER_BILL365.GetCommandParamValue("O_MESSAGE"));
                    if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BTN_EXEC_FIX_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Fix(IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_NO"), 
                        IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_NO"), 
                        IGR_TAX_BILL_ISSUE.GetCellValue("HOMETAX_ISSUE_NO"));
        }

        private void IGR_TAX_BILL_ISSUE_CellDoubleClick(object pSender)
        {
            if (IGR_TAX_BILL_ISSUE.RowIndex < 0)
            {
                return;
            }

            Show_Detail(IGR_TAX_BILL_ISSUE.GetCellValue("TAX_BILL_ISSUE_NO"));
        }

        private void BTN_SYNC_TAX_BILL_MASTER_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //매출 동기화.
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_INIT_TAX_BILL_ISSUE.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_INIT_TAX_BILL_ISSUE.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_INIT_TAX_BILL_ISSUE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            Search_DB();
        }

        private void BTN_SYNC_BILL365_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            
            IDC_INIT_TRANSFER_BILL365.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_INIT_TRANSFER_BILL365.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_INIT_TRANSFER_BILL365.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            
            if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            Search_DB();
        }

        private void IGR_TAX_BILL_ISSUE_LINE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            decimal vVAT_RATE = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE.GetCellValue("VAT_RATE"));
            decimal vQTY = 0;
            decimal vUNIT_PRICE = 0;
            decimal vSUPPLY_AMOUNT = 0;
            decimal vVAT_AMOUNT = 0;
            if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("QTY"))
            {
                vQTY = iConv.ISDecimaltoZero(e.NewValue,1);
                vUNIT_PRICE = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue("UNIT_PRICE"), 0);
                vSUPPLY_AMOUNT = vQTY * vUNIT_PRICE;
                vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("SUPPLY_AMOUNT", vSUPPLY_AMOUNT);
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vSUPPLY_AMOUNT);
            }
            else if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("UNIT_PRICE"))
            {
                vQTY = iConv.ISDecimaltoZero(IGR_TAX_BILL_ISSUE_LINE.GetCellValue("QTY"), 1);
                vUNIT_PRICE = iConv.ISDecimaltoZero(e.NewValue,1);
                vSUPPLY_AMOUNT = vQTY * vUNIT_PRICE;
                vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);

                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("SUPPLY_AMOUNT", vSUPPLY_AMOUNT);
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vSUPPLY_AMOUNT);
            }
            else if (e.ColIndex == IGR_TAX_BILL_ISSUE_LINE.GetColumnToIndex("SUPPLY_AMOUNT"))
            {
                vSUPPLY_AMOUNT = iConv.ISDecimaltoZero(e.NewValue, 0);
                vVAT_AMOUNT = Math.Round(vSUPPLY_AMOUNT * vVAT_RATE);
                 
                IGR_TAX_BILL_ISSUE_LINE.SetCellValue("VAT_AMOUNT", vSUPPLY_AMOUNT);
            }
        }

        private void V_CHECK_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_CHECK_YN = IGR_TAX_BILL_ISSUE.GetColumnToIndex("CHECK_YN");
            for (int vRow = 0; vRow < IGR_TAX_BILL_ISSUE.RowCount; vRow++)
            {
                IGR_TAX_BILL_ISSUE.SetCellValue(vRow, vIDX_CHECK_YN, V1_CHECK_FLAG.CheckBoxString);  
            }
            
            IGR_TAX_BILL_ISSUE.LastConfirmChanges();
            IDA_TAX_BILL_ISSUE.OraSelectData.AcceptChanges();
            IDA_TAX_BILL_ISSUE.Refillable = true;
        }

        #endregion

        #region ----- Lookup Event ------

        private void SetCommon(object pGROUP_CODE, object pENABLED_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pENABLED_YN);
        }

        private void ILA_TB_ISSUE_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_ISSUE_STATUS", "Y");   
        }

        private void ILA_TB_HT_ISSUE_STATUS_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_HT_ISSUE_STATUS", "Y");   
        }

        private void ILA_TB_VAT_TYPE_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_VAT_TYPE", "Y");   
        }

        private void ILA_TB_VAT_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_VAT_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_TB_REQ_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_REQ_TYPE", "Y");   
        }

        private void ILA_HOMETAX_SEND_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommon("TB_HT_SEND_TYPE", "Y");   
        }

        private void ILA_CUSTOMER_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_SUPP_CUST_TYPE", "C");
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_CUSTOMER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_CUSTOMER.SetLookupParamValue("W_SUPP_CUST_TYPE", "C");
            ILD_CUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        
        private void ILA_TB_VAT_TYPE_SelectedRowData(object pSender)
        {
            decimal vVAT_AMOUNT = Sync_VAT_Amount(IGR_TAX_BILL_ISSUE.GetCellValue("VAT_TAX_TYPE"), IGR_TAX_BILL_ISSUE.GetCellValue("SUPPLY_AMOUNT"));
            IGR_TAX_BILL_ISSUE.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);
        }

        #endregion

        #region ----- Adapter Event ------

        private void IDA_TAX_BILL_ISSUE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Init_BTN_EXEC_FIX(false);  //수정세금계산서 발행.
            }
            else
            {
                if (iConv.ISNull(pBindingManager.DataRow["HOMETAX_ISSUE_STATUS"]) == "2")
                {
                    Init_BTN_EXEC_FIX(true);  //수정세금계산서 발행.
                }
                else
                {
                    Init_BTN_EXEC_FIX(false);  //수정세금계산서 발행.
                } 
            }
        }

        private void IDA_TAX_BILL_ISSUE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ISSUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10144"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["VENDOR_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SUPPLY_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["VAT_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10281"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["TAX_BILL_VAT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10614"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BUY_USER_EMAIL"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10615"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void IDA_TAX_BILL_ISSUE_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["WRITE_MM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10616"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WRITE_DD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10617"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["TAX_BILL_ISSUE_NO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:=Tax bill issue no(세금계산서 발행번호")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ITEM_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SUPPLY_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10517"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["VAT_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10281"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["UNIT_PRICE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10618"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        #endregion



    }
}