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
using System.IO;
using Syncfusion.GridExcelConverter;

namespace FCMF0509
{
    public partial class FCMF0509 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCurrency_Code;

        #endregion;

        #region ----- Constructor -----

        public FCMF0509()
        {
            InitializeComponent();
        }

        public FCMF0509(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_LC_LIST.TabIndex)
            {
                IDA_LC_LIST.Fill();
                IGR_LC_LIST.Focus();
            }
            else
            {
                idaLC_MASTER.Fill();
                igrLC_LIST.Focus();
            }
        }

        private void Insert_LC()
        {
            CURRENCY_CODE.EditValue = mCurrency_Code;

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TRANS_STATUS");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            TRANS_TRANS_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            TRANS_STATUS.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            OPEN_AMOUNT.EditValue = 0;
            OPEN_CURR_AMOUNT.EditValue = 0;
            OPEN_EXPENSE_AMOUNT.EditValue = 0;

            LC_NUM.Focus();
        }

        private void Open_Amount()
        {
            decimal pExchange_Rate = iString.ISDecimaltoZero(EXCHANGE_RATE.EditValue, 1);
            decimal pOpen_Curr_Amount = iString.ISDecimaltoZero(OPEN_CURR_AMOUNT.EditValue);
            decimal mOpenAmount = Math.Round(pExchange_Rate * pOpen_Curr_Amount, 2);
            OPEN_AMOUNT.EditValue = mOpenAmount;
        }

        private void Init_Currency_Amount()
        {
            if (iString.ISNull(CURRENCY_CODE.EditValue) == string.Empty || CURRENCY_CODE.EditValue.ToString() == mCurrency_Code.ToString())
            {
                EXCHANGE_RATE.EditValue = DBNull.Value;

                if (iString.ISDecimaltoZero(OPEN_CURR_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    OPEN_CURR_AMOUNT.EditValue = 0;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                OPEN_CURR_AMOUNT.ReadOnly = true;
                OPEN_CURR_AMOUNT.Insertable = false;
                OPEN_CURR_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                OPEN_CURR_AMOUNT.TabStop = false;
            }
            else
            {
                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                OPEN_CURR_AMOUNT.ReadOnly = false;
                OPEN_CURR_AMOUNT.Insertable = true;
                OPEN_CURR_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                OPEN_CURR_AMOUNT.TabStop = true;
            }
            EXCHANGE_RATE.Refresh();
            OPEN_CURR_AMOUNT.Refresh();
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

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

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx vGrid)
        {
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.RestoreDirectory = true;
            vSaveFileDialog.Filter = "Excel file(*.xls)|*.xls";
            vSaveFileDialog.DefaultExt = "xls";

            if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                vExport.GridToExcel(vGrid.BaseGrid, vSaveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = vSaveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaLC_MASTER.IsFocused)
                    {
                        idaLC_MASTER.AddOver();
                        Insert_LC();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaLC_MASTER.IsFocused)
                    {
                        idaLC_MASTER.AddUnder();
                        Insert_LC();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaLC_MASTER.IsFocused)
                    {
                        idaLC_MASTER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaLC_MASTER.IsFocused)
                    {
                        idaLC_MASTER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaLC_MASTER.IsFocused)
                    {
                        if (idaLC_MASTER.CurrentRow.RowState == DataRowState.Added)
                        {
                            idaLC_MASTER.Delete();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_LC_LIST);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0509_Load(object sender, EventArgs e)
        {
            idaLC_MASTER.FillSchema();
        }

        private void FCMF0509_Shown(object sender, EventArgs e)
        {
            IDC_BASE_CURRENCY.ExecuteNonQuery();
            mCurrency_Code = IDC_BASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE");            
        }

        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Open_Amount();
        }

        private void OPEN_CURR_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Open_Amount();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaBANK_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaTRANS_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "TRANS_STATUS");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaSUPPLIER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildSUPPLIER.SetLookupParamValue("W_SUPP_CUST_TYPE", "S");
            ildSUPPLIER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_SelectedRowData(object pSender)
        {
            Init_Currency_Amount();
        }
        
        #endregion

        #region ----- Adapter Event -----

        private void idaLC_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["LC_NUM"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(LC_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            //if (e.Row["TRANS_STATUS"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(TRANS_TRANS_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
            //    e.Cancel = true;
            //    return;
            //}
            if (e.Row["BANK_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BANK_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["OPEN_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OPEN_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["CURRENCY_CODE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CURRENCY_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["OPEN_AMOUNT"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OPEN_AMOUNT))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            } 
        }

        private void idaLC_MASTER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Currency_Amount();
        }

        #endregion

    }
}