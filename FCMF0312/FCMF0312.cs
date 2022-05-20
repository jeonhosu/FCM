using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.IO;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0312
{
    public partial class FCMF0312 : Office2007Form
    {       
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCurrency_Code;
        
        #endregion;

        #region ----- Constructor -----

        public FCMF0312()
        {
            InitializeComponent();
        }

        public FCMF0312(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void Search_DB()
        {            
            string vASSET_CODE = iConv.ISNull(igrASSET_MASTER.GetCellValue("ASSET_CODE"));
            int vCOL_IDX = igrASSET_MASTER.GetColumnToIndex("ASSET_CODE");

            //초기화.
            igrASSET_MASTER.LastConfirmChanges();
            IDA_ASSET_MASTER.OraSelectData.AcceptChanges();
            IDA_ASSET_MASTER.Refillable = true;

            igrASSET_HISTORY.LastConfirmChanges();
            IDA_ASSET_HISTORY.OraSelectData.AcceptChanges();
            IDA_ASSET_HISTORY.Refillable = true;

            IGR_ASSET_DPR_ACCOUNT.LastConfirmChanges();
            IDA_ASSET_DPR_ACCOUNT.OraSelectData.AcceptChanges();
            IDA_ASSET_DPR_ACCOUNT.Refillable = true;

            IGR_ASSET_DPR_RATE.LastConfirmChanges();
            IDA_ASSET_DPR_RATE.OraSelectData.AcceptChanges();
            IDA_ASSET_DPR_RATE.Refillable = true;
 
            IDA_ASSET_MASTER.Fill();
            if (iConv.ISNull(vASSET_CODE) != string.Empty)
            {
                for (int i = 0; i < igrASSET_MASTER.RowCount; i++)
                {
                    if (vASSET_CODE == iConv.ISNull(igrASSET_MASTER.GetCellValue(i, vCOL_IDX)))
                    {
                        igrASSET_MASTER.CurrentCellMoveTo(i, vCOL_IDX);
                        igrASSET_MASTER.CurrentCellActivate(i, vCOL_IDX);
                        return;
                    }
                }
            }
            igrASSET_MASTER.Focus();
        }

        private void Insert_Asset_Master()
        {
            CURRENCY_CODE.EditValue = mCurrency_Code;
            Init_Currency_Amount();

            ACQUIRE_DATE.EditValue = iDate.ISGetDate(DateTime.Today);
            REGISTER_DATE.EditValue = iDate.ISGetDate(DateTime.Today);
            QTY.EditValue = 0;
            AMOUNT.EditValue = 0;

            IDC_ASSET_STATUS.SetCommandParamValue("W_GROUP_CODE", "ASSET_STATUS");
            IDC_ASSET_STATUS.ExecuteNonQuery();
            ASSET_STATUS_CODE.EditValue = IDC_ASSET_STATUS.GetCommandParamValue("O_CODE");
            ASSET_STATUS_DESC.EditValue = IDC_ASSET_STATUS.GetCommandParamValue("O_CODE_NAME");

            TB_ASSET_MASTER.SelectedIndex = 0;
            TB_ASSET_MASTER.SelectedTab.Focus();

            ASSET_DESC.Focus();
        }

        private void Insert_Asset_History()
        {
            // 자산마스터 내용 INSERT.
            igrASSET_HISTORY.SetCellValue("CHARGE_DATE", DateTime.Today);
            igrASSET_HISTORY.SetCellValue("CURRENCY_CODE", CURRENCY_CODE.EditValue);
            igrASSET_HISTORY.SetCellValue("EXCHANGE_RATE", EXCHANGE_RATE.EditValue);
            igrASSET_HISTORY.SetCellValue("CURR_AMOUNT", CURR_AMOUNT.EditValue);
            igrASSET_HISTORY.SetCellValue("AMOUNT", 0);
            igrASSET_HISTORY.SetCellValue("QTY", QTY.EditValue);
            igrASSET_HISTORY.SetCellValue("EXPENSE_TYPE", EXPENSE_TYPE.EditValue);
            igrASSET_HISTORY.SetCellValue("EXPENSE_DESC", EXPENSE_TYPE_DESC.EditValue);
            igrASSET_HISTORY.SetCellValue("LOCATION_ID", LOCATION_ID.EditValue);
            igrASSET_HISTORY.SetCellValue("LOCATION_NAME", LOCATION_NAME.EditValue);
            igrASSET_HISTORY.SetCellValue("MANAGE_DEPT_ID", MANAGE_DEPT_ID.EditValue);
            igrASSET_HISTORY.SetCellValue("MANAGE_DEPT_NAME", MANAGE_DEPT_NAME.EditValue);
            igrASSET_HISTORY.SetCellValue("USE_DEPT_ID", USE_DEPT_ID.EditValue);
            igrASSET_HISTORY.SetCellValue("USE_DEPT_NAME", USE_DEPT_NAME.EditValue);
            igrASSET_HISTORY.SetCellValue("FIRST_USER", FIRST_USER.EditValue);
            igrASSET_HISTORY.SetCellValue("SECOND_USER", SECOND_USER.EditValue);
            igrASSET_HISTORY.SetCellValue("COST_CENTER_ID", COST_CENTER_ID.EditValue);
            igrASSET_HISTORY.SetCellValue("CC_CODE", CC_CODE.EditValue);
            igrASSET_HISTORY.SetCellValue("COST_CENTER_NAME", CC_DESC.EditValue);
            
            Init_H_Currency_Amount();

            int vCOL_IDX = igrASSET_HISTORY.GetColumnToIndex("CHARGE_DATE");
            igrASSET_HISTORY.CurrentCellMoveTo(vCOL_IDX);
            igrASSET_HISTORY.CurrentCellActivate(vCOL_IDX);
        }

        private void Insert_DPR_ACCOUNT()
        {
            IGR_ASSET_DPR_ACCOUNT.SetCellValue("COST_CENTER_ID", COST_CENTER_ID.EditValue);
            IGR_ASSET_DPR_ACCOUNT.SetCellValue("COST_CENTER_CODE", CC_CODE.EditValue);
            IGR_ASSET_DPR_ACCOUNT.SetCellValue("COST_CENTER_DESC", CC_DESC.EditValue);
            IGR_ASSET_DPR_ACCOUNT.SetCellValue("ENABLED_FLAG", "Y");
            IGR_ASSET_DPR_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", ACQUIRE_DATE.EditValue);
            IGR_ASSET_DPR_ACCOUNT.Focus();
        }

        private void Insert_DPR_Rate()
        {
            IGR_ASSET_DPR_RATE.SetCellValue("EFFECTIVE_DATE", ACQUIRE_DATE.EditValue);
            IGR_ASSET_DPR_RATE.Focus();
        }

        private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }
        
        private void Init_Currency_Amount()
        {
            if (iConv.ISNull(CURRENCY_CODE.EditValue) == string.Empty || CURRENCY_CODE.EditValue.ToString() == iConv.ISNull(mCurrency_Code, "KRW"))
            {
                if (iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue, 0) != 0)
                {
                    EXCHANGE_RATE.EditValue = DBNull.Value;
                }
                if (iConv.ISDecimaltoZero(CURR_AMOUNT.EditValue) != Convert.ToDecimal(0))
                {
                    CURR_AMOUNT.EditValue = 0;
                }
                EXCHANGE_RATE.ReadOnly = true;
                EXCHANGE_RATE.Insertable = false;
                EXCHANGE_RATE.Updatable = false;

                CURR_AMOUNT.ReadOnly = true;
                CURR_AMOUNT.Insertable = false;
                CURR_AMOUNT.Updatable = false;

                EXCHANGE_RATE.TabStop = false;
                CURR_AMOUNT.TabStop = false;
            }
            else
            {
                EXCHANGE_RATE.ReadOnly = false;
                EXCHANGE_RATE.Insertable = true;
                EXCHANGE_RATE.Updatable = true;

                CURR_AMOUNT.ReadOnly = false;
                CURR_AMOUNT.Insertable = true;
                CURR_AMOUNT.Updatable = true;

                EXCHANGE_RATE.TabStop = true;
                CURR_AMOUNT.TabStop = true;
            }
            EXCHANGE_RATE.Refresh();
            CURR_AMOUNT.Refresh();
        }

        private void Init_H_Currency_Amount()
        {
            if (iConv.ISNull(igrASSET_HISTORY.GetCellValue("CURRENCY_CODE")) == string.Empty 
                || iConv.ISNull(igrASSET_HISTORY.GetCellValue("CURRENCY_CODE")) == mCurrency_Code.ToString())
            {                
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("EXCHANGE_RATE")].ReadOnly = 1;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("EXCHANGE_RATE")].Insertable = 0;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("EXCHANGE_RATE")].Updatable = 0;

                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("CURR_AMOUNT")].ReadOnly = 1;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("CURR_AMOUNT")].Insertable = 0;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("CURR_AMOUNT")].Updatable = 0;
            }
            else
            {
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("EXCHANGE_RATE")].ReadOnly = 0;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("EXCHANGE_RATE")].Insertable = 1;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("EXCHANGE_RATE")].Updatable = 1;

                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("CURR_AMOUNT")].ReadOnly = 0;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("CURR_AMOUNT")].Insertable = 1;
                igrASSET_HISTORY.GridAdvExColElement[igrASSET_HISTORY.GetColumnToIndex("CURR_AMOUNT")].Updatable = 1;
            }
            igrASSET_HISTORY.Refresh();
        }

        private void Init_Asset_Amount()
        {
            decimal mAMOUNT = 0;
            if (iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue) != 0 &&
                iConv.ISDecimaltoZero(CURR_AMOUNT.EditValue) != 0)
            {
                if (iConv.ISDecimaltoZero(AMOUNT.EditValue) == 0)
                {
                    mAMOUNT = iConv.ISDecimaltoZero(EXCHANGE_RATE.EditValue) * iConv.ISDecimaltoZero(CURR_AMOUNT.EditValue);
                    mAMOUNT = Math.Round(mAMOUNT, 0);
                    AMOUNT.EditValue = mAMOUNT;
                }
            }
        }

        private void Set_Tab_Focus()
        {
            
        }

        private decimal Get_Last_Book_Amount(object pAsset_Amount, object pLast_Book_Rate)
        {
            IDC_GET_LAST_BOOK_AMOUNT.SetCommandParamValue("W_ASSET_AMOUNT", pAsset_Amount);
            IDC_GET_LAST_BOOK_AMOUNT.SetCommandParamValue("W_LAST_BOOK_RATE", pLast_Book_Rate);
            IDC_GET_LAST_BOOK_AMOUNT.ExecuteNonQuery();
            decimal vLast_Book_Amount = iConv.ISDecimaltoZero(IDC_GET_LAST_BOOK_AMOUNT.GetCommandParamValue("O_LAST_BOOK_AMOUNT"));
            return vLast_Book_Amount;
        }

        private void Get_DPR_Rate(object pDPR_TYPE, object pUSEFUL_TYPE, object pUSEFUL_LIFE)
        {
            IDC_DPR_RATE.SetCommandParamValue("W_DPR_TYPE", pDPR_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_TYPE", pUSEFUL_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_LIFE", pUSEFUL_LIFE);
            IDC_DPR_RATE.ExecuteNonQuery();
            IGR_ASSET_DPR_RATE.SetCellValue("DPR_RATE", IDC_DPR_RATE.GetCommandParamValue("O_DPR_RATE"));
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
        }

        #endregion;


        #region ----- Excel Upload : Asset Master -----

        private void Select_Excel_File_10()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FILE_PATH_MASTER.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    FILE_PATH_MASTER.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private void Excel_Upload_10()
        {
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            bool vXL_Load_OK = false;            

            if (iConv.ISNull(FILE_PATH_MASTER.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH_MASTER))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
                       
            string vOPenFileName = FILE_PATH_MASTER.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }


            //기존자료 삭제.
            vSTATUS = "F";
            vMESSAGE = string.Empty;

            IDC_DELETE_ASSET_MASTER_TEMP.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_DELETE_ASSET_MASTER_TEMP.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_DELETE_ASSET_MASTER_TEMP.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_TRANS_ASSET_MASTER.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
            
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            
            // 업로드 아답터 fill //
            IDA_ASSET_MASTER_UPLOAD.Cancel();
            IDA_ASSET_MASTER_UPLOAD.Fill();
            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL_10(IDA_ASSET_MASTER_UPLOAD, 2);
                    if (vXL_Load_OK == false)
                    {
                        IDA_ASSET_MASTER_UPLOAD.Cancel();
                    }
                    else
                    {
                        IDA_ASSET_MASTER_UPLOAD.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_ASSET_MASTER_UPLOAD.Cancel();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }
            vXL_Upload.DisposeXL();


            if (IDA_ASSET_MASTER_UPLOAD.IsUpdateCompleted == true)
            {
                vSTATUS = "F";
                vMESSAGE = string.Empty;

                IDC_SET_TRANS_ASSET_MASTER.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_SET_TRANS_ASSET_MASTER.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_TRANS_ASSET_MASTER.GetCommandParamValue("O_MESSAGE"));

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (IDC_SET_TRANS_ASSET_MASTER.ExcuteError || vSTATUS == "F")
                {
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- Excel Upload : Asset history -----

        private void Select_Excel_File_20()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FILE_PATH_HISTORY.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    FILE_PATH_HISTORY.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private void Excel_Upload_20()
        {
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            bool vXL_Load_OK = false;

            if (iConv.ISNull(FILE_PATH_HISTORY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH_MASTER))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vOPenFileName = FILE_PATH_HISTORY.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }


            //기존자료 삭제.
            vSTATUS = "F";
            vMESSAGE = string.Empty;

            IDC_DELETE_ASSET_HISTORY_TEMP.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_DELETE_ASSET_HISTORY_TEMP.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_DELETE_ASSET_HISTORY_TEMP.GetCommandParamValue("O_MESSAGE"));
            if (IDC_DELETE_ASSET_HISTORY_TEMP.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
                return;
            }

            // 업로드 아답터 fill //
            IDA_ASSET_HISTORY_UPLOAD.Cancel();
            IDA_ASSET_HISTORY_UPLOAD.Fill();            
            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL_20(IDA_ASSET_HISTORY_UPLOAD, 2);
                    if (vXL_Load_OK == false)
                    {
                        IDA_ASSET_HISTORY_UPLOAD.Cancel();
                    }
                    else
                    {
                        IDA_ASSET_HISTORY_UPLOAD.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_ASSET_HISTORY_UPLOAD.Cancel();
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }
            vXL_Upload.DisposeXL();


            if (IDA_ASSET_HISTORY_UPLOAD.IsUpdateCompleted == true)
            {
                vSTATUS = "F";
                vMESSAGE = string.Empty;

                IDC_SET_TRANS_ASSET_HISTORY.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_SET_TRANS_ASSET_HISTORY.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_TRANS_ASSET_HISTORY.GetCommandParamValue("O_MESSAGE"));

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (IDC_DELETE_ASSET_HISTORY_TEMP.ExcuteError || vSTATUS == "F")
                {
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } 
                    return;
                }
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion
        
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
                    if (IDA_ASSET_MASTER.IsFocused)
                    {
                        IDA_ASSET_MASTER.AddOver();
                        Insert_Asset_Master();
                    }
                    else if (IDA_ASSET_HISTORY.IsFocused)
                    {
                        IDA_ASSET_HISTORY.AddOver();
                        Insert_Asset_History();
                    }
                    else if (IDA_ASSET_DPR_ACCOUNT.IsFocused)
                    {
                        IDA_ASSET_DPR_ACCOUNT.AddOver();
                        Insert_DPR_ACCOUNT();
                    }
                    else if (IDA_ASSET_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_DPR_RATE.AddOver();
                        Insert_DPR_Rate();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ASSET_MASTER.IsFocused)
                    {
                        IDA_ASSET_MASTER.AddUnder();
                        Insert_Asset_Master();
                    }
                    else if (IDA_ASSET_HISTORY.IsFocused)
                    {
                        IDA_ASSET_HISTORY.AddUnder();
                        Insert_Asset_History();
                    }
                    else if (IDA_ASSET_DPR_ACCOUNT.IsFocused)
                    {
                        IDA_ASSET_DPR_ACCOUNT.AddUnder();
                        Insert_DPR_ACCOUNT();
                    }
                    else if (IDA_ASSET_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_DPR_RATE.AddUnder();
                        Insert_DPR_Rate();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ASSET_MASTER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_MASTER.IsFocused)
                    {
                        IDA_ASSET_DPR_RATE.Cancel();
                        IDA_ASSET_DPR_ACCOUNT.Cancel(); 
                        IDA_ASSET_HISTORY.Cancel();
                        IDA_ASSET_MASTER.Cancel();
                    }
                    else if (IDA_ASSET_HISTORY.IsFocused)
                    {
                        IDA_ASSET_HISTORY.Cancel();
                    }
                    else if (IDA_ASSET_DPR_ACCOUNT.IsFocused)
                    {
                        IDA_ASSET_DPR_ACCOUNT.Cancel(); 
                    }
                    else if (IDA_ASSET_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_DPR_RATE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ASSET_MASTER.IsFocused)
                    {
                        IDA_ASSET_MASTER.Delete();
                    }
                    else if (IDA_ASSET_HISTORY.IsFocused)
                    {
                        IDA_ASSET_HISTORY.Delete();
                    }
                    else if (IDA_ASSET_DPR_ACCOUNT.IsFocused)
                    {
                        IDA_ASSET_DPR_ACCOUNT.Delete();
                    }
                    else if (IDA_ASSET_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_DPR_RATE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        
        private void FCMF0312_Load(object sender, EventArgs e)
        {
            W_CIP_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_CIP_FLAG.EditValue = W_CIP_ALL.RadioButtonString; 
            
            IDA_ASSET_MASTER.FillSchema();
            IDA_ASSET_HISTORY.FillSchema();
            IDA_ASSET_DPR_ACCOUNT.FillSchema();
            IDA_ASSET_DPR_RATE.FillSchema();
        }

        private void FCMF0312_Shown(object sender, EventArgs e)
        {
            IDC_BASE_CURRENCY.ExecuteNonQuery();
            mCurrency_Code = IDC_BASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE");

            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "ASSET_STATUS");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            W_ASSET_STATUS_CODE.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE");
            W_ASSET_STATUS_NAME.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");

            BTN_SELECT_2.BringToFront();
            W_ENABLED_FLAG_2.BringToFront();
        }

        private void IGR_ASSET_DPR_RATE_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            int vIDX_USEFUL_TYPE = IGR_ASSET_DPR_RATE.GetColumnToIndex("USEFUL_TYPE");
            int vIDX_USEFUL_LIFE = IGR_ASSET_DPR_RATE.GetColumnToIndex("USEFUL_LIFE");
            if (e.ColIndex == vIDX_USEFUL_TYPE)
            {
                Get_DPR_Rate(IGR_ASSET_DPR_RATE.GetCellValue("DPR_TYPE"), e.CellValue, IGR_ASSET_DPR_RATE.GetCellValue("USEFUL_LIFE"));
            }
            else if (e.ColIndex == vIDX_USEFUL_LIFE)
            {
                Get_DPR_Rate(IGR_ASSET_DPR_RATE.GetCellValue("DPR_TYPE"), IGR_ASSET_DPR_RATE.GetCellValue("USEFUL_TYPE"), e.CellValue);
            }
        }

        private void IGR_ASSET_DPR_RATE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_LAST_BOOK_RATE = IGR_ASSET_DPR_RATE.GetColumnToIndex("LAST_BOOK_RATE");
            if (vIDX_LAST_BOOK_RATE == e.ColIndex)
            {
                decimal vASSET_AMOUNT = iConv.ISDecimaltoZero(AMOUNT.EditValue, 0); 
                decimal vLAST_BOOK_AMOUNT = Get_Last_Book_Amount(vASSET_AMOUNT, e.NewValue);

                IGR_ASSET_DPR_RATE.SetCellValue("LAST_BOOK_AMOUNT", vLAST_BOOK_AMOUNT);
            }
        }

        private void itbASSET_MASTER_Click(object sender, EventArgs e)
        {
            Set_Tab_Focus();
        }

        private void ACQUIRE_DATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iConv.ISNull(ACQUIRE_DATE.EditValue) != String.Empty && CIP_FLAG.CheckedState == ISUtil.Enum.CheckedState.Unchecked)
            {
                REGISTER_DATE.EditValue = ACQUIRE_DATE.EditValue;
            }
        }

        private void EXCHANGE_RATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Asset_Amount();
        }

        private void CURR_AMOUNT_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            Init_Asset_Amount();
        }

        private void tabPageAdv1_Click(object sender, EventArgs e)
        {
            FILE_PATH_MASTER.EditValue = string.Empty;
            FILE_PATH_HISTORY.EditValue = string.Empty;
        }

        private void BTN_SELECT_MASTER_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_MASTER_TEMP.Fill();
        }

        private void BTN_FILE_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File_10();
        }

        private void BTN_UPLOAD_EXEC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Excel_Upload_10();  // 자산대장 업로드 //   
        }

        private void BTN_SELECT_HISTORY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_HISTORY_TEMP.Fill();
        }

        private void BTN_FILE_SELECT_HISTORY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File_20();
        }

        private void BTN_UPLOAD_EXEC_HISTORY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Excel_Upload_20();  // 자산대장 업로드 //   
        }

        private void W_CIP_ALL_Click(object sender, EventArgs e)
        {
            if (W_CIP_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_ALL.RadioButtonString;
            }
        }

        private void W_CIP_NO_Click(object sender, EventArgs e)
        {
            if (W_CIP_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_NO.RadioButtonString;
            }
        }

        private void W_CIP_YES_Click(object sender, EventArgs e)
        {
            if (W_CIP_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CIP_FLAG.EditValue = W_CIP_YES.RadioButtonString;
            }
        }

        private void CIP_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IDA_ASSET_MASTER.CurrentRow.RowState == DataRowState.Added)
            {
                if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
                {
                    PARENT_ASSET_CODE.Insertable = false;
                }
                else
                {
                    PARENT_ASSET_CODE.Insertable = true;
                }
            }
            else
            {
                PARENT_ASSET_CODE.Insertable = false;
            }
        }

        private void BTN_SELECT_2_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_DPR_ACCOUNT.Fill();
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaASSET_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_ASSET_CATEGORY_CLASS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", W_ASSET_CATEGORY_ID.EditValue);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 2);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaASSET_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY_INPUT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCOSTCENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildCOSTCENTER.SetLookupParamValue("W_STD_DATE", ACQUIRE_DATE.EditValue);
        }

        private void ilaEXPENSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("EXPENSE_TYPE", "Y");
        }

        private void ilaIFRS_DPR_METHOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("DPR_METHOD_TYPE", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaUSE_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildDEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaDPR_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("DPR_TYPE", "Y");
        }

        private void ILA_DPR_TYPE_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("DPR_TYPE", "Y");
        }

        private void ILA_USEFUL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("USEFUL_TYPE", "Y");
        }

        private void ilaASSET_LOCATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_LOCATION", "Y");
        }

        private void ilaCURRENCY_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaASSET_STATUS_NAME_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_STATUS", "Y");
        }

        private void ilaCURRENCY_CODE_SelectedRowData(object pSender)
        {
            Init_Currency_Amount();
        }

        private void ilaASSET_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_STATUS", "Y");
        }

        private void ilaCUSTOMER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUSTOMER.SetLookupParamValue("W_SUPP_CUST_TYPE", "S");
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_LEASE_COMPANY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUSTOMER.SetLookupParamValue("W_SUPP_CUST_TYPE", "S");
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaFIRST_USER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_START_DATE", DateTime.Today);
            ildPERSON.SetLookupParamValue("W_END_DATE", iDate.ISMonth_1st(DateTime.Today));;
        }

        private void ILA_COSTCENTER_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOSTCENTER.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ilaHAVE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_HAVE_TYPE", "Y");
        }

        private void ilaTAX_DED_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("TAX_DED_TYPE", "Y");
        }

        private void ilaDPR_TYPE_SelectedRowData(object pSender)
        {
            Get_DPR_Rate(IGR_ASSET_DPR_RATE.GetCellValue("DPR_TYPE"), IGR_ASSET_DPR_RATE.GetCellValue("USEFUL_TYPE"), IGR_ASSET_DPR_RATE.GetCellValue("USEFUL_LIFE"));
        }

        private void ilaH_ASSET_CHARGE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_CHARGE", "Y");
        }

        private void ilaH_CURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ila_H_EXPENSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("EXPENSE_TYPE", "Y");
        }

        private void ilaH_ASSET_LOCATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_LOCATION", "Y");
        }

        private void ilaH_MANAGE_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildDEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaH_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildDEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaH_CURRENCY_SelectedRowData(object pSender)
        {
            Init_H_Currency_Amount();
        }

        private void ilaMANAGE_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildDEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaUSE_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildDEPT_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_ASSET_AST_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (CIP_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                ILD_ASSET_AST_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_ID", 0);
            }
            else
            {
                ILD_ASSET_AST_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_ID", AST_CATEGORY_ID.EditValue);
            }
        }

        private void ILA_DPR_ACCOUNT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DPR_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaOPERATION_DIVISION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("OPERATION_DIVISION", "Y");
        }

        private void ilaH_OPERATION_DIVISION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("OPERATION_DIVISION", "Y");
        }
        #endregion

        #region ----- Adapeter Event -----

        private void idaASSET_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ASSET_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10201"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                ASSET_DESC.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["EXPENSE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10216"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                EXPENSE_TYPE_DESC.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["AST_CATEGORY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10093"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                AST_CATEGORY_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["ACQUIRE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10203"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                ACQUIRE_DATE.Focus();
                return;
            }
            //if (iString.ISNull(e.Row["REGISTER_DATE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10204"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    REGISTER_DATE.Focus();
            //    return;
            //}
            //if (iString.ISNull(e.Row["LOCATION_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10205"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    LOCATION_NAME.Focus();
            //    return;
            //}            
            //if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    CURRENCY_CODE.Focus();
            //    return;
            //}       
            if (iConv.ISNull(e.Row["QTY"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10204"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                REGISTER_DATE.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10208"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                AMOUNT.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["CIP_FLAG"]) == "Y" && iConv.ISNull(e.Row["PARENT_ASSET_ID"]) != string.Empty)
            {
                //건설중인 자산은 부모자산을 선택 할 수 없다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10568"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                PARENT_ASSET_CODE.Focus();
                return;
            }
        }

        private void idaASSET_MASTER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                if (iConv.ISNull(e.Row["ASSET_CODE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10209"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaASSET_MASTER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            if (pBindingManager.DataRow.RowState == DataRowState.Added)
            {
                CIP_FLAG.ReadOnly = false;
                PARENT_ASSET_CODE.Insertable = true;
            }
            else
            {
                CIP_FLAG.ReadOnly = true;
                PARENT_ASSET_CODE.Insertable = false;
            }
        }

        private void idaASSET_HISTORY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["CHARGE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10223"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrASSET_HISTORY.CurrentCellMoveTo(igrASSET_HISTORY.RowIndex, igrASSET_HISTORY.GetColumnToIndex("CHARGE_DATE"));
                igrASSET_HISTORY.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["CHARGE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10224"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrASSET_HISTORY.CurrentCellMoveTo(igrASSET_HISTORY.RowIndex, igrASSET_HISTORY.GetColumnToIndex("CHARGE_NAME"));
                igrASSET_HISTORY.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10520"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrASSET_HISTORY.CurrentCellMoveTo(igrASSET_HISTORY.RowIndex, igrASSET_HISTORY.GetColumnToIndex("AMOUNT"));
                igrASSET_HISTORY.Focus();
                return;
            }
        }

        private void idaASSET_HISTORY_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        private void IDA_ASSET_DPR_ACCOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ASSET_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10399"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; 
                return;
            }
            if (iConv.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["COST_CENTER_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10018"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DIST_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10574"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_ASSET_DPR_ACCOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return; 
            }
        }

        private void idaASSET_DPR_RATE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["DPR_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10097"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["USEFUL_TYPE"]) == string.Empty)
            {
                int vIDX_COL = IGR_ASSET_DPR_RATE.GetColumnToIndex("USEFUL_TYPE_DESC");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ASSET_DPR_RATE, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["USEFUL_LIFE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE"]) == string.Empty)
            {
                int vIDX_COL = IGR_ASSET_DPR_RATE.GetColumnToIndex("EFFECTIVE_DATE");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ASSET_DPR_RATE, vIDX_COL))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaASSET_HISTORY_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (igrASSET_HISTORY.RowIndex > -1)
            {
                Init_H_Currency_Amount();
            }
        }


        #endregion
         
    }
}