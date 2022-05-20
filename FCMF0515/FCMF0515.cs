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

namespace FCMF0515
{
    public partial class FCMF0515 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0515()
        {
            InitializeComponent();
        }

        public FCMF0515(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (itbBILL.SelectedTab.TabIndex == 1)
            {
                IDA_BILL_MASTER.Fill();
                IGR_BILL_LIST.Focus();
            }
            else if (itbBILL.SelectedTab.TabIndex == 2)
            {
                IDA_BILL_MASTER_UPLOAD.Fill();
                IGR_BILL_MASTER_UPLOAD.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Insert_Data()
        {
            // 어음상태.
            IDC_DV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_STATUS");
            IDC_DV_COMMON.ExecuteNonQuery();
            BILL_STATUS_NAME.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE_NAME");
            BILL_STATUS.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE");

            //자타구분.
            IDC_DV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_MODE");
            IDC_DV_COMMON.ExecuteNonQuery();
            BILL_MODE_NAME.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE_NAME");
            BILL_MODE.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE");


            //어음구분
            //idcCOMMON_CODE_NAME.SetCommandParamValue("W_GROUP_CODE", "BILL_TYPE");
            //idcCOMMON_CODE_NAME.SetCommandParamValue("W_CODE", BILL_TYPE_0.EditValue);
            //idcCOMMON_CODE_NAME.ExecuteNonQuery();
            //BILL_TYPE_NAME.EditValue = idcCOMMON_CODE_NAME.GetCommandParamValue("O_RETURN_VALUE");
            //if (iString.ISNull(BILL_TYPE.EditValue) == string.Empty)
            //{
            //    //DEFAULT VALUE 설정 : 어음구분.
            //    IDC_DV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_TYPE");
            //    IDC_DV_COMMON.ExecuteNonQuery();
            //    BILL_TYPE_NAME.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE_NAME");
            //    BILL_TYPE.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE");
            //}

            if (iString.ISNull(BILL_TYPE.EditValue) == string.Empty)
            {
                BILL_TYPE_NAME.EditValue = BILL_TYPE_NAME_0.EditValue;
                BILL_TYPE.EditValue = BILL_TYPE_0.EditValue;
            }

            VAT_ISSUE_DATE.EditValue = DateTime.Today;
            ISSUE_DATE.EditValue = DateTime.Today;
            DUE_DATE.EditValue = DateTime.Today;
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

        #region ----- Excel Upload -----

        private void Select_Excel_File()
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
                    UPLOAD_FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private bool Excel_Upload()
        {
            bool vResult = false;

            string vStatus = "F";
            string vMessage = string.Empty;
            bool vXL_Load_OK = false;

            if (iString.ISNull(UPLOAD_FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPLOAD_FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            //기존자료 삭제.
            IDC_DELETE_BILL_MASTER_UPLOAD.ExecuteNonQuery();
            vStatus = IDC_DELETE_BILL_MASTER_UPLOAD.GetCommandParamValue("O_STATUS").ToString();
            vMessage = iString.ISNull(IDC_DELETE_BILL_MASTER_UPLOAD.GetCommandParamValue("O_MESSAGE"));
            if (IDC_DELETE_BILL_MASTER_UPLOAD.ExcuteError == true || vStatus == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return vResult;
            }

            //기존작업 취소.
            IDA_BILL_MASTER_UPLOAD_TEMP.Cancel();
            string vOPenFileName = UPLOAD_FILE_PATH.EditValue.ToString();
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
                return vResult;
            }

            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL(IDA_BILL_MASTER_UPLOAD_TEMP, 2);
                    if (vXL_Load_OK == false)
                    {
                        IDA_BILL_MASTER_UPLOAD_TEMP.Cancel();
                    }
                    else
                    {
                        IDA_BILL_MASTER_UPLOAD_TEMP.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return vResult;
            }
            vXL_Upload.DisposeXL();
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            vResult = true;
            return vResult;
        }

        private bool Set_Bill_Master_Transfer()
        {
            bool vResult = false;

            DialogResult dlgResult;
            dlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlgResult == DialogResult.No)
            {
                return vResult;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            object vMESSAGE = string.Empty;

            IDC_BILL_MASTER_TRANSFER.ExecuteNonQuery();
            vSTATUS = IDC_BILL_MASTER_TRANSFER.GetCommandParamValue("O_STATUS").ToString();
            vMESSAGE = IDC_BILL_MASTER_TRANSFER.GetCommandParamValue("O_MESSAGE");
            if (IDC_BILL_MASTER_TRANSFER.ExcuteError == true || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (iString.ISNull(vMESSAGE) != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return vResult;
            }
            vResult = true;
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            return vResult;
        }

        private bool Cancel_Bill_Master_Transfer()
        {
            bool vResult = false;

            DialogResult dlgResult;
            dlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlgResult == DialogResult.No)
            {
                return vResult;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            object vMESSAGE = String.Empty;

            IDC_CANCEL_BILL_MASTER_TRANSFER.ExecuteNonQuery();
            vSTATUS = IDC_CANCEL_BILL_MASTER_TRANSFER.GetCommandParamValue("O_STATUS").ToString();
            vMESSAGE = IDC_CANCEL_BILL_MASTER_TRANSFER.GetCommandParamValue("O_MESSAGE");
            if (IDC_CANCEL_BILL_MASTER_TRANSFER.ExcuteError == true || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                if (iString.ISNull(vMESSAGE) != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return vResult;
            }
            vResult = true;
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            return vResult;
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
                    if (IDA_BILL_MASTER.IsFocused)
                    {
                        IDA_BILL_MASTER.AddOver();
                        Insert_Data();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BILL_MASTER.IsFocused)
                    {
                        IDA_BILL_MASTER.AddUnder();
                        Insert_Data();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_BILL_MASTER.IsFocused)
                    {
                        IDA_BILL_MASTER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BILL_MASTER.IsFocused)
                    {
                        IDA_BILL_MASTER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    IDA_BILL_MASTER.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (IDA_BILL_MASTER.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (IDA_BILL_MASTER.IsFocused)
                    {
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0515_Load(object sender, EventArgs e)
        {
            IDA_BILL_MASTER.FillSchema();
        }

        private void FCMF0515_Shown(object sender, EventArgs e)
        {

        }

        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_BILL_MASTER_UPLOAD.SetSelectParamValue("P_SOB_ID", -1);
            IDA_BILL_MASTER_UPLOAD.Fill();

            IDA_BILL_MASTER_UPLOAD_TEMP.SetDeleteParamValue("P_SOB_ID", -1);
            IDA_BILL_MASTER_UPLOAD_TEMP.Fill();
            if (Excel_Upload() == true)
            {
                Search_DB();
            }
        }

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_SLIP_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Bill_Master_Transfer();
            Search_DB();
        }

        private void BTN_SLIP_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Cancel_Bill_Master_Transfer();
            Search_DB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaBILL_CLASS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_CLASS", "N");
        }

        private void ilaBILL_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_TYPE", "N");
        }

        private void ilaBILL_NUM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ilaVENDOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaBILL_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_STATUS", "N");
        }
                
        private void ilaBILL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_TYPE", "Y");
        }

        private void ilaBILL_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_STATUS", "Y");
        }

        private void ilaVENDOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBILL_MODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_MODE", "Y");
        }

        private void ilaSUPP_CUST_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBANK_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaRECEIPT_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaKEEP_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaRECEIPT_DEPT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaPERSON_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
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
            if (iString.ISNumtoZero(e.Row["BILL_AMOUNT"]) == 0)
            {// 어음금액
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10146"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BILL_AMOUNT.Focus();
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BILL_MODE"]) == string.Empty)
            {// 자타구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10353"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BILL_MODE_NAME.Focus();
                e.Cancel = true;
                return;
            }
        }

        #endregion


    }
}