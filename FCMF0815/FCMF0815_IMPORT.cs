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
using System.IO;

namespace FCMF0815
{
    public partial class FCMF0815_IMPORT : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Constructor -----

        public FCMF0815_IMPORT(Form pMainForm, ISAppInterface pAppInterface, object pSESSION_ID
                                , object pVAT_SELECT_PERIOD, object pVAT_SELECT_PERIOD_NAME
                                , object pVAT_ISSUE_DATE_FR, object pVAT_ISSUE_DATE_TO)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_VAT_SELECT_PERIOD.EditValue = pVAT_SELECT_PERIOD;
            V_VAT_SELECT_PERIOD_DESC.EditValue = pVAT_SELECT_PERIOD_NAME;
            V_ISSUE_DATE_FR.EditValue = pVAT_ISSUE_DATE_FR;
            V_ISSUE_DATE_TO.EditValue = pVAT_ISSUE_DATE_TO;
        }

        #endregion;

        #region ----- Property / Method ----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        #endregion

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

                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx";
                openFileDialog1.DefaultExt = "xlsx";
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

            if (iConv.ISNull(UPLOAD_FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPLOAD_FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(V_START_ROW.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_START_ROW))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(V_VAT_CATEGORY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_VAT_CATEGORY))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            //if (iConv.ISNull(V_STD_DATE.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_STD_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return vResult;
            //}

            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            string vVAT_GUBUN = string.Empty;
            string vVAT_TYPE = string.Empty;
            if (iConv.ISNull(V_VAT_CATEGORY.EditValue) == "P_TAX_BILL")
            {
                vVAT_GUBUN = "1";
                vVAT_TYPE = "TAX_BILL";
            }
            else if (iConv.ISNull(V_VAT_CATEGORY.EditValue) == "P_BILL")
            {
                vVAT_GUBUN = "1";
                vVAT_TYPE = "BILL";
            }
            else if (iConv.ISNull(V_VAT_CATEGORY.EditValue) == "S_TAX_BILL")
            {
                vVAT_GUBUN = "2";
                vVAT_TYPE = "TAX_BILL";
            }
            else if (iConv.ISNull(V_VAT_CATEGORY.EditValue) == "S_BILL")
            {
                vVAT_GUBUN = "2";
                vVAT_TYPE = "BILL";
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_VAT_CATEGORY))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            if (V_DEL_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {                
                if (iConv.ISNull(V_ISSUE_DATE_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return vResult;
                }
                if (iConv.ISNull(V_ISSUE_DATE_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return vResult;
                }
                if (iDate.ISGetDate(V_ISSUE_DATE_FR.EditValue) > iDate.ISGetDate(V_ISSUE_DATE_TO.EditValue))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return vResult;
                }

                V_MESSAGE.PromptText = "Pre-Data Deleting Start....";
                IDC_DELETE_EXCEL.SetCommandParamValue("P_VAT_GUBUN", vVAT_GUBUN);
                IDC_DELETE_EXCEL.SetCommandParamValue("P_VAT_TYPE", vVAT_TYPE);
                IDC_DELETE_EXCEL.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_DELETE_EXCEL.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_DELETE_EXCEL.GetCommandParamValue("O_MESSAGE"));
                if(vSTATUS == "F")
                {                    
                    if (vMESSAGE != String.Empty)
                    {
                        V_MESSAGE.PromptText = vMESSAGE;
                    }
                    Application.DoEvents();
                    return vResult;
                }  

                V_MESSAGE.PromptText = "Pre-Data Deleting End....";
                Application.DoEvents(); 
            }

            bool vXL_Load_OK = false;
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
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return vResult;
            }

            vSTATUS = "F";
            vMESSAGE = string.Empty; 

            V_MESSAGE.PromptText = "Importing Start....";
            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL(IDC_IMPORT_EXCEL, iConv.ISNumtoZero(V_START_ROW.EditValue, 2), V_PB_INTERFACE, V_MESSAGE,
                                                    vVAT_GUBUN, vVAT_TYPE);
                    if (vXL_Load_OK == false)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        vResult = false;
                        return vResult;
                    }
                    vResult = true;
                    V_MESSAGE.PromptText = "Importing Completed....";  
                }
            }
            catch (Exception ex)
            { 
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                vResult = false;
                return vResult;
            }
            vXL_Upload.DisposeXL();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return vResult;
        }

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----

        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
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
        #endregion

        #region ----- Form Event -----
         
        private void FCMF0815_IMPORT_Shown(object sender, EventArgs e)
        {
            IDC_GET_LOCAL_DATETIME_P.ExecuteNonQuery();
            V_SYSDATE.EditValue = IDC_GET_LOCAL_DATETIME_P.GetCommandParamValue("X_LOCAL_DATE");
            V_SYSDATE.BringToFront();

            V_MESSAGE.PromptText = "";
            V_PB_INTERFACE.BarFillPercent = 0; 
            Application.DoEvents();
        }

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            if (Excel_Upload() == true)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            } 
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void V_P_TAX_BILL_Click(object sender, EventArgs e)
        {
            if (V_P_TAX_BILL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_VAT_CATEGORY.EditValue = V_P_TAX_BILL.RadioCheckedString;
            }
        }

        private void V_P_BILL_Click(object sender, EventArgs e)
        {
            if (V_P_BILL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_VAT_CATEGORY.EditValue = V_P_BILL.RadioCheckedString;
            }
        }

        private void V_S_TAX_BILL_Click(object sender, EventArgs e)
        {
            if (V_S_TAX_BILL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_VAT_CATEGORY.EditValue = V_S_TAX_BILL.RadioCheckedString;
            }
        }

        private void V_S_BILL_Click(object sender, EventArgs e)
        {
            if (V_S_BILL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                V_VAT_CATEGORY.EditValue = V_S_BILL.RadioCheckedString;
            }
        }

        #endregion

        #region ----- Data Adapter Event -----

        #endregion

    }
}