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

namespace FCMF0821
{
    public partial class FCMF0821 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0821()
        {
            InitializeComponent();
        }

        public FCMF0821(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void Set_Default_Value()
        {
            //세금계산서 발행기간.
            DateTime vGetDateTime = GetDate();
            W_PERIOD_YEAR.EditValue = iDate.ISYear(vGetDateTime);

            //사업장 구분.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
            idcDV_COMMON.ExecuteNonQuery();
            W_TAX_CODE_NAME.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            W_TAX_CODE.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");

            //영세율적용근거.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "VAT_APPLY_BASE");
            idcDV_COMMON.ExecuteNonQuery();
            APPLY_BASE_DESC_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            APPLY_BASE_CODE_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");

            //영세율제출서류명.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "VAT_SUBMIT_DOCUMENT");
            idcDV_COMMON.ExecuteNonQuery();
            SUBMIT_DOCUMENT_DESC_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            SUBMIT_DOCUMENT_CODE_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");

            //법정서식제출불능사유.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "VAT_SUBMIT_IMPOSSIBLE");
            idcDV_COMMON.ExecuteNonQuery();
            SUBMIT_IMPOSSIBLE_DESC_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            SUBMIT_IMPOSSIBLE_CODE_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");
            
            W_WRITE_DATE.EditValue = vGetDateTime;

            //부가세 과세구분//
            IDC_GET_VAT_LEVIER_TYPE_P.ExecuteNonQuery();
            string vVAT_LEVIER_TYPE = iString.ISNull(IDC_GET_VAT_LEVIER_TYPE_P.GetCommandParamValue("O_VAT_LEVIER_TYPE"));
            if (vVAT_LEVIER_TYPE == "5")
            {
                V_BUSINESS_UNIT_TAX_YN.Visible = true;
                V_BUSINESS_UNIT_TAX_YN.BringToFront();
            }
            else
            {
                V_BUSINESS_UNIT_TAX_YN.Visible = false;
            }
        }

        private void SEARCH_DB()
        {
            if (iString.ISNull(W_TAX_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TAX_CODE_NAME.Focus();
                return;
            }

            if (iString.ISNull(W_VAT_PERIOD_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10487"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (Convert.ToDateTime(W_ISSUE_DATE_FR.EditValue) > Convert.ToDateTime(W_ISSUE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            idaFOREIGN_CURR_HEADER.Fill();
            igrFOREIGN_CURR_LINE.Focus();
        }

        private bool VAT_PERIOD_CHECK()
        {
            //신고기간 검증.
            string vCHECK_YN = "N";
            idcVAT_PERIOD_CHECK.ExecuteNonQuery();
            vCHECK_YN = iString.ISNull(idcVAT_PERIOD_CHECK.GetCommandParamValue("O_YN"));
            if (vCHECK_YN == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10396"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return false;
            }
            return true;
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Show_Slip_Detail()
        {
            int mSLIP_HEADER_ID = iString.ISNumtoZero(igrFOREIGN_CURR_LINE.GetCellValue("INTERFACE_HEADER_ID"));
            if (mSLIP_HEADER_ID != Convert.ToInt32(0))
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
                vFCMF0204.Show();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.UseWaitCursor = false;
            }
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

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = pGrid.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                idcVAT_PERIOD.ExecuteNonQuery();
                object vPeriod = idcVAT_PERIOD.GetCommandParamValue("O_PERIOD");

                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0821_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    // 헤더 인쇄.
                    idaOPERATING_UNIT.Fill();
                    if (idaOPERATING_UNIT.SelectRows.Count > 0)
                    {
                        xlPrinting.HeaderWrite(idaOPERATING_UNIT, vPeriod
                                            , APPLY_BASE_DESC.EditValue, SUBMIT_IMPOSSIBLE_DESC.EditValue
                                            , SUBMIT_DOCUMENT_DESC.EditValue, SUBMIT_DATE.EditValue
                                            , ATTACH_CONTENTS_1.EditValue, ATTACH_CONTENTS_2.EditValue, ATTACH_CONTENTS_3.EditValue);
                    }
                    // 실제 인쇄
                    vPageNumber = xlPrinting.LineWrite(pGrid);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE("FOREIGN_CURRENCY_");
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                //신고기간 검증.
                if (VAT_PERIOD_CHECK() == false)
                {
                    return;
                }

                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaFOREIGN_CURR_LINE.IsFocused)
                    {
                        idaFOREIGN_CURR_LINE.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaFOREIGN_CURR_LINE.IsFocused)
                    {
                        idaFOREIGN_CURR_LINE.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaFOREIGN_CURR_HEADER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaFOREIGN_CURR_LINE.IsFocused)
                    {
                        idaFOREIGN_CURR_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaFOREIGN_CURR_LINE.IsFocused)
                    {
                        idaFOREIGN_CURR_LINE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT", igrFOREIGN_CURR_LINE);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE", igrFOREIGN_CURR_LINE);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0821_Load(object sender, EventArgs e)
        {
            idaFOREIGN_CURR_HEADER.FillSchema();
            idaFOREIGN_CURR_LINE.FillSchema();
        }

        private void FCMF0821_Shown(object sender, EventArgs e)
        {
            Set_Default_Value();
        }

        private void igrVAT_EXPORT_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        private void ibtnSET_FOREIGN_CURR_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_TAX_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TAX_CODE_NAME.Focus();
                return;
            }

            if (iString.ISNull(W_VAT_PERIOD_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10487"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (Convert.ToDateTime(W_ISSUE_DATE_FR.EditValue) > Convert.ToDateTime(W_ISSUE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            //신고기간 검증.
            if (VAT_PERIOD_CHECK() == false)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
            isDataTransaction1.BeginTran();
            idcSET_FOREIGN_CURR.ExecuteNonQuery();
            mSTATUS = iString.ISNull(idcSET_FOREIGN_CURR.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(idcSET_FOREIGN_CURR.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
            if (idcSET_FOREIGN_CURR.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            isDataTransaction1.Commit();
            if (mMESSAGE != String.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SUBMIT_DATE_0_EditValueChanged(object pSender)
        {
            if (idaFOREIGN_CURR_HEADER.OraSelectData.Rows.Count > 0)
            {
                SUBMIT_DATE.EditValue = SUBMIT_DATE_0.EditValue;
            }
        }

        #endregion

        #region ----- Lookup Event : Search -----

        private void ilaTAX_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        }

        private void ilaTAX_CODE_0_SelectedRowData(object pSender)
        {
            W_VAT_PERIOD_DESC.EditValue = string.Empty;
            W_VAT_PERIOD_ID.EditValue = DBNull.Value;
            W_ISSUE_DATE_FR.EditValue = DBNull.Value;
            W_ISSUE_DATE_TO.EditValue = DBNull.Value;
        }

        private void ilaVAT_APPLY_BASE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_APPLY_BASE", "Y");
        }

        private void ilaVAT_SUBMIT_IMPOSSIBLE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_SUBMIT_IMPOSSIBLE", "Y");
        }

        private void ilaVAT_SUBMIT_DOCUMENT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_SUBMIT_DOCUMENT", "Y");
        }

        private void ilaVAT_APPLY_BASE_0_SelectedRowData(object pSender)
        {
            if (idaFOREIGN_CURR_HEADER.OraSelectData.Rows.Count > 0)
            {
                APPLY_BASE_CODE.EditValue = APPLY_BASE_CODE_0.EditValue;
                APPLY_BASE_DESC.EditValue = APPLY_BASE_DESC_0.EditValue;
            }
        }

        private void ilaVAT_SUBMIT_IMPOSSIBLE_0_SelectedRowData(object pSender)
        {
            if (idaFOREIGN_CURR_HEADER.OraSelectData.Rows.Count > 0)
            {
                SUBMIT_IMPOSSIBLE_CODE.EditValue = SUBMIT_IMPOSSIBLE_CODE_0.EditValue;
                SUBMIT_IMPOSSIBLE_DESC.EditValue = SUBMIT_IMPOSSIBLE_DESC_0.EditValue;
            }
        }

        private void ilaVAT_SUBMIT_DOCUMENT_0_SelectedRowData(object pSender)
        {
            if (idaFOREIGN_CURR_HEADER.OraSelectData.Rows.Count > 0)
            {
                SUBMIT_DOCUMENT_CODE.EditValue = SUBMIT_DOCUMENT_CODE_0.EditValue;
                SUBMIT_DOCUMENT_DESC.EditValue = SUBMIT_DOCUMENT_DESC_0.EditValue;
            }
        }

        #endregion

        #region ----- Lookup Event : Grid -----
        
        #endregion

        #region ----- Adapter Event : VAT_FOREIGN_CURR -----

        private void idaVAT_EXPORT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["TAX_CODE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DOCUMENT_NUM"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10280"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["SHIPPING_DATE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10242", "&&VALUE:=Shipping Date(선적일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CURRENCY_CODE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10124"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EXCHANGE_RATE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10268"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaVAT_EXPORT_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        #endregion 

    }
}