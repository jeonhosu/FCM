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

namespace FCMF0813
{
    public partial class FCMF0813 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0813()
        {
            InitializeComponent();
        }

        public FCMF0813(Form pMainForm, ISAppInterface pAppInterface)
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
            idcDV_TAX_CODE.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
            idcDV_TAX_CODE.ExecuteNonQuery();
            W_TAX_CODE_NAME.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE_NAME");
            W_TAX_CODE.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE");

            WRITE_DATE.EditValue = vGetDateTime;

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

            //idaBUSINESS_MASTER.Fill();
            idaVAT_DECLARATION.Fill();
            idaTAX_STANDARD.Fill();
            if (itbVAT_DECLARATION.SelectedTab.TabIndex == 1)
            {
                ISSUE_PERIOD_FR.Focus();
            }
            else if (itbVAT_DECLARATION.SelectedTab.TabIndex == 2)
            {
                BALANCE_TAX_VAT.Focus();
            }
            else if (itbVAT_DECLARATION.SelectedTab.TabIndex == 3)
            {
                SS_TAX_INVOICE_AMT.Focus();
            }
            else if (itbVAT_DECLARATION.SelectedTab.TabIndex == 4)
            {
                igrTAX_STANDARD.Focus();
            }
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

        private void XLPrinting_1(string pOutChoice, ISDataAdapter pData1, ISDataAdapter pData2)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = pData1.OraSelectData.Rows.Count;

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

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                idcVAT_PERIOD.ExecuteNonQuery();
                string vPeriod = string.Format("( {0} )", idcVAT_PERIOD.GetCommandParamValue("O_PERIOD"));
                string vISSUE_PERIOD = String.Format("({0:D2}월 {1:D2}일 ~ {2:D2}월 {3:D2}일)", ISSUE_PERIOD_FR.DateTimeValue.Month, ISSUE_PERIOD_FR.DateTimeValue.Day, ISSUE_DATE_TO.DateTimeValue.Month, ISSUE_DATE_TO.DateTimeValue.Day);
                string vWrite_Date = String.Format("{0:D2}년 {1:D2}월 {2:D2}일", iDate.ISGetDate(WRITE_DATE.EditValue).Year, iDate.ISGetDate(WRITE_DATE.EditValue).Month, iDate.ISGetDate(WRITE_DATE.EditValue).Day);

                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2014-01-01"))
                {
                    xlPrinting.OpenFileNameExcel = "FCMF0813_001.xlsx";                    
                }
                else if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2016-01-01"))
                {
                    xlPrinting.OpenFileNameExcel = "FCMF0813_002.xlsx";
                }
                else if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2021-04-01"))
                {
                    xlPrinting.OpenFileNameExcel = "FCMF0813_003.xlsx";
                }
                else  
                {
                    //2021년도 1기 확정 변경 양식 적용 
                    xlPrinting.OpenFileNameExcel = "FCMF0813_013.xlsx";                    
                } 
                //-------------------------------------------------------------------------------------
                // 파일 오픈. 
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    // 헤더 인쇄.
                    if (pData1.CurrentRows.Count > 0)
                    {
                        if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2014-01-01"))
                        {
                            xlPrinting.HeaderWrite(pData1, vPeriod, vISSUE_PERIOD);
                        }
                        else if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2021-04-01"))
                        {
                            xlPrinting.HeaderWrite_201401(pData1, vPeriod, vISSUE_PERIOD);
                        }
                        else
                        {
                            xlPrinting.HeaderWrite_013(pData1, vPeriod, vISSUE_PERIOD, vWrite_Date);
                        }
                    }

                    //과세표준인쇄.
                    idaPRINT_TAX_STANDARD.Fill();
                    if (igrPRINT_TAX_STANDARD.RowCount > 0)
                    {
                        if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2014-01-01"))
                        {
                            xlPrinting.XLLine_3(igrPRINT_TAX_STANDARD);
                        }
                        else if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2021-04-01"))
                        {
                            xlPrinting.XLLine_3_201401(igrPRINT_TAX_STANDARD);
                        }
                        else
                        {
                            xlPrinting.XLLine_3_013(igrPRINT_TAX_STANDARD);
                        }
                    }

                    // 실제 인쇄
                    if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2014-01-01"))
                    {
                        vPageNumber = xlPrinting.LineWrite(pData1, pData2);
                    }
                    else if (iDate.ISGetDate(W_ISSUE_DATE_FR.EditValue) < iDate.ISGetDate("2021-04-01"))
                    {
                        vPageNumber = xlPrinting.LineWrite_201401(pData1, pData2);
                    }
                    else
                    {
                        vPageNumber = xlPrinting.LineWrite_013(pData1, pData2);
                    }

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE("VAT_1_");
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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaTAX_STANDARD.IsFocused)
                    {
                        if (iString.ISNull(DECLARATION_ID.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10343"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaTAX_STANDARD.Update();
                    }
                    else
                    {
                        idaVAT_DECLARATION.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaVAT_DECLARATION.IsFocused)
                    {
                        idaVAT_DECLARATION.Cancel();
                    }
                    else if (idaDECLARATION_ATTACH.IsFocused)
                    {
                        idaDECLARATION_ATTACH.Cancel();
                    }
                    else if (idaTAX_STANDARD.IsFocused)
                    {
                        idaTAX_STANDARD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaVAT_DECLARATION.IsFocused)
                    {
                        idaVAT_DECLARATION.Delete();
                    }
                    else if (idaDECLARATION_ATTACH.IsFocused)
                    {
                        idaDECLARATION_ATTACH.Delete();
                    }
                    else if (idaTAX_STANDARD.IsFocused)
                    {
                        idaTAX_STANDARD.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT", idaVAT_DECLARATION, idaDECLARATION_ATTACH);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE", idaVAT_DECLARATION, idaDECLARATION_ATTACH);
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void FCMF0813_Load(object sender, EventArgs e)
        {
            idaVAT_DECLARATION.FillSchema();
            idaTAX_STANDARD.FillSchema();
        }

        private void FCMF0813_Shown(object sender, EventArgs e)
        {
            W_ISSUE_DATE_FR.BringToFront();
            W_ISSUE_DATE_TO.BringToFront();

            Set_Default_Value();
        }

        private void itbVAT_DECLARATION_Click(object sender, EventArgs e)
        {
            if (itbVAT_DECLARATION.SelectedTab.TabIndex == 1)
            {
                ISSUE_PERIOD_FR.Focus();
            }
            else if (itbVAT_DECLARATION.SelectedTab.TabIndex == 2)
            {
                BALANCE_TAX_VAT.Focus();
            }
            else if (itbVAT_DECLARATION.SelectedTab.TabIndex == 3)
            {
                SS_TAX_INVOICE_AMT.Focus();
            }
            else if (itbVAT_DECLARATION.SelectedTab.TabIndex == 4)
            {
                igrTAX_STANDARD.Focus();
            }
        }

        private void ibtnSET_DECLARATION_ButtonClick(object pSender, EventArgs pEventArgs)
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
                W_ISSUE_DATE_TO.Focus();
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

            DialogResult vdlgResult;
            FCMF0813_CREATE vFCMF0813_CREATE = new FCMF0813_CREATE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vFCMF0813_CREATE.ShowDialog();

            if (vdlgResult == DialogResult.OK)
            {
                string mSTATUS = "F";
                string mMESSAGE = null;

                mMESSAGE = isMessageAdapter1.ReturnText("FCM_10376");
                mMESSAGE = string.Format("{0}\n\n{1}", mMESSAGE, isMessageAdapter1.ReturnText("FCM_10377"));
                if (MessageBoxAdv.Show(mMESSAGE, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                isDataTransaction1.BeginTran();
                idcSET_DECLARATION.SetCommandParamValue("P_WRITE_DATE", vFCMF0813_CREATE.WRITE_DATE);
                idcSET_DECLARATION.SetCommandParamValue("W_VAT_REPORT_TYPE", vFCMF0813_CREATE.VAT_REPORT_TYPE);
                idcSET_DECLARATION.SetCommandParamValue("W_VAT_LEVIER_TYPE", vFCMF0813_CREATE.VAT_LEVIER_TYPE);
                idcSET_DECLARATION.SetCommandParamValue("W_MODIFY_DESC", vFCMF0813_CREATE.MODIFY_DESC);
                idcSET_DECLARATION.ExecuteNonQuery();
                mSTATUS = iString.ISNull(idcSET_DECLARATION.GetCommandParamValue("O_STATUS"));
                mMESSAGE = iString.ISNull(idcSET_DECLARATION.GetCommandParamValue("O_MESSAGE"));

                W_MODIFY_DEGREE.EditValue = idcSET_DECLARATION.GetCommandParamValue("O_MODIFY_DEGREE");
                W_MODIFY_DESC.EditValue = idcSET_DECLARATION.GetCommandParamValue("O_MODIFY_DESC");
                W_VAT_REPORT_TYPE_DESC.EditValue = vFCMF0813_CREATE.VAT_REPORT_TYPE_DESC;
                W_VAT_REPORT_TYPE.EditValue = vFCMF0813_CREATE.VAT_REPORT_TYPE;
                
                if (idcSET_DECLARATION.ExcuteError || mSTATUS == "F")
                {
                    isDataTransaction1.RollBack();
                    if (mMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
                isDataTransaction1.Commit();
                if (mMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                SEARCH_DB();
            }
        }

        private void BTN_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(DECLARATION_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10343"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            string vStatus = "F";
            string vMessage = string.Empty;
            IDC_DELETE_DECLARATION.ExecuteNonQuery();
            vStatus = iString.ISNull(IDC_DELETE_DECLARATION.GetCommandParamValue("O_STATUS"));
            vMessage = iString.ISNull(IDC_DELETE_DECLARATION.GetCommandParamValue("O_MESSAGE"));
            if (vStatus == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            SEARCH_DB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaTAX_CODE_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        }

        private void ilaTAX_CODE_0_SelectedRowData(object pSender)
        {
            W_VAT_PERIOD_DESC.EditValue = string.Empty;
            W_VAT_PERIOD_ID.EditValue = string.Empty;
            W_ISSUE_DATE_FR.EditValue = DBNull.Value;
            W_ISSUE_DATE_TO.EditValue = DBNull.Value;
        }

        private void ILA_VAT_REPORT_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("VAT_REPORT_TYPE", "Y");
        }

        private void ILA_VAT_MAKE_GB_SelectedRowData(object pSender)
        {
            if (iString.ISNull(W_VAT_REPORT_TYPE.EditValue) == "02")
            {//수정신고//
                W_MODIFY_DESC.ReadOnly = false;
            }
            else
            {
                W_MODIFY_DESC.EditValue = string.Empty;
                W_MODIFY_DEGREE.EditValue = DBNull.Value;

                W_MODIFY_DESC.ReadOnly = true;
            }
        }

        private void ILA_BANK_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_SITE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BANK_SITE.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_ACCOUNT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event : TAX_STANDARD ------

        private void idaVAT_DECLARATION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["CLOSURE_DATE"]) == string.Empty &&
                iString.ISNull(e.Row["CLOSURE_REASON"]) != string.Empty)
            {
                MessageBoxAdv.Show("폐업사유를 입력할 경우 폐업일자는 필수입니다.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CLOSURE_REASON"]) == string.Empty &&
                iString.ISNull(e.Row["CLOSURE_DATE"]) != string.Empty)
            {
                MessageBoxAdv.Show("폐업일자를 입력할 경우 폐업사유는 필수입니다.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        
        private void idaTAX_STANDARD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
             
        }

        #endregion

    }
}