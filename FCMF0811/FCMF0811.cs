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

namespace FCMF0811
{
    public partial class FCMF0811 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0811()
        {
            InitializeComponent();
        }

        public FCMF0811(Form pMainForm, ISAppInterface pAppInterface)
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

            idaNOT_DEDUCTION.Fill();
            idaNOT_DEDUCTION_ADJUST.Fill();
            idaNOT_DEDUCTION_DETAIL.Fill();

            if (itbNOT_DEDUCTION.SelectedTab.TabIndex == 1)
            {
                igrNOT_DEDUCTION.Focus();
            }
            else if (itbNOT_DEDUCTION.SelectedTab.TabIndex == 2)
            {
                igrNOT_DEDUCTION_DETAIL.Focus();
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

        private void Set_GRID_STATUS_ROW()
        {
            if (igrNOT_DEDUCTION_ADJUST.RowCount < 1)
            {
                return;
            }
            int vSTATUS = 0;                // INSERTABLE, UPDATABLE;

            int vROW = igrNOT_DEDUCTION_ADJUST.RowIndex;
            object vNO_DED_CODE = igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_CODE");
            int vIDX_GL_AMOUNT = igrNOT_DEDUCTION_ADJUST.GetColumnToIndex("GL_AMOUNT");
            int vIDX_VAT_AMOUNT = igrNOT_DEDUCTION_ADJUST.GetColumnToIndex("VAT_AMOUNT");

            if (iString.ISNull(vNO_DED_CODE) == "990")
            {
                vSTATUS = 0;
            }
            else
            {
                vSTATUS = 1;
            }
            
            igrNOT_DEDUCTION_ADJUST.GridAdvExColElement[vIDX_GL_AMOUNT].Insertable = vSTATUS;
            igrNOT_DEDUCTION_ADJUST.GridAdvExColElement[vIDX_GL_AMOUNT].Updatable = vSTATUS;

            igrNOT_DEDUCTION_ADJUST.GridAdvExColElement[vIDX_VAT_AMOUNT].Insertable = vSTATUS;
            igrNOT_DEDUCTION_ADJUST.GridAdvExColElement[vIDX_VAT_AMOUNT].Updatable = vSTATUS;
        }

        private void SHOW_ADJUST_3()
        {
            FCMF0811_3 vFCMF0811_3 = new FCMF0811_3(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                    , W_TAX_CODE_NAME.EditValue, W_TAX_CODE.EditValue
                                                    , W_VAT_PERIOD_DESC.EditValue
                                                    , W_ISSUE_DATE_FR.EditValue
                                                    , W_ISSUE_DATE_TO.EditValue
                                                    , igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_TYPE")
                                                    , igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_DESC")
                                                    , igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_CODE"));
            vFCMF0811_3.ShowDialog();
            if (iString.ISNull(vFCMF0811_3.Get_Save_Flag) == "SAVE")
            {
                IDC_GET_ADJUST_AMOUNT.SetCommandParamValue("P_NOT_DED_TYPE", igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_TYPE"));
                IDC_GET_ADJUST_AMOUNT.SetCommandParamValue("P_NOT_DED_CODE", igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_CODE"));
                IDC_GET_ADJUST_AMOUNT.SetCommandParamValue("P_ADJUST_TYPE", "3");
                IDC_GET_ADJUST_AMOUNT.ExecuteNonQuery();
                decimal vADJUST_AMOUNT = iString.ISDecimaltoZero(IDC_GET_ADJUST_AMOUNT.GetCommandParamValue("O_ADJUST_AMOUNT"));
                if (IDC_GET_ADJUST_AMOUNT.ExcuteError)
                {
                    vADJUST_AMOUNT = 0;
                }

                decimal vVAT_RATE = 0.1M;
                decimal vVAT_AMOUNT = Math.Truncate(vADJUST_AMOUNT * vVAT_RATE);

                igrNOT_DEDUCTION_ADJUST.SetCellValue("GL_AMOUNT", vADJUST_AMOUNT);
                igrNOT_DEDUCTION_ADJUST.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);

                idaNOT_DEDUCTION_ADJUST.Update();
            }

            vFCMF0811_3.Dispose();
        }

        private void SHOW_ADJUST_4()
        {
            FCMF0811_4 vFCMF0811_4 = new FCMF0811_4(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                    , W_TAX_CODE_NAME.EditValue, W_TAX_CODE.EditValue
                                                    , W_VAT_PERIOD_DESC.EditValue
                                                    , W_ISSUE_DATE_FR.EditValue
                                                    , W_ISSUE_DATE_TO.EditValue
                                                    , igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_TYPE")
                                                    , igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_DESC")
                                                    , igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_CODE"));
            vFCMF0811_4.ShowDialog();
            if (iString.ISNull(vFCMF0811_4.Get_Save_Flag) == "SAVE")
            {
                IDC_GET_ADJUST_AMOUNT.SetCommandParamValue("P_NOT_DED_TYPE", igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_TYPE"));
                IDC_GET_ADJUST_AMOUNT.SetCommandParamValue("P_NOT_DED_CODE", igrNOT_DEDUCTION_ADJUST.GetCellValue("NOT_DED_CODE"));
                IDC_GET_ADJUST_AMOUNT.SetCommandParamValue("P_ADJUST_TYPE", "4");
                IDC_GET_ADJUST_AMOUNT.ExecuteNonQuery();
                decimal vADJUST_AMOUNT = iString.ISDecimaltoZero(IDC_GET_ADJUST_AMOUNT.GetCommandParamValue("O_ADJUST_AMOUNT"));
                if (IDC_GET_ADJUST_AMOUNT.ExcuteError)
                {
                    vADJUST_AMOUNT = 0;
                }

                decimal vVAT_RATE = 0.1M;
                decimal vVAT_AMOUNT = Math.Floor(vADJUST_AMOUNT * vVAT_RATE);

                igrNOT_DEDUCTION_ADJUST.SetCellValue("GL_AMOUNT", vADJUST_AMOUNT);
                igrNOT_DEDUCTION_ADJUST.SetCellValue("VAT_AMOUNT", vVAT_AMOUNT);
                
                idaNOT_DEDUCTION_ADJUST.Update();
            }

            vFCMF0811_4.Dispose();
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
                string vPeriod = string.Format("( {0} )", idcVAT_PERIOD.GetCommandParamValue("O_PERIOD"));
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0811_001.xls";
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
                        xlPrinting.HeaderWrite(idaOPERATING_UNIT, vPeriod);
                    }
                    // 실제 인쇄
                    vPageNumber = xlPrinting.LineWrite(pGrid, vPeriod);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE("VAT_NOT_DED_");
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
                    if (idaNOT_DEDUCTION_DETAIL.IsFocused)
                    {
                        idaNOT_DEDUCTION_DETAIL.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaNOT_DEDUCTION_DETAIL.IsFocused)
                    {
                        idaNOT_DEDUCTION_DETAIL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaNOT_DEDUCTION_DETAIL.IsFocused)
                    {
                        idaNOT_DEDUCTION_DETAIL.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT", igrNOT_DEDUCTION);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE", igrNOT_DEDUCTION);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0811_Load(object sender, EventArgs e)
        {
            idaNOT_DEDUCTION.FillSchema();
            idaNOT_DEDUCTION_DETAIL.FillSchema();
        }

        private void FCMF0811_Shown(object sender, EventArgs e)
        {
            Set_Default_Value();
        }
        
        private void itbNOT_DEDUCTION_Click(object sender, EventArgs e)
        {
            if (itbNOT_DEDUCTION.SelectedTab.TabIndex == 1)
            {
                igrNOT_DEDUCTION.Focus();
            }
            else if (itbNOT_DEDUCTION.SelectedTab.TabIndex == 2)
            {
                igrNOT_DEDUCTION_DETAIL.Focus();
            }
        }

        private void igrNOT_DEDUCTION_ADJUST_CellDoubleClick(object pSender)
        {
            if (iString.ISNull(igrNOT_DEDUCTION_ADJUST.GetCellValue("NO_DED_CODE")) == "110")
            {
                SHOW_ADJUST_3();
            }
            else if (iString.ISNull(igrNOT_DEDUCTION_ADJUST.GetCellValue("NO_DED_CODE")) == "120")
            {
                SHOW_ADJUST_4();
            }
        }

        private void igrNOT_DEDUCTION_ADJUST_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            if (igrNOT_DEDUCTION_ADJUST.RowCount < 1)
            {
                return;
            }
            
            decimal vAMOUNT = 0;
            int vIDX_GL_AMOUNT = igrNOT_DEDUCTION_ADJUST.GetColumnToIndex("GL_AMOUNT");
            int vIDX_VAT_AMOUNT = igrNOT_DEDUCTION_ADJUST.GetColumnToIndex("VAT_AMOUNT");

            Decimal vGL_RATE = iString.ISDecimaltoZero(10);
            Decimal vVAT_RATE = iString.ISDecimaltoZero(0.1);

            if (e.ColIndex == vIDX_GL_AMOUNT)
            {
                if (iString.ISDecimaltoZero(igrNOT_DEDUCTION_ADJUST.GetCellValue("VAT_AMOUNT"), 0) == 0)
                {
                    vAMOUNT = vVAT_RATE * iString.ISDecimaltoZero(e.CellValue, 0);
                    igrNOT_DEDUCTION_ADJUST.SetCellValue("VAT_AMOUNT", vAMOUNT);
                }
            }
            else if (e.ColIndex == vIDX_VAT_AMOUNT)
            {
                if (iString.ISDecimaltoZero(igrNOT_DEDUCTION_ADJUST.GetCellValue("GL_AMOUNT"), 0) == 0)
                {
                    vAMOUNT = vGL_RATE * iString.ISDecimaltoZero(e.CellValue, 0);
                    igrNOT_DEDUCTION_ADJUST.SetCellValue("GL_AMOUNT", vAMOUNT);
                }
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

        #endregion

        #region ----- Adapter Event : Depreciation Asset Detail ------

        //private void idaDPR_ASSET_DETAIL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        //{
        //    if (iString.ISNull(e.Row["DPR_ASSET_ID"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10232"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["ACQUIRE_DATE"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10203"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["CUSTOMER_ID"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10135"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["GL_AMOUNT"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10208"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["VAT_AMOUNT"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10281"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["VAT_ASSET_GB"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10282"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        //private void idaDPR_ASSET_DETAIL_PreDelete(ISPreDeleteEventArgs e)
        //{
        //    if (e.Row.RowState != DataRowState.Added)
        //    {
        //        if (iString.ISNull(e.Row["DPR_ASSET_ID"]) == string.Empty)
        //        {
        //            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10232"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            e.Cancel = true;
        //            return;
        //        }
        //    }
        //}

        private void idaNOT_DEDUCTION_ADJUST_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            //Set_GRID_STATUS_ROW();
        }

        #endregion

    }
}