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

namespace FCMF0822
{
    public partial class FCMF0822 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mVAT_DOC_TYPE = "2";
        object mVAT_DOC_TYPE_DESC;

        #endregion;

        #region ----- Constructor -----

        public FCMF0822()
        {
            InitializeComponent();
        }

        public FCMF0822(Form pMainForm, ISAppInterface pAppInterface)
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

            //DEFAULT DOCUMENT TYPE.
            idcDV_VAT_DOC_TYPE.SetCommandParamValue("W_GROUP_CODE", "VAT_DOC_TYPE");
            idcDV_VAT_DOC_TYPE.SetCommandParamValue("W_CODE", mVAT_DOC_TYPE);
            idcDV_VAT_DOC_TYPE.ExecuteNonQuery();
            mVAT_DOC_TYPE_DESC = idcDV_VAT_DOC_TYPE.GetCommandParamValue("O_RETURN_VALUE");

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
            IDA_ZERO_RATE_EISSUE_SUM.Fill();
            IDA_ZERO_RATE_EISSUE.Fill();
            IDA_PRINT_EISSUE.Fill();
            IGR_ZERO_RATE_EISSUE.Focus();
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

        private void SET_INSERT_EISSUE()
        {
            IGR_ZERO_RATE_EISSUE.SetCellValue("TAX_CODE", W_TAX_CODE.EditValue);
            IGR_ZERO_RATE_EISSUE.SetCellValue("TAX_CODE_DESC", W_TAX_CODE_NAME.EditValue);
            IGR_ZERO_RATE_EISSUE.SetCellValue("VAT_DATE_FR", W_ISSUE_DATE_FR.EditValue);
            IGR_ZERO_RATE_EISSUE.SetCellValue("VAT_DATE_TO", W_ISSUE_DATE_TO.EditValue);
            IGR_ZERO_RATE_EISSUE.SetCellValue("DOCUMENT_TYPE", mVAT_DOC_TYPE);
            IGR_ZERO_RATE_EISSUE.SetCellValue("DOCUMENT_TYPE_DESC", mVAT_DOC_TYPE_DESC);
            IGR_ZERO_RATE_EISSUE.SetCellValue("ISSUE_DATE", W_WRITE_DATE.EditValue);

            IGR_ZERO_RATE_EISSUE.CurrentCellMoveTo(IGR_ZERO_RATE_EISSUE.GetColumnToIndex("DOCUMENT_TYPE_DESC"));
            IGR_ZERO_RATE_EISSUE.CurrentCellActivate(IGR_ZERO_RATE_EISSUE.GetColumnToIndex("DOCUMENT_TYPE_DESC"));
            IGR_ZERO_RATE_EISSUE.Focus();
        }

        private void SET_TOTAL_BASE_AMOUNT(object pEXCHANGE_RATE, object pCURR_AMOUNT)
        {            
            decimal mBASE_AMOUNT = 0;        
            mBASE_AMOUNT = iString.ISDecimaltoZero(pEXCHANGE_RATE) * iString.ISDecimaltoZero(pCURR_AMOUNT);
            mBASE_AMOUNT = Math.Round(mBASE_AMOUNT, 0);
            IGR_ZERO_RATE_EISSUE.SetCellValue("TOTAL_BASE_AMOUNT", mBASE_AMOUNT);            
        }

        private void SET_THIS_BASE_AMOUNT(object pEXCHANGE_RATE, object pCURR_AMOUNT)
        {
            decimal mBASE_AMOUNT = 0;
            mBASE_AMOUNT = iString.ISDecimaltoZero(pEXCHANGE_RATE) * iString.ISDecimaltoZero(pCURR_AMOUNT);
            mBASE_AMOUNT = Math.Round(mBASE_AMOUNT, 0);
            IGR_ZERO_RATE_EISSUE.SetCellValue("THIS_BASE_AMOUNT", mBASE_AMOUNT);
        }

        private void Show_Slip_Detail()
        {
            int mSLIP_HEADER_ID = iString.ISNumtoZero(IGR_ZERO_RATE_EISSUE.GetCellValue("INTERFACE_HEADER_ID"));
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

        #region ----- XL Export Methods ----

        private void ExportXL()
        {
            int vCountRow = IDA_ZERO_RATE_EISSUE.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xlsx";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(IDA_ZERO_RATE_EISSUE.OraSelectData, vsSaveExcelFileName, vsSheetName);
                if (vXLSaveOK == true)
                {
                    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                else
                {
                    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                xlExport.XLClose();
            }
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = -1;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 0;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 5;
                    break;
            }

            return vTerritory;
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
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                idcVAT_PERIOD.ExecuteNonQuery();
                string vPeriod = string.Format("( {0} )", idcVAT_PERIOD.GetCommandParamValue("O_PERIOD"));                
                string vISSUE_PERIOD = String.Format("{0:D4}년 {1:D2}월 {2:D2}일 ~ {3:D4}년 {4:D2}월 {5:D2}일"
                                            , W_ISSUE_DATE_FR.DateTimeValue.Year, W_ISSUE_DATE_FR.DateTimeValue.Month, W_ISSUE_DATE_FR.DateTimeValue.Day
                                            , W_ISSUE_DATE_TO.DateTimeValue.Year, W_ISSUE_DATE_TO.DateTimeValue.Month, W_ISSUE_DATE_TO.DateTimeValue.Day);
                string vWRITE_DATE = String.Format("{0}", W_WRITE_DATE.DateTimeValue.ToShortDateString());
                string vWRITE_DATE_1 = String.Format("{0:D4}년 {1:D2}월 {2:D2}일", W_WRITE_DATE.DateTimeValue.Year, W_WRITE_DATE.DateTimeValue.Month, W_WRITE_DATE.DateTimeValue.Day);

                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0822_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    // 신고자 인적사항 인쇄.
                    idaOPERATING_UNIT.Fill();
                    if (idaOPERATING_UNIT.SelectRows.Count > 0)
                    {
                        xlPrinting.HeaderWrite(idaOPERATING_UNIT, vPeriod, vISSUE_PERIOD, vWRITE_DATE, vWRITE_DATE_1);
                    }
                    // 공급실적합계 인쇄.
                    IDA_PRINT_ZERO_RATE_EISSUE_SUM.Fill();
                    if (IDA_PRINT_ZERO_RATE_EISSUE_SUM.SelectRows.Count > 0)
                    {
                        xlPrinting.Header_SUM(IDA_PRINT_ZERO_RATE_EISSUE_SUM);
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
                        xlPrinting.SAVE("ZERO_RATE_eISSUE_");
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
                    //if (idaZERO_RATE_EISSUE.IsFocused)
                    //{
                    //    idaZERO_RATE_EISSUE.AddOver();
                    //    SET_INSERT_EISSUE();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if (idaZERO_RATE_EISSUE.IsFocused)
                    //{
                    //    idaZERO_RATE_EISSUE.AddUnder();
                    //    SET_INSERT_EISSUE();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_ZERO_RATE_EISSUE.IsFocused)
                    {
                        IDA_ZERO_RATE_EISSUE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ZERO_RATE_EISSUE.IsFocused)
                    {
                        IDA_ZERO_RATE_EISSUE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ZERO_RATE_EISSUE.IsFocused)
                    {
                        IDA_ZERO_RATE_EISSUE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT", igrPRINT_EISSUE);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE", igrPRINT_EISSUE);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0822_Load(object sender, EventArgs e)
        {
            IDA_ZERO_RATE_EISSUE.FillSchema();
        }

        private void FCMF0822_Shown(object sender, EventArgs e)
        {
            Set_Default_Value();
        }

        private void igrZERO_RATE_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        // 사용안함.
        private void igrVAT_ZERO_RATE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int mIDX_EXCHANGE_RATE = IGR_ZERO_RATE_EISSUE.GetColumnToIndex("EXCHANGE_RATE");
            int mIDX_TOTAL_CURR_AMOUNT = IGR_ZERO_RATE_EISSUE.GetColumnToIndex("TOTAL_CURR_AMOUNT");
            int mIDX_THIS_CURR_AMOUNT = IGR_ZERO_RATE_EISSUE.GetColumnToIndex("THIS_CURR_AMOUNT");

            if (e.ColIndex == mIDX_EXCHANGE_RATE)
            {//환율.
                SET_TOTAL_BASE_AMOUNT(e.NewValue, IGR_ZERO_RATE_EISSUE.GetCellValue("TOTAL_CURR_AMOUNT"));
                // 전호수 수정 : 당기체출금액만 수정.
                //SET_THIS_BASE_AMOUNT(e.NewValue, igrZERO_TAX_RATE.GetCellValue("TOTAL_CURR_AMOUNT"));
            }
            else if (e.ColIndex == mIDX_TOTAL_CURR_AMOUNT)
            {//당기제출금액.
                SET_TOTAL_BASE_AMOUNT(IGR_ZERO_RATE_EISSUE.GetCellValue("EXCHANGE_RATE"), e.NewValue);
            }
            // 전호수 수정 : 당기체출금액만 수정.
            //else if (e.ColIndex == mIDX_THIS_CURR_AMOUNT)
            //{//당기제출금액.
            //    SET_THIS_BASE_AMOUNT(igrZERO_TAX_RATE.GetCellValue("EXCHANGE_RATE"), e.NewValue);
            //}
        }

        
        private void ibtnSET_EISSUE_ButtonClick(object pSender, EventArgs pEventArgs)
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
            idcSET_EISSUE.ExecuteNonQuery();
            mSTATUS = iString.ISNull(idcSET_EISSUE.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(idcSET_EISSUE.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
            if (idcSET_EISSUE.ExcuteError || mSTATUS == "F")
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
        
        #endregion

        #region ----- Lookup Event : Search -----

        private void ilaTAX_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        } 

        private void ilaTAX_CODE_0_SelectedRowData(object pSender)
        {
            W_VAT_PERIOD_ID.EditValue = string.Empty;
            W_VAT_PERIOD_DESC.EditValue = DBNull.Value;
            W_ISSUE_DATE_FR.EditValue = DBNull.Value;
            W_ISSUE_DATE_TO.EditValue = DBNull.Value;
        }

        private void ilaVAT_DOCU_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_DOC_TYPE", "Y");
        }

        private void ilaCUSTOMER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event : ZERO_RATE_EISSUE -----

        private void idaZERO_RATE_EISSUE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["TAX_CODE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["DOCUMENT_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10129"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ISSUE_DATE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10242", "&&VALUE:=Issue Date(발급일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DOCUMENT_NUM"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10240"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (iString.ISNull(e.Row["CUSTOMER_ID"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaZERO_RATE_EISSUE_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        #endregion
          
    }
}