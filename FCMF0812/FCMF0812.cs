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

namespace FCMF0812
{
    public partial class FCMF0812 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;M

        #region ----- Constructor -----

        public FCMF0812()
        {
            InitializeComponent();
        }

        public FCMF0812(Form pMainForm, ISAppInterface pAppInterface)
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

        private void Default_Value()
        {
            //세금계산서 발행기간.
            DateTime vGetDateTime = GetDate();
            W_PERIOD_YEAR.EditValue = iDate.ISYear(vGetDateTime);
            
            //사업장 구분.
            idcDV_TAX_CODE.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
            idcDV_TAX_CODE.ExecuteNonQuery();

            W_TAX_CODE_NAME1.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE_NAME");
            W_TAX_CODE1.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE");

            W_TAX_CODE_NAME2.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE_NAME");
            W_TAX_CODE2.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE"); 

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
            
            if (TB_MAIN.SelectedTab.TabIndex == TP_REALTY_LEASE.TabIndex)
            {
                if (iString.ISNull(W_TAX_CODE1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_TAX_CODE_NAME1.Focus();
                    return;
                }

                idaREALTY_LEASE.Fill();
                igrREALTY_LEASE.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_HISTORY.TabIndex)
            {
                if (iString.ISNull(W_TAX_CODE2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_TAX_CODE_NAME2.Focus();
                    return;
                }

                if (iString.ISNull(W_VAT_PERIOD_ID.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10487"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_VAT_PERIOD_DESC.Focus();
                    return;
                }
                if (iString.ISNull(W_ISSUE_DATE_FR2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_ISSUE_DATE_FR2.Focus();
                    return;
                }
                if (iString.ISNull(W_ISSUE_DATE_TO2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_ISSUE_DATE_FR2.Focus();
                    return;
                }
                if (Convert.ToDateTime(W_ISSUE_DATE_FR2.EditValue) > Convert.ToDateTime(W_ISSUE_DATE_TO2.EditValue))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_ISSUE_DATE_FR2.Focus();
                    return;
                }

                //신고기간 검증.
                if (VAT_PERIOD_CHECK() == false)
                {
                    return;
                }                        
            
                idaREALTY_LEASE_HISTORY.Fill();
                igrREALTY_LEASE_HISTORY.Focus();
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
                W_ISSUE_DATE_FR2.Focus();
                return false;
            }
            return true;
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void INIT_INSERT_GRID()
        {
            igrREALTY_LEASE.SetCellValue("TAX_CODE", W_TAX_CODE1.EditValue);
            igrREALTY_LEASE.SetCellValue("TAX_NAME", W_TAX_CODE_NAME1.EditValue);

            USE_DATE_FR.EditValue = DateTime.Today;
            USE_DATE_TO.EditValue = DateTime.Today;
            
            HOUSE_NUM.Focus();
        }

        #endregion;

        #region ----- XL Export Methods ----

        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.OraSelectData.Rows.Count;
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
                bool vXLSaveOK = xlExport.XLExport(idaREALTY_LEASE.OraSelectData, vsSaveExcelFileName, vsSheetName);
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
            //신고기간 검증.
            if (VAT_PERIOD_CHECK() == false)
            {
                return;
            }
            
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
                string vUSE_PERIOD = String.Format("(기간 : {0:D2} 월 ~ {1:D2} 월)", W_ISSUE_DATE_FR2.DateTimeValue.Month, W_ISSUE_DATE_TO2.DateTimeValue.Month);
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0812_001.xlsx";
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
                        xlPrinting.HeaderWrite(idaOPERATING_UNIT, vPeriod, vUSE_PERIOD);
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
                        xlPrinting.SAVE("Realty_");
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
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaREALTY_LEASE.IsFocused)
                    {
                        if (iString.ISNull(W_TAX_CODE1.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            W_TAX_CODE1.Focus();
                            return;
                        }
                        idaREALTY_LEASE.AddOver();
                        INIT_INSERT_GRID();
                    }
                    else if (idaREALTY_LEASE_HISTORY.IsFocused)
                    {
                        idaREALTY_LEASE_HISTORY.AddOver();

                        igrREALTY_LEASE_HISTORY.CurrentCellMoveTo(igrREALTY_LEASE_HISTORY.GetColumnToIndex("CUSTOMER_DESC"));
                        igrREALTY_LEASE_HISTORY.CurrentCellActivate(igrREALTY_LEASE_HISTORY.GetColumnToIndex("CUSTOMER_DESC"));
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaREALTY_LEASE.IsFocused)
                    {
                        if (iString.ISNull(W_TAX_CODE1.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            W_TAX_CODE1.Focus();
                            return;
                        }
                        idaREALTY_LEASE.AddUnder();
                        INIT_INSERT_GRID();
                    }
                    else if (idaREALTY_LEASE_HISTORY.IsFocused)
                    {
                        idaREALTY_LEASE_HISTORY.AddUnder();

                        igrREALTY_LEASE_HISTORY.CurrentCellMoveTo(igrREALTY_LEASE_HISTORY.GetColumnToIndex("CUSTOMER_DESC"));
                        igrREALTY_LEASE_HISTORY.CurrentCellActivate(igrREALTY_LEASE_HISTORY.GetColumnToIndex("CUSTOMER_DESC"));
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaREALTY_LEASE.IsFocused)
                    {
                        idaREALTY_LEASE.Update();
                    }
                    else if (idaREALTY_LEASE_HISTORY.IsFocused)
                    {
                        idaREALTY_LEASE_HISTORY.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaREALTY_LEASE.IsFocused)
                    {
                        idaREALTY_LEASE.Cancel();
                    }
                    else if (idaREALTY_LEASE_HISTORY.IsFocused)
                    {
                        idaREALTY_LEASE_HISTORY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaREALTY_LEASE.IsFocused)
                    {
                        idaREALTY_LEASE.Delete();
                    }
                    else if (idaREALTY_LEASE_HISTORY.IsFocused)
                    {
                        idaREALTY_LEASE_HISTORY.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT", igrREALTY_LEASE_HISTORY);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE", igrREALTY_LEASE_HISTORY);
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void FCMF0812_Load(object sender, EventArgs e)
        {
            idaREALTY_LEASE.FillSchema(); 
        }

        private void FCMF0812_Shown(object sender, EventArgs e)
        {
            Default_Value();
        }

        private void ibtnSET_INTEREST_RATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult dlgResult;

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            FCMF0812_INTEREST_RATE vINTEREST_RATE = new FCMF0812_INTEREST_RATE(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            dlgResult = vINTEREST_RATE.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {

            }
            vINTEREST_RATE.Dispose();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void ibtnSET_REALTY_LEASE_HISTORY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_TAX_CODE2.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TAX_CODE_NAME2.Focus();
                return;
            }

            if (iString.ISNull(W_VAT_PERIOD_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10487"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_FR2.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR2.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_TO2.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_TO2.Focus();
                return;
            }
            if (Convert.ToDateTime(W_ISSUE_DATE_FR2.EditValue) > Convert.ToDateTime(W_ISSUE_DATE_TO2.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR2.Focus();
                return;
            }

            //신고기간 검증.
            if (VAT_PERIOD_CHECK() == false)
            {
                return;
            }

            string mSTATUS = "F";
            string mMESSAGE = null;

            idcSET_REALTY_LEASE.ExecuteNonQuery();
            mSTATUS = iString.ISNull(idcSET_REALTY_LEASE.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(idcSET_REALTY_LEASE.GetCommandParamValue("O_MESSAGE"));
            if (idcSET_REALTY_LEASE.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (mMESSAGE != String.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion

        #region ------ Lookup Event ------
        
        private void ilaTAX_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "N");
        }

        private void ilaTAX_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        }

        private void ilaTAX_CODE2_0_SelectedRowData(object pSender)
        {
            W_VAT_PERIOD_DESC.EditValue = string.Empty;
            W_VAT_PERIOD_ID.EditValue = string.Empty;
            W_ISSUE_DATE_FR2.EditValue = DBNull.Value;
            W_ISSUE_DATE_TO2.EditValue = DBNull.Value;
        }

        private void ilaCUSTOMER_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaFLOOR_TYPE1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR_TYPE", "Y");
        }

        private void ilaCUSTOMER1_SelectedRowData(object pSender)
        {
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ------ Adater Event : REALTY LEASE ------

        private void idaREALTY_LEASE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["TAX_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["HOUSE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10294"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FLOOR_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10283"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FLOOR_COUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10284"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISInt(e.Row["FLOOR_COUNT"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10580", string.Format("&&VALUE:=Number of floors(층수)")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }     
            if (iString.ISNull(e.Row["CUSTOMER_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10290"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AREA_M2"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10286"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["USE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10288"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USE_DATE_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10289"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaREALTY_LEASE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion
         
    }
}