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

namespace FCMF0516
{
    public partial class FCMF0516 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        #endregion;

        #region ----- Constructor -----

        public FCMF0516()
        {
            InitializeComponent();
        }

        public FCMF0516(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (iString.ISNull(DUE_DATE_FR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DUE_DATE_FR_0.Focus();
                return;
            }
            if (iString.ISNull(DUE_DATE_TO_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DUE_DATE_TO_0.Focus();
                return;
            }
            if (Convert.ToDateTime(DUE_DATE_FR_0.EditValue) > Convert.ToDateTime(DUE_DATE_TO_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DUE_DATE_FR_0.Focus();
                return;
            }
            if(itbBILL.SelectedTab.TabIndex == 1)
            {//지급어음
                idaPAYABLE_BILL.SetSelectParamValue("W_BILL_CLASS", "1");
                idaPAYABLE_BILL.Fill();

                igrPAYABLE_BILL.Focus();
            }
            else if(itbBILL.SelectedTab.TabIndex == 2)
            {//받을어음.
                idaRECEIVABLE_BILL.SetSelectParamValue("W_BILL_CLASS", "2");
                idaRECEIVABLE_BILL.Fill();

                igrRECEIVABLE_BILL.Focus();
            }
            
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Default_Value()
        {
            DUE_DATE_FR_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            DUE_DATE_TO_0.EditValue = DateTime.Today;

            // 어음상태.
            idcDV_COMMON.SetCommandParamValue("W_GROUP_CODE", "BILL_STATUS");
            idcDV_COMMON.ExecuteNonQuery();
            BILL_STATUS_NAME_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE_NAME");
            BILL_STATUS_0.EditValue = idcDV_COMMON.GetCommandParamValue("O_CODE");
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private object GetTerritory()
        {

            object vTerritory = "Default";
            vTerritory = isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage;
            return vTerritory;
        }

        #endregion;

        #region ----- XL Print Methods ----

        private void XLPrinting(string pOutChoice, ISGridAdvEx pGRID)
        {// pOutChoice : 출력구분.
            
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            string pDUE_DATE_FR = iDate.ISGetDate(DUE_DATE_FR_0.EditValue).ToShortDateString();
            string pDUE_DATE_TO = iDate.ISGetDate(DUE_DATE_TO_0.EditValue).ToShortDateString();
            int pTabIndex = itbBILL.SelectedTab.TabIndex;

            object vTerritory = string.Empty;

            int vCountRow = pGRID.RowCount;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutChoice.ToUpper() == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                if (pTabIndex == 1)
                {
                    vSaveFileName = String.Format("어음명세서(지급어음명세서)_{0}~{1}", pDUE_DATE_FR, pDUE_DATE_TO);
                }
                else if (pTabIndex == 2)
                {
                    vSaveFileName = String.Format("어음명세서(받을어음명세서)_{0}~{1}", pDUE_DATE_FR, pDUE_DATE_TO);
                }
                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xlsx";

                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    if (vFileName.Exists)
                    {
                        try
                        {
                            vFileName.Delete();
                        }
                        catch (Exception EX)
                        {
                            MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            vTerritory = GetTerritory();
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.

                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0516_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //인쇄 일시 - 서버 시간
                    IDC_GetDate.ExecuteNonQuery();
                    DateTime vPrint_Datetime = iDate.ISGetDate(IDC_GetDate.GetCommandParamValue("X_LOCAL_DATE"));

                    // 실제 인쇄
                    vPageNumber = xlPrinting.WriteMain(pTabIndex, pDUE_DATE_FR, pDUE_DATE_TO, vPrint_Datetime, pGRID);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        ////출력구분에 따른 선택(인쇄 or file 저장)
                        //PRINT TYPE : PREVIEW, PRINT
                        //START : 인쇄 시작 범위, END : 인쇄 종료 범위, PRINTCOPIES : 인쇄 매수 
                        xlPrinting.Printing("PREVIEW", 1, vPageNumber, 1);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = "Printing End";
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

        #region ----- XL Export Methods ----

        private void XLExport(ISGridAdvEx pGRID)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;
            string pDUE_DATE_FR_0 = DUE_DATE_FR_0.DateTimeValue.ToShortDateString();
            string pDUE_DATE_TO_0 = DUE_DATE_TO_0.DateTimeValue.ToShortDateString();


            object vTerritory = string.Empty;

            int vCountRow = pGRID.RowCount;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

            if (itbBILL.SelectedTab.TabIndex == 1)
            {
                vSaveFileName = String.Format("어음명세서(지급어음명세서)_{0}~{1}", pDUE_DATE_FR_0, pDUE_DATE_TO_0);
            }
            else if (itbBILL.SelectedTab.TabIndex == 2)
            {
                vSaveFileName = String.Format("어음명세서(받을어음명세서)_{0}~{1}", pDUE_DATE_FR_0, pDUE_DATE_TO_0);
            }
            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.DefaultExt = "xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = saveFileDialog1.FileName;
                System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                if (vFileName.Exists)
                {
                    try
                    {
                        vFileName.Delete();
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            vTerritory = GetTerritory();
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            bool isOpen = false;
            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.

                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0516_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                isOpen = false;

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            //-------------------------------------------------------------------------------------
            if (isOpen == true)
            {
                try
                {
                    if (idaPAYABLE_BILL.SelectRows.Count > 0)
                    {
                        xlPrinting.Header_ExportWrite(pDUE_DATE_FR_0, pDUE_DATE_TO_0);
                    }

                    // 실제 인쇄
                    vPageNumber = xlPrinting.ExportWrite(iString.ISNull(vTerritory), pGRID);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    xlPrinting.SAVE(vSaveFileName);

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = "Printing End";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                catch (Exception Ex)
                {
                    xlPrinting.Dispose();

                    vMessageText = Ex.Message;
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            else
            {
                vMessageText = "Excel File Open Error";
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            //-------------------------------------------------------------------------------------

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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPAYABLE_BILL.IsFocused)
                    {
                        idaPAYABLE_BILL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPAYABLE_BILL.IsFocused)
                    {
                        idaPAYABLE_BILL.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (itbBILL.SelectedTab.TabIndex == 1)
                    {
                        XLPrinting("PRINT", igrPAYABLE_BILL);
                    }
                    else if (itbBILL.SelectedTab.TabIndex == 2)
                    {
                        XLPrinting("PRINT", igrRECEIVABLE_BILL);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (itbBILL.SelectedTab.TabIndex == 1)
                    {
                        XLExport(igrPAYABLE_BILL);
                    }
                    else if (itbBILL.SelectedTab.TabIndex == 2)
                    {
                        XLExport(igrRECEIVABLE_BILL);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0516_Load(object sender, EventArgs e)
        {

        }

        private void FCMF0516_Shown(object sender, EventArgs e)
        {
            Set_Default_Value();
        }

        private void itbBILL_SelectedIndexChanged(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        #endregion

        #region ------ Lookup Event -----

        private void ilaVENDOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVENDOR.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaBILL_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BILL_STATUS", "N");
        }

        private void ilaBANK_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

    }
}