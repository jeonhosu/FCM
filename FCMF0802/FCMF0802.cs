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

namespace FCMF0802
{
    public partial class FCMF0802 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        #endregion;

        #region ----- Constructor -----

        public FCMF0802(Form pMainForm, ISAppInterface pAppInterface)
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

        private void SearchDB()
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

            idaVAT_MASTER.Fill();
            igrVAT_MASTER.Focus();
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

        private void Show_Slip_Detail()
        {
            //int mSLIP_HEADER_ID = iString.ISNumtoZero(igrVAT_MASTER.GetCellValue("SLIP_HEADER_ID"));
            //if (mSLIP_HEADER_ID != Convert.ToInt32(0))
            //{
            //    Application.UseWaitCursor = true;
            //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            //    FCMF0204.FCMF0204 vFCMF0204 = new FCMF0204.FCMF0204(this.MdiParent, isAppInterfaceAdv1.AppInterface, mSLIP_HEADER_ID);
            //    vFCMF0204.Show();

            //    this.Cursor = System.Windows.Forms.Cursors.Default;
            //    Application.UseWaitCursor = false;
            //}
        }

        private void Show_Slip_Interface_Detail()
        {
            //System.Windows.Forms.DialogResult vdlgResultValue;
            //int mHEADER_INTERFACE_ID = iString.ISNumtoZero(igrNOT_CONFIRM_VAT.GetCellValue("HEADER_INTERFACE_ID"));
            //if (mHEADER_INTERFACE_ID != Convert.ToInt32(0))
            //{
            //    Application.UseWaitCursor = true;
            //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            //    Application.DoEvents();

            //    Form vSLIP_IF_DETAIL = new SLIP_IF_DETAIL(isAppInterfaceAdv1.AppInterface, mHEADER_INTERFACE_ID);
            //    vdlgResultValue = vSLIP_IF_DETAIL.ShowDialog();
            //    vSLIP_IF_DETAIL.Dispose();

            //    Application.DoEvents();
            //    this.Cursor = System.Windows.Forms.Cursors.Default;
            //    Application.UseWaitCursor = false;
            //}
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        #endregion;

        #region ----- XL Export Methods ----

        private void ExportXL()
        {
            int vCountRow = idaVAT_MASTER.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xls";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(idaVAT_MASTER.OraSelectData, vsSaveExcelFileName, vsSheetName);
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

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1()
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                //mMainForm
                //string vPathReport = string.Empty;
                //object vObject = mMainForm.Tag;
                //if (vObject != null)
                //{
                //    vPathReport = mMainForm.Tag;
                //}
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = @"K:\00_5_XL_Print\Ex_XL_Print\XL_Print_02.xls";
                xlPrinting.XLFileOpen();

                int vTerritory = GetTerritory(igrVAT_MASTER.TerritoryLanguage);
                string vPeriodFrom = W_ISSUE_DATE_FR.EditText;
                string vPeriodTo = W_ISSUE_DATE_FR.EditText;
                int vPageNumber = xlPrinting.XLWirte(igrVAT_MASTER, vTerritory, vPeriodFrom, vPeriodTo);

                //xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                ////xlPrinting.Printing(3, 4);


                xlPrinting.Save("t_XL_"); //저장 파일명
                //vMessageText = string.Format("Err : {0}", xlPrinting.ErrorMessage);
                //MessageGrid(vMessageText);

                //xlPrinting.PreView();

                xlPrinting.Dispose();
                //-------------------------------------------------------------------------

                vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
                MessageBoxAdv.Show(vMessageText);
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
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
                    SearchDB(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    //if(idaVAT_MASTER.IsFocused)
                    //{
                    //    idaVAT_MASTER.AddOver();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if(idaVAT_MASTER.IsFocused)
                    //{
                    // idaVAT_MASTER.AddUnder();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaVAT_MASTER.IsFocused)
                    {
                        idaVAT_MASTER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaVAT_MASTER.IsFocused)
                    {
                        idaVAT_MASTER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (idaVAT_MASTER.IsFocused)
                    //{
                    //    idaVAT_MASTER.Delete();
                    //}
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0802_Load(object sender, EventArgs e)
        {
            idaVAT_MASTER.FillSchema();
        }
        
        private void FCMF0802_Shown(object sender, EventArgs e)
        {
            //세금계산서 발행기간.
            DateTime vGetDateTime = GetDate();
            W_PERIOD_YEAR.EditValue = iDate.ISYear(vGetDateTime);

            idcDV_TAX_CODE.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
            idcDV_TAX_CODE.ExecuteNonQuery();
            W_TAX_CODE_NAME.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE_NAME");
            W_TAX_CODE.EditValue = idcDV_TAX_CODE.GetCommandParamValue("O_CODE");
        }

        private void igrVAT_MASTER_CellDoubleClick(object pSender)
        {
            Show_Slip_Detail();
        }

        private void igrNOT_CONFIRM_VAT_CellDoubleClick(object pSender)
        {
            Show_Slip_Interface_Detail();
        }

        private void igrVAT_MASTER_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            decimal mTOTAL_AMOUNT;
            int mGL_COL = igrVAT_MASTER.GetColumnToIndex("GL_AMOUNT");
            int mVAT_COL = igrVAT_MASTER.GetColumnToIndex("VAT_AMOUNT");
            if (e.ColIndex == mGL_COL && iString.ISDecimaltoZero(e.OldValue) != iString.ISDecimaltoZero(e.NewValue))
            {
                mTOTAL_AMOUNT = iString.ISDecimaltoZero(e.NewValue) + iString.ISDecimaltoZero(igrVAT_MASTER.GetCellValue("VAT_AMOUNT"));
                igrVAT_MASTER.SetCellValue("TOTAL_AMOUNT", mTOTAL_AMOUNT);
            }
            else if (e.ColIndex == mVAT_COL && iString.ISDecimaltoZero(e.OldValue) != iString.ISDecimaltoZero(e.NewValue))
            {
                mTOTAL_AMOUNT = iString.ISDecimaltoZero(e.NewValue) + iString.ISDecimaltoZero(igrVAT_MASTER.GetCellValue("GL_AMOUNT"));
                igrVAT_MASTER.SetCellValue("TOTAL_AMOUNT", mTOTAL_AMOUNT);
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaTAX_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "N");
        }

        private void ilaTAX_CODE_0_SelectedRowData(object pSender)
        {
            W_VAT_PERIOD_DESC.EditValue = string.Empty;
            W_VAT_PERIOD_ID.EditValue = string.Empty;
        }

        private void ilaVAT_CLASS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_GUBUN", "N");
        }

        private void ilaCUSTOMER_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "N");
        }
         
        private void ilaVAT_TAX_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_TAX_TYPE", "Y");
        }

        private void ilaVAT_GUBUN_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_GUBUN", "Y");
        }
         
        private void ilaCUSTOMER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUSTOMER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaVAT_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["VAT_ISSUE_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Vat Issue Date(계산서 발행일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VAT_GUBUN"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Vat Class(매입매출구분"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VAT_TAX_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Vat Type(계산서유형"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CUSTOMER_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10135"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["GL_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10243", "&&VALUE:=Supply Amount(공급가액"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VAT_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10243", "&&VALUE:=VAT Amount(세액"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

        
        
    }
}