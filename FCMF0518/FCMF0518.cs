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

namespace FCMF0518
{
    public partial class FCMF0518 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0518()
        {
            InitializeComponent();
        }

        public FCMF0518(Form pMainForm, ISAppInterface pAppInterface)
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
            {//지급어음
                idaBILL_EXPIRY_VENDOR.SetSelectParamValue("W_BILL_CLASS", "1");
                idaBILL_EXPIRY_VENDOR.Fill();

                igrPAYABLE_BILL.Focus();
            }
            else if (itbBILL.SelectedTab.TabIndex == 2)
            {//받을어음.
                idaBILL_EXPIRY_VENDOR.SetSelectParamValue("W_BILL_CLASS", "2");
                idaBILL_EXPIRY_VENDOR.Fill();

                igrRECEIVABLE_BILL.Focus();
            }
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
                    if (idaBILL_EXPIRY_VENDOR.IsFocused)
                    {
                        idaBILL_EXPIRY_VENDOR.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaBILL_EXPIRY_VENDOR.IsFocused)
                    {
                        idaBILL_EXPIRY_VENDOR.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaBILL_EXPIRY_VENDOR.IsFocused)
                    {
                        idaBILL_EXPIRY_VENDOR.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBILL_EXPIRY_VENDOR.IsFocused)
                    {
                        idaBILL_EXPIRY_VENDOR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBILL_EXPIRY_VENDOR.IsFocused)
                    {
                        idaBILL_EXPIRY_VENDOR.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (itbBILL.SelectedTab.TabIndex == 1)
                    {
                        XLPrinting("PRINT", igrRECEIVABLE_BILL);
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
                        XLPrinting("FILE", igrPAYABLE_BILL);
                    }
                    else if (itbBILL.SelectedTab.TabIndex == 2)
                    {
                        XLPrinting("FILE", igrRECEIVABLE_BILL);
                    }
                }

            }
        }

        #endregion;

        #region ----- XL Print Methods ----

        private void XLPrinting(string pOutChoice, ISGridAdvEx pGRID)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;



            object vTerritory = string.Empty;

            int vCountRow = pGRID.RowCount;
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                if (itbBILL.SelectedTab.TabIndex == 1)
                {
                    vSaveFileName = String.Format("업체별받을어음현황(지급어음명세서)");
                }
                else if (itbBILL.SelectedTab.TabIndex == 2)
                {
                    vSaveFileName = String.Format("업체별받을어음현황(받을어음명세서)");
                }

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";

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
            }
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
                xlPrinting.OpenFileNameExcel = "FCMF0518_001.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {

                    // 실제 인쇄
                    vPageNumber = xlPrinting.LineWrite(iString.ISNull(vTerritory), pGRID);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
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

        #region ----- Form Event -----

        private void itbBILL_Click(object sender, EventArgs e)
        {
            Search_DB();
        }

        #endregion

    }
}