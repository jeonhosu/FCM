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

namespace FCMF0512
{
    public partial class FCMF0512 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public FCMF0512()
        {
            InitializeComponent();
        }

        public FCMF0512(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        private void Search_DB()
        {
            idaTR_DAILY_SUM_1.Fill();
            idaTR_DAILY_SUM_2.Fill();
            idaTR_DAILY_110.Fill();
            idaTR_DAILY_120.Fill();
            idaTR_DAILY_130_1.Fill();
            idaTR_DAILY_140.Fill();
            idaTR_DAILY_210_1.Fill();
            idaTR_DAILY_210_2.Fill();
            idaTR_DAILY_210_3.Fill();
            idaTR_DAILY_210_4.Fill();
            idaTR_DAILY_SLIP.Fill();
            idaTR_DAILY_FUND_MOVE.Fill();
        }
                
        //private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        //{
        //    ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
        //    ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        //}
        #endregion;

        #region ----- XL Export Methods ----

        private void ExportXL(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //int vCountRow = pGrid.RowCount;
            //if (vCountRow < 1)
            //{
            //    return;
            //}

            //string vsMessage = string.Empty;
            //string vsSheetName = "Slip_Line";

            //saveFileDialog1.Title = "Excel_Save";
            //saveFileDialog1.FileName = "XL_00";
            //saveFileDialog1.DefaultExt = "xls";
            //System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            //saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            //saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
            //if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    //string vsSaveExcelFileName = saveFileDialog1.FileName;
            //    //XL.XLPrint xlExport = new XL.XLPrint();
            //    //bool vXLSaveOK = xlExport.XLExport(pAdapter.OraSelectData, vsSaveExcelFileName, vsSheetName);
            //    //if (vXLSaveOK == true)
            //    //{
            //    //    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
            //    //    MessageBoxAdv.Show(vsMessage);
            //    //}
            //    //else
            //    //{
            //    //    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
            //    //    MessageBoxAdv.Show(vsMessage);
            //    //}
            //    //xlExport.XLClose();
            //}
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

        #endregion;

        #region ----- Week Name Method ----

        private string WeekName(System.DateTime pDate)
        {
            string vWeekName = string.Empty;

            switch (pDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    vWeekName = "월";
                    break;
                case DayOfWeek.Tuesday:
                    vWeekName = "화";
                    break;
                case DayOfWeek.Wednesday:
                    vWeekName = "수";
                    break;
                case DayOfWeek.Thursday:
                    vWeekName = "목";
                    break;
                case DayOfWeek.Friday:
                    vWeekName = "금";
                    break;
                case DayOfWeek.Saturday:
                    vWeekName = "토";
                    break;
                case DayOfWeek.Sunday:
                    vWeekName = "일";
                    break;
            }

            return vWeekName;
        }

        #endregion;

        #region ----- XL Print 0 Methods ----

        //현금 및 예금 현황
        private void XLPrinting00(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexTab)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = pGrid.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                if (vIndexTab == 1)
                {
                    //현금 및 예금 현황
                    XLPrinting02 xlPrinting = new XLPrinting02(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 2)
                {
                    // 정기 예.적금 현황
                    XLPrinting03 xlPrinting = new XLPrinting03(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 3)
                {
                    // 받을 어음 현황
                    XLPrinting04 xlPrinting = new XLPrinting04(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 4)
                {
                    // 지급 어음 // 일반 대출
                    XLPrinting05 xlPrinting = new XLPrinting05(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 5)
                {
                    // 일반 대출
                    XLPrinting06 xlPrinting = new XLPrinting06(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 6)
                {
                    // 사채
                    //XLPrinting07 xlPrinting = new XLPrinting07(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 7)
                {
                    // 한도 대출
                    XLPrinting08 xlPrinting = new XLPrinting08(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 8)
                {
                    // 회전대
                    XLPrinting09 xlPrinting = new XLPrinting09(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 9)
                {
                    // 자금 입/출내역
                    XLPrinting10 xlPrinting = new XLPrinting10(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                else if (vIndexTab == 10)
                {
                    // 이체 입/출내역
                    XLPrinting11 xlPrinting = new XLPrinting11(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
                }
                //-------------------------------------------------------------------------------------

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(pGrid, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 1 Methods ----

        //현금 및 예금 현황
        private void XLPrinting01(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vWeekName = WeekName(GL_DATE_0.DateTimeValue);
            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일[{3}]", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day, vWeekName);

            int vCountRowGrid = pGrid.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting01 xlPrinting = new XLPrinting01(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(pGrid, igrTR_DAILY_SUM_2, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 2 Methods ----

        //현금 및 예금 현황
        private void XLPrinting02()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrTR_DAILY_110.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting02 xlPrinting = new XLPrinting02(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrTR_DAILY_110, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 3 Methods ----

        // 정기 예.적금 현황
        private void XLPrinting03()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrDEPOSIT.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting03 xlPrinting = new XLPrinting03(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrDEPOSIT, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 4 Methods ----

        // 받을 어음 현황
        private void XLPrinting04()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrTR_DAILY_130_1.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting04 xlPrinting = new XLPrinting04(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrTR_DAILY_130_1, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 5 Methods ----

        // 지급 어음
        private void XLPrinting05()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrPAYALBE_BILL.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting05 xlPrinting = new XLPrinting05(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrPAYALBE_BILL, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 6 Methods ----

        // 일반 대출
        private void XLPrinting06()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrLOAN_1.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting06 xlPrinting = new XLPrinting06(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrLOAN_1, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 8 Methods ----

        // 한도 대출
        private void XLPrinting08()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrLOAN_2.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting08 xlPrinting = new XLPrinting08(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrLOAN_2, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 9 Methods ----

        // 회전대
        private void XLPrinting09()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrLOAN_3.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting09 xlPrinting = new XLPrinting09(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrLOAN_3, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 10 Methods ----

        // 자금 입/출내역
        private void XLPrinting10()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrTR_SLIP.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting10 xlPrinting = new XLPrinting10(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrTR_SLIP, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 11 Methods ----

        // 이체 입/출내역
        private void XLPrinting11()
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            string vDate = string.Format("{0}년 {1:D2}월 {2:D2}일", GL_DATE_0.DateTimeValue.Year, GL_DATE_0.DateTimeValue.Month, GL_DATE_0.DateTimeValue.Day);

            int vCountRowGrid = igrFUND_MOVE.RowCount;
            if (vCountRowGrid > 0)
            {
                vMessageText = string.Format("Printing Starting");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                XLPrinting11 xlPrinting = new XLPrinting11(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

                try
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "FCMF0512_001.xls";
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    bool isOpen = xlPrinting.XLFileOpen();
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    if (isOpen == true)
                    {
                        vPageNumber = xlPrinting.LineWrite(igrFUND_MOVE, vDate);

                        ////[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        //xlPrinting.Printing(1, vPageNumber);

                        ////[SAVE]
                        xlPrinting.Save("SLIP_"); //저장 파일명


                        vPageTotal = vPageTotal + vPageNumber;
                    }
                    //-------------------------------------------------------------------------------------

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    xlPrinting.Dispose();
                }
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End [Total Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

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
                    Search_DB();
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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    int vIndexTab = itbTR_DAILY.SelectedIndex;
                    if (vIndexTab == 0)
                    {
                        XLPrinting01(vIndexTab); //자금일보
                    }
                    else if (vIndexTab == 1)
                    {
                        XLPrinting02(vIndexTab); //현금 및 예금 현황
                    }
                    else if (vIndexTab == 2)
                    {
                        XLPrinting03(vIndexTab); // 정기 예.적금 현황
                    }
                    else if (vIndexTab == 3)
                    {
                        XLPrinting04(vIndexTab); // 받을 어음 현황
                    }
                    else if (vIndexTab == 4)
                    {
                        XLPrinting05(vIndexTab); // 지급 어음 // 일반 대출
                    }
                    else if (vIndexTab == 5)
                    {
                        XLPrinting06(vIndexTab); // 일반 대출
                    }
                    else if (vIndexTab == 6)
                    {
                        //XLPrinting07(vIndexTab); // 사채
                    }
                    else if (vIndexTab == 7)
                    {
                        XLPrinting08(vIndexTab); // 한도 대출
                    }
                    else if (vIndexTab == 8)
                    {
                        XLPrinting09(vIndexTab); // 회전대
                    }
                    else if (vIndexTab == 9)
                    {
                        XLPrinting10(vIndexTab); // 자금 입/출내역
                    }
                    else if (vIndexTab == 10)
                    {
                        XLPrinting11(vIndexTab); // 이체 입/출내역
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    int vIndexTab = itbTR_DAILY.SelectedIndex;
                    if (vIndexTab == 0)
                    {
                        ExportXL(igrTR_DAILY_SUM_1);
                    }
                    else if (vIndexTab == 1)
                    {
                        ExportXL(igrTR_DAILY_110);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        
        private void FCMF0512_Load(object sender, EventArgs e)
        {
        
        }

        private void FCMF0512_Shown(object sender, EventArgs e)
        {
            GL_DATE_0.EditValue = DateTime.Today;

            System.DateTime vtmpDate = new DateTime(2010, 12, 10);
            GL_DATE_0.EditValue = vtmpDate;
        }

        private void ibtnTR_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object mMessage;
            idcTR_DAILY_SUM.ExecuteNonQuery();
            mMessage = idcTR_DAILY_SUM.GetCommandParamValue("O_MESSAGE");
            if (iString.ISNull(mMessage) != string.Empty)
            {
                MessageBoxAdv.Show(mMessage.ToString(),  "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion
        
        #region ----- Lookup Event -----
        
        #endregion

        #region ----- Adapeter Event -----
       
        private void idaASSET_CATEGORY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ASSET_CATEGORY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10093"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ASSET_CATEGORY_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10094"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ASSET_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10095"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DPR_METHOD_TYPE"]) != string.Empty && iString.ISNull(e.Row["PROGRESS_YEAR_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["IFRS_DPR_METHOD_TYPE"]) != string.Empty && iString.ISNull(e.Row["IFRS_PROGRESS_YEAR_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10088", "&&VALUE:=Account Name(계정)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["RESIDUAL_VALUE_AMOUNT"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10088", "&&VALUE:=Residual Value Amount(잔존가액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_TO"]) != string.Empty &&
               Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaASSET_CATEGORY_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion
    }
}