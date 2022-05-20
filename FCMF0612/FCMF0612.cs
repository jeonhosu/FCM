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

namespace FCMF0612
{
    public partial class FCMF0612 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mCAP = "N";

        #endregion;

        #region ----- Constructor -----

        public FCMF0612()
        {
            InitializeComponent();
        }

        public FCMF0612(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Set_Default_Value()
        {
            // Budget Select Type.
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "BUDGET_CAPACITY");
            idcDEFAULT_VALUE.ExecuteNonQuery();
        }

        private void SearchDB()
        {
            if (iString.ISNull(BUDGET_YEAR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_YEAR_0.Focus();
                return;
            }

            cbCHECK_YN.CheckBoxValue = "N";
            igrPLAN_YEAR.LastConfirmChanges();
            idaPLAN_YEAR_APPROVE.OraSelectData.AcceptChanges();
            idaPLAN_YEAR_APPROVE.Refillable = true;

            idaPLAN_YEAR_APPROVE.Fill();
            idaPLAN_MONTH_APPROVE.Fill();

            Set_Plan_Month_Header();    //헤더 설정.
            Set_Total_Amount();
            Set_Tab_Focus();
        }

        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Plan_Month_Header()
        {
            int mStart_Col = 7;
            idaMONTH_HEADER.Fill();
            if (idaMONTH_HEADER.SelectRows.Count == 0)
            {
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "01");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "02");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "03");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "04");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "05");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "06");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "07");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "08");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "09");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "10");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "11");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, "12");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 12].HeaderElement[0].Default = string.Format("{0}-{1}", BUDGET_YEAR_0.EditValue, isMessageAdapter1.ReturnText("EAPP_10045"));
            }
            else
            {
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_1"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_2"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_3"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_4"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_5"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_6"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_7"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_8"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_9"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_10"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_11"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["MONTH_12"]);
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 12].HeaderElement[0].Default = iString.ISNull(idaMONTH_HEADER.CurrentRow["YEAR_TOTAL"]);
            }
            igrPLAN_MONTH.ResetDraw = true;
        }

        private void Set_Total_Amount()
        {
            decimal vTotal_Amount = 0;
            object vAmount;
            int vIDXCol;
            // 년예산.
            vIDXCol = igrPLAN_YEAR.GetColumnToIndex("YEAR_AMOUNT");
            if (vIDXCol == -1)
            {
                return;
            }  
            for (int r = 0; r < idaPLAN_YEAR_APPROVE.SelectRows.Count; r++)
            {
                vAmount = 0;
                vAmount = igrPLAN_YEAR.GetCellValue(r, vIDXCol);
                vTotal_Amount = vTotal_Amount + iString.ISDecimaltoZero(vAmount);
            }
            YEAR_TOTAL_AMOUNT.EditValue = vTotal_Amount;

            // 월예산.
            vTotal_Amount = 0;
            vAmount = 0;
            vIDXCol = -1;
            vIDXCol = igrPLAN_MONTH.GetColumnToIndex("YEAR_TOTAL");
            if (vIDXCol == -1)
            {
                return;
            }
            for (int r = 0; r < idaPLAN_MONTH_APPROVE.SelectRows.Count; r++)
            {
                vAmount = 0;
                vAmount = igrPLAN_MONTH.GetCellValue(r, vIDXCol);
                vTotal_Amount = vTotal_Amount + iString.ISDecimaltoZero(vAmount);
            }
            MONTH_TOTAL_AMOUNT.EditValue = vTotal_Amount;
        }

        private void Set_Grid_Year_Item_Status(DataRow pDataRow)
        {
            bool mEnabled_YN = true;
            int mIDX_Col;

            //선택승인.
            cbCHECK_YN.Enabled = false;
            mIDX_Col = igrPLAN_YEAR.GetColumnToIndex("CHECK_YN");
            igrPLAN_YEAR.GridAdvExColElement[mIDX_Col].Insertable = 0;
            igrPLAN_YEAR.GridAdvExColElement[mIDX_Col].Updatable = 0;
            igrPLAN_YEAR.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            if (pDataRow != null)
            {
                if (mCAP == "N")
                {
                    mEnabled_YN = false;
                }
                else if (iString.ISNull(icbALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString())
                {
                    mEnabled_YN = false;
                }

                if (mEnabled_YN == true)
                {
                    //선택승인.
                    cbCHECK_YN.Enabled = true;
                    mIDX_Col = igrPLAN_YEAR.GetColumnToIndex("CHECK_YN");
                    igrPLAN_YEAR.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    igrPLAN_YEAR.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    igrPLAN_YEAR.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
            }
            igrPLAN_YEAR.ResetDraw = true;
        }

        private void Set_Grid_Item_Status(DataRow pDataRow)
        {
            bool mEnabled_YN = true;
            int mStart_Col = 7;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].ReadOnly = true;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].Insertable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].Updatable = 0;
            igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].ReadOnly = true;
            if (pDataRow != null)
            {
                if (mCAP == "N")
                {
                    mEnabled_YN = false;
                }
                else if (iString.ISNull(icbALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString() ||
                    (iString.ISNull(pDataRow["APPROVE_STATUS"]) != "A".ToString() &&
                    iString.ISNull(pDataRow["APPROVE_STATUS"]) != "N".ToString()))
                {
                    mEnabled_YN = false;
                }

                if (iString.ISNull(pDataRow["MONTH_1_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_2_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_3_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_4_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_5_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_6_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_7_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_8_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_9_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_10_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_11_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].ReadOnly = false;
                }
                if (iString.ISNull(pDataRow["MONTH_12_YN"]) == "Y" && mEnabled_YN == true)
                {
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].Insertable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].Updatable = 1;
                    igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].ReadOnly = false;
                }
            }
            igrPLAN_MONTH.ResetDraw = true;
        }

        private void Set_Tab_Focus()
        {
            if (itbBUDGET_PLAN.SelectedTab.TabIndex == 1)
            {
                igrPLAN_YEAR.Focus();
            }
            else if (itbBUDGET_PLAN.SelectedTab.TabIndex == 2)
            {
                igrPLAN_YEAR.LastConfirmChanges();
                idaPLAN_YEAR_APPROVE.OraSelectData.AcceptChanges();
                idaPLAN_YEAR_APPROVE.Refillable = true;

                igrPLAN_MONTH.Focus();
            }
        }

        private void Insert_BUDGET_YEAR_PLAN()
        {
            int mIDX_Col;
            igrPLAN_YEAR.SetCellValue("BUDGET_YEAR", BUDGET_YEAR_0.EditValue);

            mIDX_Col = igrPLAN_YEAR.GetColumnToIndex("BUDGET_YEAR");
            igrPLAN_YEAR.CurrentCellMoveTo(mIDX_Col);
            igrPLAN_YEAR.CurrentCellActivate(mIDX_Col);
            igrPLAN_YEAR.Focus();
        }

        private void Set_CheckBox()
        {
            int mIDX_Col = igrPLAN_YEAR.GetColumnToIndex("CHECK_YN");
            object mCheck_YN = cbCHECK_YN.CheckBoxValue;
            for (int r = 0; r < igrPLAN_YEAR.RowCount; r++)
            {
                igrPLAN_YEAR.SetCellValue(r, mIDX_Col, mCheck_YN);
            }
        }

        private void Get_Cap()
        {
            IDC_BUDGET_MANAGER_CAP.ExecuteNonQuery();
            mCAP = iString.ISNull(IDC_BUDGET_MANAGER_CAP.GetCommandParamValue("O_CAP"));
        }

        #endregion;

        #region ----- XL Print Method -----

        private void XLPrinting(string pOutChoice)
        {
            object vPRINT_TYPE = string.Empty;
            DialogResult dlgResult;
            FCMF0612_PRINT vFCMF0612_PRINT = new FCMF0612_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgResult = vFCMF0612_PRINT.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                vPRINT_TYPE = vFCMF0612_PRINT.Get_Print_Type;
                if (iString.ISNull(vPRINT_TYPE) == "D")
                {
                    //부서별
                    XLPrinting_1(pOutChoice);
                }
                else if (iString.ISNull(vPRINT_TYPE) == "A")
                {
                    //계정별
                    XLPrinting_2(pOutChoice);
                }
            }
            vFCMF0612_PRINT.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_1(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            IDA_PRINT_APPROVE_DEPT.Fill();
            int vCountRow = IDA_PRINT_APPROVE_DEPT.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Budget_assign_depart";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                vMessageText = string.Format(" Writing Starting...");
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0612_001.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    object vBUDGET_YEAR = BUDGET_YEAR_0.EditValue;

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_1(vBUDGET_YEAR);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(IDA_PRINT_APPROVE_DEPT);

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

        private void XLPrinting_2(string pOutChoice)
        {
            //예산신청내역 - 계정별
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            IDA_PRINT_APPROVE_ACCOUNT.Fill();
            int vCountRow = IDA_PRINT_APPROVE_ACCOUNT.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Budget_assign_account";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                vMessageText = string.Format(" Writing Starting...");
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0612_002.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    object vBUDGET_YEAR = BUDGET_YEAR_0.EditValue;

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_2(vBUDGET_YEAR);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_2(IDA_PRINT_APPROVE_ACCOUNT);

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
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaPLAN_MONTH_APPROVE.IsFocused)
                    {
                        idaPLAN_MONTH_APPROVE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if ( idaPLAN_MONTH_APPROVE.IsFocused)
                    {
                        idaPLAN_MONTH_APPROVE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPLAN_MONTH_APPROVE.IsFocused)
                    {
                        idaPLAN_MONTH_APPROVE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0612_Load(object sender, EventArgs e)
        {
            
        }

        private void FCMF0612_Shown(object sender, EventArgs e)
        {
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            APPROVE_STATUS_9.EditValue = "A";

            EMAIL_STATUS.EditValue = "N";
            icbALL_RECORD_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            BUDGET_YEAR_0.EditValue = DateTime.Today.Year;
            Set_Plan_Month_Header();           
            
            Get_Cap();

            idaBUDGET_ACCOUNT.FillSchema();
            idaPLAN_YEAR_APPROVE.FillSchema();
            idaPLAN_MONTH_APPROVE.FillSchema();
        }

        private void irbAPPR_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            APPROVE_STATUS_9.EditValue = iStatus.RadioButtonString;

            //버튼 제어.
            if (mCAP == "N")
            {
                ibtOK.Enabled = false;
                ibtCANCEL.Enabled = false;
            }
            else if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "A")
            {
                ibtOK.Enabled = true;
                ibtCANCEL.Enabled = false;
            }
            else if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "B")
            {
                ibtOK.Enabled = true;
                ibtCANCEL.Enabled = true;
            }
            else if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "C")
            {
                ibtOK.Enabled = false;
                ibtCANCEL.Enabled = true;
            }
            else
            {
                ibtOK.Enabled = false;
                ibtCANCEL.Enabled = false;
            }
            SearchDB();
        }
        
        private void igrPLAN_YEAR_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_CHECK_FLAG = igrPLAN_YEAR.GetColumnToIndex("CHECK_FLAG");
            if (e.ColIndex == vIDX_CHECK_FLAG)
            {
                igrPLAN_YEAR.LastConfirmChanges();
                idaPLAN_YEAR_APPROVE.OraSelectData.AcceptChanges();
                idaPLAN_YEAR_APPROVE.Refillable = true;
            }            
        }
        
        private void ibtOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }

            if (igrPLAN_YEAR.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            isDataTransaction1.BeginTran();
            int vIDX_CHECK_YN = igrPLAN_YEAR.GetColumnToIndex("CHECK_YN");
            int vIDX_DEPT_ID = igrPLAN_YEAR.GetColumnToIndex("DEPT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = igrPLAN_YEAR.GetColumnToIndex("ACCOUNT_CONTROL_ID");

            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < igrPLAN_YEAR.RowCount; i++)
            {
                if (iString.ISNull(igrPLAN_YEAR.GetCellValue(i, vIDX_CHECK_YN), "N") == "Y")
                {
                    igrPLAN_YEAR.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    igrPLAN_YEAR.CurrentCellActivate(i, vIDX_CHECK_YN);

                    idcAPPROVE_STATUS.SetCommandParamValue("W_DEPT_ID", igrPLAN_YEAR.GetCellValue(i, vIDX_DEPT_ID));
                    idcAPPROVE_STATUS.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", igrPLAN_YEAR.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    idcAPPROVE_STATUS.SetCommandParamValue("P_APPROVE_FLAG", "OK");
                    idcAPPROVE_STATUS.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(idcAPPROVE_STATUS.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(idcAPPROVE_STATUS.GetCommandParamValue("O_MESSAGE"));
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (idcAPPROVE_STATUS.ExcuteError || vSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }
            isDataTransaction1.Commit();
            SearchDB();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            ////Email 전송.
            //isDataTransaction1.BeginTran();

            //vSTATUS = "F";
            //vMESSAGE = null;
            //IDC_EMAIL_SEND.SetCommandParamValue("P_MAIL_STATUS", "CREATE");
            //IDC_EMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "DOMESTIC");
            //IDC_EMAIL_SEND.ExecuteNonQuery();
            //vSTATUS = iString.ISNull(IDC_EMAIL_SEND.GetCommandParamValue("O_STATUS"));
            //vMESSAGE = IDC_EMAIL_SEND.GetCommandParamValue("O_MESSAGE");
            //if (IDC_EMAIL_SEND.ExcuteError || vSTATUS == "F")
            //{
            //    isDataTransaction1.RollBack();

            //    //이메일 전송상태 변경 - 오류 때문에 강제 변경.
            //    isDataTransaction1.BeginTran();
            //    IDC_UPDATE_EMAIL_STATUS.SetCommandParamValue("P_MAIL_STATUS", "CREATE");
            //    IDC_UPDATE_EMAIL_STATUS.SetCommandParamValue("P_SOURCE_TYPE", "DOMESTIC");
            //    IDC_UPDATE_EMAIL_STATUS.ExecuteNonQuery();
            //    isDataTransaction1.Commit();

            //    Application.UseWaitCursor = false;
            //    this.Cursor = System.Windows.Forms.Cursors.Default;
            //    Application.DoEvents();
            //    if (iString.ISNull(vMESSAGE) != string.Empty)
            //    {
            //        MessageBoxAdv.Show(iString.ISNull(vMESSAGE), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //    return;
            //}

            //isDataTransaction1.Commit();
            //Application.UseWaitCursor = false;
            //this.Cursor = System.Windows.Forms.Cursors.Default;
            //Application.DoEvents();
            //SEARCH_DB();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_9.EditValue) == "C".ToString())
            {
                EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            isDataTransaction1.BeginTran();
            int vIDX_CHECK_YN = igrPLAN_YEAR.GetColumnToIndex("CHECK_YN");
            int vIDX_DEPT_ID = igrPLAN_YEAR.GetColumnToIndex("DEPT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = igrPLAN_YEAR.GetColumnToIndex("ACCOUNT_CONTROL_ID");

            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < igrPLAN_YEAR.RowCount; i++)
            {
                if (iString.ISNull(igrPLAN_YEAR.GetCellValue(i, vIDX_CHECK_YN), "N") == "Y")
                {
                    igrPLAN_YEAR.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    igrPLAN_YEAR.CurrentCellActivate(i, vIDX_CHECK_YN);

                    idcAPPROVE_STATUS.SetCommandParamValue("W_DEPT_ID", igrPLAN_YEAR.GetCellValue(i, vIDX_DEPT_ID));
                    idcAPPROVE_STATUS.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", igrPLAN_YEAR.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    idcAPPROVE_STATUS.SetCommandParamValue("P_APPROVE_FLAG", "CANCEL");
                    idcAPPROVE_STATUS.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(idcAPPROVE_STATUS.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(idcAPPROVE_STATUS.GetCommandParamValue("O_MESSAGE"));

                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (idcAPPROVE_STATUS.ExcuteError || vSTATUS == "F")
                    {
                        isDataTransaction1.RollBack();
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }
            isDataTransaction1.Commit();
            SearchDB();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }
        
        private void icbCHECK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Set_CheckBox();
        }

        private void itbBUDGET_PLAN_Click(object sender, EventArgs e)
        {
            Set_Tab_Focus();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", ACCOUNT_CODE_FR_0.EditValue);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_FR_0_SelectedRowData(object pSender)
        {
            ACCOUNT_DESC_TO_0.EditValue = ACCOUNT_DESC_FR_0.EditValue;
            ACCOUNT_CODE_TO_0.EditValue = ACCOUNT_CODE_FR_0.EditValue;
            ACCOUNT_CONTROL_ID_TO_0.EditValue = ACCOUNT_CONTROL_ID_FR_0.EditValue;
        }

        private void ilaDEPT_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(string.Format("{0}-01", BUDGET_YEAR_0.EditValue)));
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(string.Format("{0}-12", BUDGET_YEAR_0.EditValue)));
        }

        private void ilaDEPT_FR_0_SelectedRowData(object pSender)
        {
            DEPT_NAME_TO_0.EditValue = DEPT_NAME_FR_0.EditValue;
            DEPT_CODE_TO_0.EditValue = DEPT_CODE_FR_0.EditValue;
            DEPT_ID_TO_0.EditValue = DEPT_ID_FR_0.EditValue;
        }

        private void ilaDEPT_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", DEPT_CODE_FR_0.EditValue);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(string.Format("{0}-01", BUDGET_YEAR_0.EditValue)));
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(string.Format("{0}-12", BUDGET_YEAR_0.EditValue)));
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaPLAN_YEAR_APPROVE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Year_Item_Status(pBindingManager.DataRow);
        }

        private void idaPLAN_MONTH_APPROVE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(BUDGET_YEAR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Department(예산부서)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Account Code(예산 계정)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPLAN_MONTH_APPROVE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Item_Status(pBindingManager.DataRow);
        }

        private void idaBUDGET_ACCOUNT_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            idaPLAN_YEAR_APPROVE.Fill();
            idaPLAN_MONTH_APPROVE.Fill();
        }

        private void idaPLAN_MONTH_APPROVE_UpdateCompleted(object pSender)
        {
            SearchDB();
        }
        
        #endregion

    }
}