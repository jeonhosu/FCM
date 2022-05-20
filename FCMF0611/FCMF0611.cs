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

namespace FCMF0611
{
    public partial class FCMF0611 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0611()
        {
            InitializeComponent();
        }

        public FCMF0611(Form pMainForm, ISAppInterface pAppInterface)
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

            V_APPROVE_STATUS.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
            V_APPROVE_STATUS_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
        }

        private void SearchDB()
        {
            if (iString.ISNull(V_BUDGET_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_BUDGET_YEAR.Focus();
                return;
            }

            idaBUDGET_PLAN_YEAR.Fill();
            idaBUDGET_PLAN_MONTH.Fill();

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
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 0].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "01");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 1].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "02");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 2].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "03");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 3].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "04");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 4].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "05");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 5].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "06");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 6].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "07");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 7].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "08");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 8].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "09");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 9].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "10");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 10].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "11");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 11].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, "12");
                igrPLAN_MONTH.GridAdvExColElement[mStart_Col + 12].HeaderElement[0].Default = string.Format("{0}-{1}", V_BUDGET_YEAR.EditValue, isMessageAdapter1.ReturnText("EAPP_10045"));
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
            for (int r = 0; r < idaBUDGET_PLAN_YEAR.SelectRows.Count; r++)
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
            for (int r = 0; r < idaBUDGET_PLAN_MONTH.SelectRows.Count; r++)
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
            int mStart_Col = 8;
            igrPLAN_YEAR.GridAdvExColElement[mStart_Col + 0].Insertable = 0;
            igrPLAN_YEAR.GridAdvExColElement[mStart_Col + 0].Updatable = 0;
            igrPLAN_YEAR.GridAdvExColElement[mStart_Col + 0].ReadOnly = true;
            if (pDataRow != null)
            {
                if (iString.ISNull(V_ALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString() ||
                    (iString.ISNull(pDataRow["APPROVE_STATUS"]) != "A".ToString() &&
                    iString.ISNull(pDataRow["APPROVE_STATUS"]) != "N".ToString()))
                {
                    if (pDataRow.RowState != DataRowState.Added)
                    {
                        mEnabled_YN = false;
                    }
                }

                if (mEnabled_YN == true)
                {
                    igrPLAN_YEAR.GridAdvExColElement[mStart_Col + 0].Insertable = 1;
                    igrPLAN_YEAR.GridAdvExColElement[mStart_Col + 0].Updatable = 1;
                    igrPLAN_YEAR.GridAdvExColElement[mStart_Col + 0].ReadOnly = false;
                }                
            }
            igrPLAN_MONTH.ResetDraw = true;
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
                if (iString.ISNull(V_ALL_RECORD_FLAG.CheckBoxValue) == "Y".ToString() ||
                    (iString.ISNull(pDataRow["APPROVE_STATUS"]) != "A".ToString() &&
                    iString.ISNull(pDataRow["APPROVE_STATUS"]) != "N".ToString()))
                {
                    if (pDataRow.RowState != DataRowState.Added)
                    {
                        mEnabled_YN = false;
                    }
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
                igrPLAN_MONTH.Focus();
            }
        }

        private void Insert_BUDGET_PLAN_YEAR()
        {
            int mIDX_Col;
            igrPLAN_YEAR.SetCellValue("BUDGET_YEAR", V_BUDGET_YEAR.EditValue);


            mIDX_Col = igrPLAN_YEAR.GetColumnToIndex("DEPT_NAME");
            igrPLAN_YEAR.CurrentCellMoveTo(mIDX_Col);
            igrPLAN_YEAR.CurrentCellActivate(mIDX_Col);
            igrPLAN_YEAR.Focus();
        }

        private void Insert_BUDGET_PLAN_MONTH()
        {
            int mIDX_Col;
            mIDX_Col = igrPLAN_MONTH.GetColumnToIndex("DEPT_NAME");

            igrPLAN_MONTH.CurrentCellMoveTo(mIDX_Col);
            igrPLAN_MONTH.CurrentCellActivate(mIDX_Col);
            igrPLAN_MONTH.Focus();
        }

        private void Create_Plan_Month()
        {
            string mMESSAGE;
            if (iString.ISNull(V_BUDGET_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_BUDGET_YEAR.Focus();
                return;
            }            

            idcBUDGET_PERIOD.ExecuteNonQuery();
            mMESSAGE = iString.ISNull(idcBUDGET_PERIOD.GetCommandParamValue("O_MESSAGE"));
            if (mMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            SearchDB();
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

        #region ----- XL Print 1 Method -----

        private void XLPrinting(string pOutChoice)
        {
            object vPRINT_TYPE = string.Empty;
            DialogResult dlgResult;
            FCMF0611_PRINT vFCMF0611_PRINT = new FCMF0611_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgResult = vFCMF0611_PRINT.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                vPRINT_TYPE = vFCMF0611_PRINT.Get_Print_Type;
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
            vFCMF0611_PRINT.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_1(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            IDA_PRINT_PLAN_DEPT.Fill();
            int vCountRow = IDA_PRINT_PLAN_DEPT.OraSelectData.Rows.Count;
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
                vSaveFileName = "Budget_request_depart";

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
                xlPrinting.OpenFileNameExcel = "FCMF0611_001.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    object vBUDGET_YEAR = V_BUDGET_YEAR.EditValue;

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_1(vBUDGET_YEAR);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(IDA_PRINT_PLAN_DEPT);

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

            IDA_PRINT_PLAN_ACCOUNT.Fill();
            int vCountRow = IDA_PRINT_PLAN_ACCOUNT.OraSelectData.Rows.Count;
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
                vSaveFileName = "Budget_request_account";

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
                xlPrinting.OpenFileNameExcel = "FCMF0611_002.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //헤더 데이터 설정
                    object vBUDGET_YEAR = V_BUDGET_YEAR.EditValue;

                    //헤더 인쇄
                    xlPrinting.HeaderWrite_2(vBUDGET_YEAR);
                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_2(IDA_PRINT_PLAN_ACCOUNT);

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
                    if (idaBUDGET_PLAN_YEAR.IsFocused)
                    {
                        idaBUDGET_PLAN_YEAR.AddOver();
                        Insert_BUDGET_PLAN_YEAR();
                    }                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaBUDGET_PLAN_YEAR.IsFocused)
                    {
                        idaBUDGET_PLAN_YEAR.AddUnder();
                        Insert_BUDGET_PLAN_YEAR();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaBUDGET_PLAN_YEAR.IsFocused)
                    {
                        idaBUDGET_PLAN_YEAR.Update();
                    }
                    else if (idaBUDGET_PLAN_MONTH.IsFocused)
                    {
                        idaBUDGET_PLAN_MONTH.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBUDGET_PLAN_YEAR.IsFocused)
                    {
                        idaBUDGET_PLAN_YEAR.Cancel();
                    }
                    else if ( idaBUDGET_PLAN_MONTH.IsFocused)
                    {
                        idaBUDGET_PLAN_MONTH.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBUDGET_PLAN_YEAR.IsFocused)
                    {
                        idaBUDGET_PLAN_YEAR.Delete();
                    }
                    else if (idaBUDGET_PLAN_MONTH.IsFocused)
                    {
                        idaBUDGET_PLAN_MONTH.Delete();
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

        private void FCMF0611_Load(object sender, EventArgs e)
        {
            idaBUDGET_ACCOUNT.FillSchema();
            idaBUDGET_PLAN_YEAR.FillSchema();
            idaBUDGET_PLAN_MONTH.FillSchema();
        }

        private void FCMF0611_Shown(object sender, EventArgs e)
        {
            V_BUDGET_YEAR.EditValue = DateTime.Today.Year;
            Set_Plan_Month_Header();
        }

        private void ibtnCONFIRM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string mMESSAGE;
            if (iString.ISNull(V_BUDGET_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_BUDGET_YEAR.Focus();
                return;
            }

            idcBUDGET_PLAN_YEAR_CONFIRM.ExecuteNonQuery();
            mMESSAGE = iString.ISNull(idcBUDGET_PLAN_YEAR_CONFIRM.GetCommandParamValue("O_MESSAGE"));
            if (mMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }   
        }

        private void ibtnEXECUTE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaBUDGET_PLAN_YEAR.Update();
        }

        private void ibtREQ_APPROVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            idaBUDGET_ACCOUNT.Update();

            string vSTATUS;
            string vMESSAGE;
            isDataTransaction1.BeginTran();
            idcAPPROVE_REQUEST.ExecuteNonQuery();
            vSTATUS = iString.ISNull(idcAPPROVE_REQUEST.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(idcAPPROVE_REQUEST.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (idcAPPROVE_REQUEST.ExcuteError || vSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            isDataTransaction1.Commit();
            SearchDB();
        }

        private void itbBUDGET_PLAN_Click(object sender, EventArgs e)
        {
            Set_Tab_Focus();
        }

        private void V_EXCEL_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_BUDGET_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_BUDGET_YEAR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_BUDGET_YEAR.Focus();
                return;
            }

            DialogResult vdlgResult = DialogResult.None;
            FCMF0611_UPLOAD vFCMF0611_UPLOAD = new FCMF0611_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface, V_BUDGET_YEAR.EditValue);
            vdlgResult = vFCMF0611_UPLOAD.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            vFCMF0611_UPLOAD.Dispose();
            SearchDB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaACCOUNT_CONTROL_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_FR_0_SelectedRowData(object pSender)
        {
            V_ACCOUNT_DESC_TO.EditValue = V_ACCOUNT_DESC_FR.EditValue;
            V_ACCOUNT_CODE_TO.EditValue = V_ACCOUNT_CODE_FR.EditValue;
            V_ACCOUNT_CONTROL_ID_TO.EditValue = V_ACCOUNT_CONTROL_ID_FR.EditValue;
        }

        private void ilaACCOUNT_CONTROL_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", V_ACCOUNT_CODE_FR.EditValue);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDEPT_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(string.Format("{0}-01", V_BUDGET_YEAR.EditValue)));
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(string.Format("{0}-12", V_BUDGET_YEAR.EditValue)));
        }

        private void ilaDEPT_FR_0_SelectedRowData(object pSender)
        {
            V_DEPT_NAME_TO.EditValue = V_DEPT_NAME_FR.EditValue;
            V_DEPT_CODE_TO.EditValue = V_DEPT_CODE_FR.EditValue;
            V_DEPT_ID_TO.EditValue = V_DEPT_ID_FR.EditValue;
        }

        private void ilaDEPT_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", V_DEPT_CODE_FR.EditValue);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(string.Format("{0}-01", V_BUDGET_YEAR.EditValue)));
            ildDEPT_FR_TO.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(string.Format("{0}-12", V_BUDGET_YEAR.EditValue)));
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDEPT_MONTH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_FR_TO.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT_FR_TO.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT_FR_TO.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BUDGET_CAPACITY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBUDGET_YEAR_PLAN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_YEAR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Budget Year(예산년도)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
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
            //if (iString.ISNull(e.Row["YEAR_AMOUNT"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Amount(예산금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

        private void idaBUDGET_PLAN_MONTH_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(V_BUDGET_YEAR.EditValue) == string.Empty)
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

        private void idaBUDGET_PLAN_YEAR_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Year_Item_Status(pBindingManager.DataRow);
        }

        private void idaBUDGET_PLAN_MONTH_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Set_Grid_Item_Status(pBindingManager.DataRow);
        }

        private void idaBUDGET_PLAN_YEAR_UpdateCompleted(object pSender)
        {
            SearchDB();
        }

        #endregion

    }
}