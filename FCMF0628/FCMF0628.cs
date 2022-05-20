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
using System.Text;
using System.Net;
using System.IO;

namespace FCMF0628
{
    public partial class FCMF0628 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        bool mSUB_SHOW_FLAG = false;
        object mSESSION_ID;

        #endregion;

        #region ----- Constructor -----

        public FCMF0628()
        {
            InitializeComponent();
        }

        public FCMF0628(Form pMainForm, ISAppInterface pAppInterface)
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

            //APPROVE_STATUS_0.EditValue =idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
            //APPROVE_STATUS_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
        }

        private void SearchDB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_DETAIL.TabIndex)
            {
                SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
            }
            else
            {
                if (iString.ISNull(W_BUDGET_PERIOD_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_BUDGET_PERIOD_FR.Focus();
                    return;
                }
                IDA_BUDGET_APPLY_HEADER.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", -1);
                IDA_BUDGET_APPLY_HEADER.Fill();

                IDA_BUDGET_APPLY_LIST.Fill();
                Set_Total_Amount();
                IGR_BUDGET_APPLY_LINE.Focus();
            }
        }

        private void SearchDB_DTL(object pBUDGET_ADD_HEADER_ID)
        {
            if (iString.ISNull(pBUDGET_ADD_HEADER_ID) != string.Empty)
            {
                TB_MAIN.SelectedIndex = 1;
                TB_MAIN.SelectedTab.Focus();

                IDA_BUDGET_APPLY_HEADER.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", pBUDGET_ADD_HEADER_ID); 
                try
                {
                    IDA_BUDGET_APPLY_HEADER.Fill();
                }
                catch (Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }

                BUDGET_PERIOD.Focus();    
            }
        }

        private void Budget_Add_Header_Insert()
        {
            IGR_BUDGET_APPLY_LINE.SetCellValue("BUDGET_PERIOD", W_BUDGET_PERIOD_TO.EditValue);

            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus(); 

            BUDGET_PERIOD.Focus();
        }

        private void Budget_Add_Line_Insert()
        {
            IGR_BUDGET_APPLY_LINE.Focus();
        }

        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ILD_COMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter_W(object pGroupCode, object pWhere, object pEnabled_YN)
        {
            ILD_COMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ILD_COMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ILD_COMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Total_Amount()
        {
            decimal vBUDGET_AMOUNT = 0;
            decimal vAPPLY_SUM_AMOUNT = 0;
            decimal vAPPLY_AMOUNT = 0;
            decimal vREMAIN_AMOUNT = 0;
            object vAmount;
            int vIDX_BUDGET_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("BUDGET_AMOUNT");
            int vIDX_APPLY_SUM_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPLY_SUM_AMOUNT");
            int vIDX_APPLY_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPLY_AMOUNT");
            int vIDX_REMAIN_AMOUNT = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("REMAIN_AMOUNT");
            for (int r = 0; r < IGR_BUDGET_APPLY_LINE.RowCount; r++)
            {
                //기초예산액.
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_BUDGET_AMOUNT);
                vBUDGET_AMOUNT = vBUDGET_AMOUNT + iString.ISDecimaltoZero(vAmount);

                //누적사용액
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_APPLY_SUM_AMOUNT);
                vAPPLY_SUM_AMOUNT = vAPPLY_SUM_AMOUNT + iString.ISDecimaltoZero(vAmount);

                //신청예산액
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_APPLY_AMOUNT);
                vAPPLY_AMOUNT = vAPPLY_AMOUNT + iString.ISDecimaltoZero(vAmount);

                //신청예산액
                vAmount = 0;
                vAmount = IGR_BUDGET_APPLY_LINE.GetCellValue(r, vIDX_REMAIN_AMOUNT);
                vREMAIN_AMOUNT = vREMAIN_AMOUNT + iString.ISDecimaltoZero(vAmount);
            }
            V_BUDGET_AMOUNT.EditValue = vBUDGET_AMOUNT;
            V_APPLY_SUM_AMOUNT.EditValue = vAPPLY_SUM_AMOUNT;
            V_APPLY_AMOUNT.EditValue = vAPPLY_AMOUNT;
            V_REMAIN_AMOUNT.EditValue = vREMAIN_AMOUNT;
        }

        private void EXE_BUDGET_ADD_STATUS(object pPERIOD_NAME, object pAPPROVE_STATUS, object pAPPROVE_FLAG)
        {
            IDA_BUDGET_APPLY_LIST.Update(); //수정사항 반영.

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_CHECK_YN = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("CHECK_YN");
            int vIDX_BUDGET_TYPE = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("BUDGET_TYPE");
            int vIDX_BUDGET_PERIOD = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("BUDGET_PERIOD");
            int vIDX_DEPT_ID = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("DEPT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_BUDGET_APPLY_LINE.RowCount; i++)
            {
                if (iString.ISNull(IGR_BUDGET_APPLY_LINE.GetCellValue(i, vIDX_CHECK_YN), "N") == "Y")
                {
                    IGR_BUDGET_APPLY_LINE.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    IGR_BUDGET_APPLY_LINE.CurrentCellActivate(i, vIDX_CHECK_YN);

                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_APPLY_LINE.GetCellValue(i, vIDX_BUDGET_TYPE));
                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_APPLY_LINE.GetCellValue(i, vIDX_BUDGET_PERIOD));
                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_APPLY_LINE.GetCellValue(i, vIDX_DEPT_ID));
                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_APPLY_LINE.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("P_APPROVE_STATUS", pAPPROVE_STATUS);
                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("P_APPROVE_FLAG", pAPPROVE_FLAG);
                    IDA_BUDGET_APPLY_STATUS.SetCommandParamValue("P_CHECK_YN", IGR_BUDGET_APPLY_LINE.GetCellValue(i, vIDX_CHECK_YN));
                    IDA_BUDGET_APPLY_STATUS.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(IDA_BUDGET_APPLY_STATUS.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(IDA_BUDGET_APPLY_STATUS.GetCommandParamValue("O_MESSAGE"));
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (IDA_BUDGET_APPLY_STATUS.ExcuteError || vSTATUS == "F")
                    {
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
            SearchDB();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void Set_Grid_Item_Status(DataRow pDataRow)
        {
            bool mEnabled_YN = true;
            int vIDX_CHECK = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("CHECK_YN");
            int mIDX_Col;

            //// 신청금액.
            //mIDX_Col = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("AMOUNT");
            //IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            //IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            //IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

            // 신청사유.
            mIDX_Col = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("CAUSE_NAME");
            IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 비고.
            mIDX_Col = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("REMARK");
            IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            if (pDataRow != null)
            {
                if ((iString.ISNull(pDataRow["APPROVE_STATUS"]) != "A".ToString() &&
                    iString.ISNull(pDataRow["APPROVE_STATUS"]) != "N".ToString()))
                {
                    if (pDataRow.RowState != DataRowState.Added)
                    {
                        mEnabled_YN = false;
                    }
                }

                if (iString.ISNull(W_APPROVE_STATUS.EditValue) == string.Empty)
                {
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[vIDX_CHECK].Insertable = 0;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[vIDX_CHECK].Updatable = 0;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[vIDX_CHECK].ReadOnly = true;
                }
                else
                {
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[vIDX_CHECK].Insertable = 1;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[vIDX_CHECK].Updatable = 1;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[vIDX_CHECK].ReadOnly = false;
                }

                if (mEnabled_YN == true)
                {
                    //// 신청금액.
                    //mIDX_Col = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("AMOUNT");
                    //IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    //IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    //IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 신청사유.
                    mIDX_Col = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("CAUSE_NAME");
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 비고.
                    mIDX_Col = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("REMARK");
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    IGR_BUDGET_APPLY_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
            }
            IGR_BUDGET_APPLY_LINE.ResetDraw = true;
        }

        private void Select_Check_YN(ISGridAdvEx pGrid, object pCHECK_FLAG)
        {
            int vIDX_CHECK = pGrid.GetColumnToIndex("CHECK_YN");
            int vIDX_APPROVE_STATUS = pGrid.GetColumnToIndex("APPROVE_STATUS");
            object vAPPROVE_STATUS = string.Empty;
            for (int i = 0; i < pGrid.RowCount; i++)
            {
                vAPPROVE_STATUS = pGrid.GetCellValue(i, vIDX_APPROVE_STATUS);
                if (iString.ISNull(W_APPROVE_STATUS.EditValue) != string.Empty)
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, pCHECK_FLAG);
                }
                else
                {
                    pGrid.SetCellValue(i, vIDX_CHECK, "N");
                }
            }

            pGrid.LastConfirmChanges();
            IDA_BUDGET_APPLY_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_LIST.Refillable = true;
        }

        private bool Check_Added()
        {
            Boolean Row_Added_Status = false;
            
            //헤더 체크 
            for (int r = 0; r < IDA_BUDGET_APPLY_HEADER.SelectRows.Count; r++)
            {
                if (IDA_BUDGET_APPLY_HEADER.SelectRows[r].RowState == DataRowState.Added ||
                    IDA_BUDGET_APPLY_HEADER.SelectRows[r].RowState == DataRowState.Modified)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10169"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            //헤더 변경없으면 라인 체크 
            if (Row_Added_Status == false)
            {
                for (int r = 0; r < IDA_BUDGET_APPLY_LINE.SelectRows.Count; r++)
                {
                    if (IDA_BUDGET_APPLY_LINE.SelectRows[r].RowState == DataRowState.Added ||
                        IDA_BUDGET_APPLY_LINE.SelectRows[r].RowState == DataRowState.Modified)
                    {
                        Row_Added_Status = true;
                    }
                }
                if (Row_Added_Status == true)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10169"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            return (Row_Added_Status);
        }

        private void Init_Sub_Panel(bool pShow_Flag, string pSub_Panel)
        {
            if (mSUB_SHOW_FLAG == true && pShow_Flag == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pShow_Flag == true)
            {
                try
                {
                    if (pSub_Panel == "COPY_BUDGET")
                    {
                        //GB_COPY_DOCUMENT.Left = 180;
                        //GB_COPY_DOCUMENT.Top = 95;

                        //GB_COPY_DOCUMENT.Width = 550;
                        //GB_COPY_DOCUMENT.Height = 195;

                        //GB_COPY_DOCUMENT.Border3DStyle = Border3DStyle.Bump;
                        //GB_COPY_DOCUMENT.BorderStyle = BorderStyle.Fixed3D;

                        //GB_COPY_DOCUMENT.Visible = true;
                    }

                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                TB_MAIN.Enabled = false;                 
            }
            else
            {
                try
                {
                    //if (pSub_Panel == "ALL")
                    //{
                    //    GB_COPY_DOCUMENT.Visible = false;
                    //}
                    //else if (pSub_Panel == "COPY_BUDGET")
                    //{
                    //    GB_COPY_DOCUMENT.Visible = false;
                    //}

                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }

                TB_MAIN.Enabled = true;        
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        //계정잔액명세서 SHOW
        private void Get_Temp_Slip()
        {
            IDA_BUDGET_APPLY_HEADER.Update();

            //delete temp data : 계정잔액 대상 산출과 전표 저장 완료시 변경//
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }
             
            DialogResult vRESULT;
            FCMF0628_SLIP vFCMF0628_SLIP = new FCMF0628_SLIP(isAppInterfaceAdv1.AppInterface, mSESSION_ID,  
                                                             BUDGET_DEPT_ID.EditValue, BUDGET_DEPT_CODE.EditValue, BUDGET_DEPT_NAME.EditValue,  
                                                             BUDGET_PERIOD.EditValue, 
                                                             BUDGET_APPLY_HEADER_ID.EditValue, 
                                                             IGR_BUDGET_APPLY_SLIP);
            vRESULT = vFCMF0628_SLIP.ShowDialog();
            if (vRESULT != DialogResult.OK)
            {
                return;
                //Set_Insert_Slip(BUDGET_DEPT_ID.EditValue, vACCOUNT_CONTROL_ID); 
            }
            vFCMF0628_SLIP.Dispose();
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        //계정잔액명세서 SHOW

        private void Delete_Temp_Slip()
        {
            IDA_BUDGET_APPLY_HEADER.Update();

            //delete temp data : 계정잔액 대상 산출과 전표 저장 완료시 변경//
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            DialogResult vRESULT;
            FCMF0628_SLIP_DEL vFCMF0628_SLIP_DEL = new FCMF0628_SLIP_DEL(isAppInterfaceAdv1.AppInterface, mSESSION_ID,
                                                                         BUDGET_DEPT_ID.EditValue, BUDGET_DEPT_CODE.EditValue, BUDGET_DEPT_NAME.EditValue,
                                                                         BUDGET_PERIOD.EditValue,
                                                                         BUDGET_APPLY_HEADER_ID.EditValue, 
                                                                         IGR_BUDGET_APPLY_SLIP);
            vRESULT = vFCMF0628_SLIP_DEL.ShowDialog();
            if (vRESULT != DialogResult.OK)
            {
                return;
                //Set_Insert_Slip(BUDGET_DEPT_ID.EditValue, vACCOUNT_CONTROL_ID); 
            }
            vFCMF0628_SLIP_DEL.Dispose();
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }  
        
        private void Set_Insert_Slip(object pBUDGET_DEPT_ID, object pACCOUNT_CONTROL_ID)
        {
            IDA_BUDGET_APPLY_SLIP_GT.SetSelectParamValue("P_SESSION_ID", mSESSION_ID);
            IDA_BUDGET_APPLY_SLIP_GT.SetSelectParamValue("P_BUDGET_PERIOD", BUDGET_PERIOD.EditValue);
            IDA_BUDGET_APPLY_SLIP_GT.SetSelectParamValue("P_BUDGET_DEPT_ID", pBUDGET_DEPT_ID);
            IDA_BUDGET_APPLY_SLIP_GT.SetSelectParamValue("P_ACCOUNT_CONTROL_ID", pACCOUNT_CONTROL_ID);
            IDA_BUDGET_APPLY_SLIP_GT.Fill();
            if (IDA_BUDGET_APPLY_SLIP_GT.SelectRows.Count < 1)
            {
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Not found data, Check data");
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int Row_Count = IGR_BUDGET_APPLY_SLIP.RowCount; 

            IGR_BUDGET_APPLY_SLIP.BeginUpdate();
            IDA_BUDGET_APPLY_SLIP.MoveLast(IGR_BUDGET_APPLY_SLIP.Name);
            try
            {
                for (int i = 0; i < IDA_BUDGET_APPLY_SLIP_GT.CurrentRows.Count; i++)
                {
                    IDA_BUDGET_APPLY_SLIP.AddUnder();
                    for (int c = 0; c < IGR_BUDGET_APPLY_SLIP.GridAdvExColElement.Count; c++)
                    {
                        if (IGR_BUDGET_APPLY_SLIP.GridAdvExColElement[c].DataColumn.ToString() != "BUDGET_APPLY_LINE_ID")
                        {
                            IGR_BUDGET_APPLY_SLIP.SetCellValue(i + Row_Count, c, IDA_BUDGET_APPLY_SLIP_GT.OraDataSet().Rows[i][c]);
                        }
                    } 
                }
            }
            catch (Exception ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                IGR_BUDGET_APPLY_SLIP.EndUpdate();
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            IGR_BUDGET_APPLY_SLIP.EndUpdate();
 
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();            
        }
         
        private void Show_Detail(object pPERIOD_NAME, object pSLIP_NUM  
                                , object pBUDGET_DEPT_NAME, object pBUDGET_DEPT_ID)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            FCMF0628_DETAIL vFCMF0628_DETAIL = new FCMF0628_DETAIL(isAppInterfaceAdv1.AppInterface, pPERIOD_NAME, pSLIP_NUM  
                                                                , pBUDGET_DEPT_NAME, pBUDGET_DEPT_ID); 

            dlgRESULT = vFCMF0628_DETAIL.ShowDialog();
            vFCMF0628_DETAIL.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void Sync_Button(object pPERIOD_NAME)
        {
            if (iDate.ISMonth_1st(pPERIOD_NAME) < iDate.ISMonth_1st("2017-07"))
            {
                BTN_CANCEL_REQ_APPROVAL.Visible = true;
            }
            else
            {
                BTN_CANCEL_REQ_APPROVAL.Visible = false;
            } 
        }

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


        #region ----- XL Print Methods ----

        private void XLPrinting_Main(string pOutput_Type)
        {
            object vBUDGET_PERIOD;
            if(TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                vBUDGET_PERIOD = IGR_BUDGET_APPLY_LIST.GetCellValue("BUDGET_PERIOD");
            }
            else
            {
                vBUDGET_PERIOD = BUDGET_PERIOD.EditValue;
            }
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(vBUDGET_PERIOD));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "FCMF0628");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            XLPrinting(pOutput_Type);             
        }

        private void XLPrinting(string pOutput_Type)
        {
            //string vDefaultPrinter = GetDefaultPrinter();
            //PD.PrinterSettings = PS;
            //if (PD.ShowDialog() == DialogResult.OK)
            //{
            //    SetDefaultPrinter(PD.PrinterSettings.PrinterName);
            //} 

            string vMessageText = string.Empty;
            int vPageNumber = 0;

            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            object vBUDGET_APPLY_HEADER_ID = BUDGET_APPLY_HEADER_ID.EditValue;
            if (TB_MAIN.SelectedTab.TabIndex == TP_LIST.TabIndex)
            {
                vBUDGET_APPLY_HEADER_ID = IGR_BUDGET_APPLY_LIST.GetCellValue("BUDGET_APPLY_HEADER_ID");
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1); 

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0628_001.xlsx"; 
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    //회계법인명.
                    IDC_GET_COMPANY_NAME_P.ExecuteNonQuery();
                    object vSOB_DESC = IDC_GET_COMPANY_NAME_P.GetCommandParamValue("O_SOB_DESC");

                    IDA_PRINT_BUDGET_APPLY_HEADER.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", vBUDGET_APPLY_HEADER_ID);
                    IDA_PRINT_BUDGET_APPLY_HEADER.Fill();

                    IDA_PRINT_BUDGET_APPLY_LINE.SetSelectParamValue("P_BUDGET_APPLY_HEADER_ID", IDA_PRINT_BUDGET_APPLY_HEADER.CurrentRow["BUDGET_APPLY_HEADER_ID"]);
                    IDA_PRINT_BUDGET_APPLY_LINE.Fill();

                    IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_TYPE", IDA_PRINT_BUDGET_APPLY_HEADER.CurrentRow["BUDGET_TYPE"]);
                    IDA_PRINT_APPROVAL_STEP_PERSON.SetSelectParamValue("P_BUDGET_HEADER_ID", IDA_PRINT_BUDGET_APPLY_HEADER.CurrentRow["BUDGET_APPLY_HEADER_ID"]);
                    IDA_PRINT_APPROVAL_STEP_PERSON.Fill();

                    vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_BUDGET_APPLY_HEADER, IDA_PRINT_BUDGET_APPLY_LINE, IDA_PRINT_APPROVAL_STEP_PERSON, vSOB_DESC, vLOCAL_DATE);

                    if (pOutput_Type == "PRINT")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.Printing(1, vPageNumber);

                    }
                    else if (pOutput_Type == "FILE")
                    {
                        ////[SAVE]
                        xlPrinting.SAVE("Budget_Req_"); //저장 파일명
                    }
                     
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }

                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                //SetDefaultPrinter(vDefaultPrinter);

                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            //SetDefaultPrinter(vDefaultPrinter);
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
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
                    if (IDA_BUDGET_APPLY_HEADER.IsFocused)
                    { 
                        IDA_BUDGET_APPLY_HEADER.AddOver();
                        Budget_Add_Header_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BUDGET_APPLY_HEADER.IsFocused)
                    { 
                        IDA_BUDGET_APPLY_HEADER.AddUnder();
                        Budget_Add_Header_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_APPLY_HEADER.Update();
                    }
                    catch
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_APPLY_LINE.IsFocused)
                    {
                        IDA_BUDGET_APPLY_LINE.Cancel();
                    }
                    else if (IDA_BUDGET_APPLY_SLIP.IsFocused)
                    {
                        IDA_BUDGET_APPLY_SLIP.Cancel();
                    }
                    else
                    {
                        IDA_BUDGET_APPLY_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_APPLY_LINE.IsFocused)
                    {
                        IDA_BUDGET_APPLY_LINE.Delete();
                    }
                    else if (IDA_BUDGET_APPLY_SLIP.IsFocused)
                    {
                        IDA_BUDGET_APPLY_SLIP.Delete();
                    }
                    else
                    {
                        IDA_BUDGET_APPLY_HEADER.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0628_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_APPLY_LIST.FillSchema();
            IDA_BUDGET_APPLY_HEADER.FillSchema();
            IDA_BUDGET_APPLY_LINE.FillSchema();

            //서브판넬 
            Init_Sub_Panel(false, "ALL");
        }

        private void FCMF0628_Shown(object sender, EventArgs e)
        {
            W_BUDGET_PERIOD_FR.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_BUDGET_PERIOD_TO.EditValue = iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 1)); 
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;

            BTN_CHG_APPROVAL_STEP.BringToFront();
            BTN_SELECT_TEMP_SLIP.BringToFront();
            BTN_DELETE_SLIP.BringToFront();

            IDC_GET_SESSION_ID_P.ExecuteNonQuery();
            mSESSION_ID = IDC_GET_SESSION_ID_P.GetCommandParamValue("O_SESSION_ID");

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }
         
        private void irbALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRADIO = sender as ISRadioButtonAdv;
            W_APPROVE_STATUS.EditValue = vRADIO.RadioButtonValue;

            //버튼제어 및 체크박스 제어.
            if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "N")
            {
                BTN_REQ_APPROVAL.Enabled = true;
                BTN_CHG_APPROVAL_STEP.Enabled = true;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
                BTN_SELECT_TEMP_SLIP.Enabled = true;
            } 
            else if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "Y")
            {
                BTN_REQ_APPROVAL.Enabled = false;
                BTN_CHG_APPROVAL_STEP.Enabled = false;
                BTN_CANCEL_REQ_APPROVAL.Enabled = true;
                BTN_SELECT_TEMP_SLIP.Enabled = false;
            }
            else
            {
                BTN_REQ_APPROVAL.Enabled = false;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
                BTN_CHG_APPROVAL_STEP.Enabled = false;
                BTN_SELECT_TEMP_SLIP.Enabled = false;
            }
            SearchDB();
        }
         
        private void IGR_BUDGET_ADD_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_BUDGET_APPLY_LIST.RowCount > 0)
            {
                SearchDB_DTL(IGR_BUDGET_APPLY_LIST.GetCellValue("BUDGET_APPLY_HEADER_ID"));
            }
        }

        private void ibtREQ_APPROVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_BUDGET_APPLY_LIST.Update();

            object mValue;
            int mRowCount = IGR_BUDGET_APPLY_LINE.RowCount;
            int mIDX_COL = IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPROVE_STATUS");

            for (int R = 0; R < mRowCount; R++)
            {
                if (iString.ISNull(IGR_BUDGET_APPLY_LINE.GetCellValue(R, mIDX_COL)) == "N".ToString())
                {// 승인미요청 건에 대해서 승인 처리.
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_APPLY_LINE.GetCellValue(R, IGR_BUDGET_APPLY_LINE.GetColumnToIndex("BUDGET_TYPE")));
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_APPLY_LINE.GetCellValue(R, IGR_BUDGET_APPLY_LINE.GetColumnToIndex("BUDGET_PERIOD")));
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_APPLY_LINE.GetCellValue(R, IGR_BUDGET_APPLY_LINE.GetColumnToIndex("DEPT_ID")));
                    IDC_APPROVE_REQUEST.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_APPLY_LINE.GetCellValue(R, IGR_BUDGET_APPLY_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID")));
                    IDC_APPROVE_REQUEST.ExecuteNonQuery();

                    mValue = DBNull.Value;
                    mValue = IDC_APPROVE_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                    IGR_BUDGET_APPLY_LINE.SetCellValue(R, IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPROVE_STATUS"), mValue);

                    mValue = DBNull.Value;
                    mValue = IDC_APPROVE_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                    IGR_BUDGET_APPLY_LINE.SetCellValue(R, IGR_BUDGET_APPLY_LINE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                }
            }
            IDA_BUDGET_APPLY_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_APPLY_LIST.Refillable = true;
        }

        private bool SendAgit(string strMessage)
        {
            try
            {
                //HttpWebResponse wRes;

                //Uri uri = new Uri("http://dev.seilpcb.co.kr/Linkage/Approval_SEIL1.aspx?1705-057&param2=2013040102"); // URL 설정
                //HttpWebRequest wReq = (HttpWebRequest)WebRequest.Create(uri); // HttpWebRequest 생성
                //wReq.Method = "POST"; // 전송 방식 "GET" 과 "POST" 중 POST 방식으로 보내야 하기 때문에 POST로 설정
 
                //byte[] bArray = Encoding.UTF8.GetBytes(strMessage);
  
                //Stream dtStream = wReq.GetRequestStream();
                //dtStream .Write(bArray , 0, bArray .Length);
                //dtStream .Close();
 
                //using (wRes = (HttpWebResponse)wReq.GetResponse())
                //{
                //    Stream respPostStream = wRes.GetResponseStream();
                //    StreamReader readerPost = new StreamReader(respPostStream, Encoding.GetEncoding("EUC-KR"), true);
 
                //    String resResult = readerPost.ReadToEnd();
                //}

                string vGW_URL = string.Format("http://gw.seilpcb.co.kr/Linkage/Approval_SEIL1.aspx?param1={0}&param2={1}", BUDGET_APPLY_NUM.EditValue, PERSON_NUM.EditValue);
                System.Diagnostics.Process.Start(vGW_URL);

                //http://dev.seilpcb.co.kr/Linkage/Approval_SEIL1.aspx?param1=BUDGET_APPLY_NUM키&param2=사번
                //http://dev.seilpcb.co.kr/Linkage/Approval_SEIL1.aspx?param1=1705-057&param2=2013040102
                //G/W URL 호출
                //string vGW_URL = "http://dev.seilpcb.co.kr/Linkage/Approval_SEIL1.aspx";

                //StringBuilder vPostPara = new StringBuilder();
                //vPostPara.Append("param1=" + BUDGET_APPLY_NUM.EditValue);   
                //vPostPara.Append("&param2=" + PERSON_NUM.EditValue);

                //Encoding vEncoding = Encoding.UTF8;
                //byte[] vResult = vEncoding.GetBytes(vPostPara.ToString());

                //HttpWebRequest vRequest = (HttpWebRequest)WebRequest.Create(vGW_URL);
                //vRequest.Method = "POST";
                //vRequest.ContentType = "application/x-www-form-urlencoded";
                //vRequest.ContentLength = vResult.Length;

                //System.IO.Stream vPostSendStream = vRequest.GetRequestStream();
                //vPostSendStream.Write(vResult, 0, vResult.Length);
                //vPostSendStream.Close();

                //vRequest.AllowAutoRedirect = true;
                
                //HttpWebResponse vResponse = (HttpWebResponse)vRequest.GetResponse();

                //System.IO.Stream vResponseStream = vResponse.GetResponseStream();
                //System.IO.StreamReader vReaderPost = new System.IO.StreamReader(vResponseStream, Encoding.Default);
                //string vResultPost = vReaderPost.ReadToEnd();
                //if (vResponse.StatusCode == HttpStatusCode.OK)
                //{

                //}
                //else
                //{
                //    return false;
                //} 
            }
            catch (WebException Ex)
            {
                //예외처리는 특별히 하지 않았음. 귀찮아서.
                 MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        } 

        private void BTN_REQ_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue,0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            if (iDate.ISMonth_1st(BUDGET_PERIOD.EditValue) >= iDate.ISMonth_1st("2017-07"))
            {
                if (SendAgit("") == false)
                {
                    return;
                }
            }

            IDC_EXEC_BUDGET_APPLY_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_EXEC_BUDGET_APPLY_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_EXEC_BUDGET_APPLY_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            
            if (IDC_EXEC_BUDGET_APPLY_REQ.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            } 
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        private void BTN_CANCEL_REQ_APPROVAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_CANCEL_BUDGET_APPLY_REQ.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CANCEL_BUDGET_APPLY_REQ.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CANCEL_BUDGET_APPLY_REQ.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_BUDGET_APPLY_REQ.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            } 
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }

        private void BTN_CHG_APPROVAL_STEP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(BUDGET_PERIOD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_PERIOD.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISDecimaltoZero(APPROVAL_STEP_SEQ.EditValue, 0) == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(APPROVAL_STEP_SEQ))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_TYPE_NAME.Focus();
                return;
            }
            if (iString.ISNull(BUDGET_APPLY_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_APPLY_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BUDGET_APPLY_NUM.Focus();
                return;
            }

            DialogResult dlgResult = DialogResult.None;
            FCMF0628_APPR_STEP vFCMF0628_APPR_STEP = new FCMF0628_APPR_STEP(isAppInterfaceAdv1.AppInterface,
                                                                            BUDGET_APPLY_NUM.EditValue, APPROVAL_STEP_SEQ.EditValue,
                                                                            BUDGET_PERIOD.EditValue, BUDGET_APPLY_HEADER_ID.EditValue,
                                                                            BUDGET_TYPE_NAME.EditValue, BUDGET_TYPE.EditValue, 
                                                                            BUDGET_DEPT_NAME.EditValue, BUDGET_DEPT_CODE.EditValue, BUDGET_DEPT_ID.EditValue);
            dlgResult = vFCMF0628_APPR_STEP.ShowDialog();
            Application.DoEvents();

            vFCMF0628_APPR_STEP.Dispose();
            SearchDB_DTL(BUDGET_APPLY_HEADER_ID.EditValue);
        }
        
        private void BTN_SELECT_TEMP_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Get_Temp_Slip();
        }

        private void BTN_DELETE_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Delete_Temp_Slip();
        }

        private void IGR_BUDGET_APPLY_SLIP_CellDoubleClick(object pSender)
        {
            if (IGR_BUDGET_APPLY_SLIP.Row > 0)
            {
                Show_Detail(BUDGET_PERIOD.EditValue, IGR_BUDGET_APPLY_SLIP.GetCellValue("SLIP_NUM")
                            , BUDGET_DEPT_NAME.EditValue, BUDGET_DEPT_ID.EditValue);
            }
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ILA_PERIOD_FR_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_PERIOD_TO_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", W_BUDGET_PERIOD_FR.EditValue);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_PERIOD_NAME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ILA_PERIOD_NAME_COPY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERIOD_NAME.SetLookupParamValue("W_START_YYYYMM", DBNull.Value);
            ILD_PERIOD_NAME.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));
        }

        private void ilaPERIOD_NAME_0_SelectedRowData(object pSender)
        {
            SearchDB();
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(W_BUDGET_PERIOD_FR.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(W_BUDGET_PERIOD_TO.EditValue));
        }

        private void ilaBUDGET_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'APPLY'", "Y");
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUDGET_CAPACITY", DBNull.Value, "Y");
        }

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_CONTROL_YN", "N");
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBUDGET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'APPLY'", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }

        private void ILA_DEPT_COPY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ILD_DEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ILD_DEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(BUDGET_PERIOD.EditValue));
            ILD_DEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(BUDGET_PERIOD.EditValue));
        }


        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCAUSE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_CAUSE", "Value1 = 'APPLY'", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_BUDGET_ADD_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager == null)
            {
                V_BUDGET_AMOUNT.EditValue = 0;
                V_APPLY_SUM_AMOUNT.EditValue = 0;
                V_APPLY_AMOUNT.EditValue = 0;
                V_REMAIN_AMOUNT.EditValue = 0;
            }
            Set_Total_Amount();
        }

        private void IDA_BUDGET_ADD_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BUDGET_PERIOD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_PERIOD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["BUDGET_DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUDGET_DEPT_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REMARK))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_BUDGET_ADD_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion


        private void ILA_PERIOD_NAME_SelectedRowData(object pSender)
        {
            Sync_Button(BUDGET_PERIOD.EditValue);
        }

        private void IDA_BUDGET_APPLY_HEADER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Sync_Button("2017-07");
                return;
            }

            Sync_Button(pBindingManager.DataRow["BUDGET_PERIOD"]);
        }

        private void TB_MAIN_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}