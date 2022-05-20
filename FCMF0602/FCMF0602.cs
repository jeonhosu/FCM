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

namespace FCMF0602
{
    public partial class FCMF0602 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0602()
        {
            InitializeComponent();
        }

        public FCMF0602(Form pMainForm, ISAppInterface pAppInterface)
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
            if (iString.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
             
            IDA_BUDGET_ADD_LIST.Fill();
            Set_Total_Amount();
            IGR_BUDGET_ADD_LINE.Focus();
        }

        private void Budget_Add_Insert()
        {
            IGR_BUDGET_ADD_LINE.SetCellValue("BUDGET_TYPE", W_BUDGET_TYPE.EditValue);
            IGR_BUDGET_ADD_LINE.SetCellValue("BUDGET_TYPE_NAME", W_BUDGET_TYPE_NAME.EditValue);
            IGR_BUDGET_ADD_LINE.SetCellValue("BUDGET_PERIOD", W_PERIOD_NAME.EditValue);

            int mIDX_COL = IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_TYPE_NAME");
            IGR_BUDGET_ADD_LINE.CurrentCellMoveTo(mIDX_COL);
            IGR_BUDGET_ADD_LINE.CurrentCellActivate(mIDX_COL);
        }

        private void SetCommonParameter(object pGroupCode, object pCodeName, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON.SetLookupParamValue("W_CODE_NAME", pCodeName);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SetCommonParameter_W(object pGroupCode, object pWhere, object pEnabled_YN)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Total_Amount()
        {
            decimal vTotal_Amount = 0;
            object vAmount;
            int vIDXCol = IGR_BUDGET_ADD_LINE.GetColumnToIndex("AMOUNT");
            for (int r = 0; r < IDA_BUDGET_ADD_LIST.SelectRows.Count; r++)
            {
                vAmount = 0;
                vAmount = IGR_BUDGET_ADD_LINE.GetCellValue(r, vIDXCol);
                vTotal_Amount = vTotal_Amount + iString.ISDecimaltoZero(vAmount);
            }
            TOTAL_AMOUNT.EditValue = vTotal_Amount;
        }

        private void EXE_BUDGET_ADD_STATUS(object pPERIOD_NAME, object pAPPROVE_STATUS, object pAPPROVE_FLAG)
        {
            IDA_BUDGET_ADD_LIST.Update(); //수정사항 반영.

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_CHECK_YN = IGR_BUDGET_ADD_LINE.GetColumnToIndex("CHECK_YN");
            int vIDX_BUDGET_TYPE = IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_TYPE");
            int vIDX_BUDGET_PERIOD = IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_PERIOD");
            int vIDX_DEPT_ID = IGR_BUDGET_ADD_LINE.GetColumnToIndex("DEPT_ID");
            int vIDX_ACCOUNT_CONTROL_ID = IGR_BUDGET_ADD_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID");
            
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_BUDGET_ADD_LINE.RowCount; i++)
            {
                if (iString.ISNull(IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_CHECK_YN), "N") == "Y")
                {
                    IGR_BUDGET_ADD_LINE.CurrentCellMoveTo(i, vIDX_CHECK_YN);
                    IGR_BUDGET_ADD_LINE.CurrentCellActivate(i, vIDX_CHECK_YN);

                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_BUDGET_TYPE));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_BUDGET_PERIOD));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_DEPT_ID));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_ACCOUNT_CONTROL_ID));
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("P_APPROVE_STATUS", pAPPROVE_STATUS);
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("P_APPROVE_FLAG", pAPPROVE_FLAG);
                    idcBUDGET_ADD_STATUS.SetCommandParamValue("P_CHECK_YN", IGR_BUDGET_ADD_LINE.GetCellValue(i, vIDX_CHECK_YN));
                    idcBUDGET_ADD_STATUS.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(idcBUDGET_ADD_STATUS.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(idcBUDGET_ADD_STATUS.GetCommandParamValue("O_MESSAGE"));
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();

                    if (idcBUDGET_ADD_STATUS.ExcuteError || vSTATUS == "F")
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
            int vIDX_CHECK = IGR_BUDGET_ADD_LINE.GetColumnToIndex("CHECK_YN");
            int mIDX_Col;

            // 신청금액.
            mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("AMOUNT");
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;

            // 신청사유.
            mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("CAUSE_NAME");
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
            // 비고.
            mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("REMARK");
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 0;
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 0;
            IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = true;
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
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[vIDX_CHECK].Insertable = 0;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[vIDX_CHECK].Updatable = 0;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[vIDX_CHECK].ReadOnly = true;
                }
                else
                {
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[vIDX_CHECK].Insertable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[vIDX_CHECK].Updatable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[vIDX_CHECK].ReadOnly = false;
                }

                if (mEnabled_YN == true)
                {
                    // 신청금액.
                    mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("AMOUNT");
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 신청사유.
                    mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("CAUSE_NAME");
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                    // 비고.
                    mIDX_Col = IGR_BUDGET_ADD_LINE.GetColumnToIndex("REMARK");
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Insertable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].Updatable = 1;
                    IGR_BUDGET_ADD_LINE.GridAdvExColElement[mIDX_Col].ReadOnly = false;
                }
            }
            IGR_BUDGET_ADD_LINE.ResetDraw = true;
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
            IDA_BUDGET_ADD_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_ADD_LIST.Refillable = true;
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
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.AddOver();
                    }
                    else
                    {
                        IDA_BUDGET_ADD_HEADER.AddOver();
                        Budget_Add_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.AddUnder();
                    }
                    else
                    {
                        IDA_BUDGET_ADD_HEADER.AddUnder();
                        Budget_Add_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_BUDGET_ADD_HEADER.Update();
                    }
                    catch
                    {
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.Cancel();
                    }
                    else
                    {
                        IDA_BUDGET_ADD_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BUDGET_ADD_LINE.IsFocused)
                    {
                        IDA_BUDGET_ADD_LINE.Delete();
                    }
                    else
                    {
                        IDA_BUDGET_ADD_HEADER.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0602_Load(object sender, EventArgs e)
        {
            IDA_BUDGET_ADD_LIST.FillSchema();
        }

        private void FCMF0602_Shown(object sender, EventArgs e)
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today); 
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }
         
        private void irbALL_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRADIO = sender as ISRadioButtonAdv;
            W_APPROVE_STATUS.EditValue = vRADIO.RadioButtonValue;

            //버튼제어 및 체크박스 제어.
            if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "A")
            {
                BTN_REQ_APPROVAL.Enabled = true;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
            }
            else if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "B")
            {
                BTN_REQ_APPROVAL.Enabled = true;
                BTN_CANCEL_REQ_APPROVAL.Enabled = true;
            }
            else if (iString.ISNull(W_APPROVE_STATUS.EditValue) == "C")
            {
                BTN_REQ_APPROVAL.Enabled = false;
                BTN_CANCEL_REQ_APPROVAL.Enabled = true;
            }
            else
            {
                BTN_REQ_APPROVAL.Enabled = false;
                BTN_CANCEL_REQ_APPROVAL.Enabled = false;
            }
            SearchDB();
        }
         
        private void IGR_BUDGET_ADD_LIST_CellDoubleClick(object pSender)
        {

        }

        private void ibtREQ_APPROVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_BUDGET_ADD_LIST.Update();

            object mValue;
            int mRowCount = IGR_BUDGET_ADD_LINE.RowCount;
            int mIDX_COL = IGR_BUDGET_ADD_LINE.GetColumnToIndex("APPROVE_STATUS");

            for (int R = 0; R < mRowCount; R++)
            {
                if (iString.ISNull(IGR_BUDGET_ADD_LINE.GetCellValue(R, mIDX_COL)) == "N".ToString())
                {// 승인미요청 건에 대해서 승인 처리.
                    idcAPPROVE_REQUEST.SetCommandParamValue("W_BUDGET_TYPE", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_TYPE")));
                    idcAPPROVE_REQUEST.SetCommandParamValue("W_BUDGET_PERIOD", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("BUDGET_PERIOD")));
                    idcAPPROVE_REQUEST.SetCommandParamValue("W_DEPT_ID", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("DEPT_ID")));
                    idcAPPROVE_REQUEST.SetCommandParamValue("W_ACCOUNT_CONTROL_ID", IGR_BUDGET_ADD_LINE.GetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("ACCOUNT_CONTROL_ID")));
                    idcAPPROVE_REQUEST.ExecuteNonQuery();

                    mValue = DBNull.Value;
                    mValue = idcAPPROVE_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                    IGR_BUDGET_ADD_LINE.SetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("APPROVE_STATUS"), mValue);

                    mValue = DBNull.Value;
                    mValue = idcAPPROVE_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                    IGR_BUDGET_ADD_LINE.SetCellValue(R, IGR_BUDGET_ADD_LINE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                }
            }
            IDA_BUDGET_ADD_LIST.OraSelectData.AcceptChanges();
            IDA_BUDGET_ADD_LIST.Refillable = true;
        }
        
        private void ibtnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            EXE_BUDGET_ADD_STATUS(W_PERIOD_NAME.EditValue, "A", "OK");
        }

        private void ibtnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            EXE_BUDGET_ADD_STATUS(W_PERIOD_NAME.EditValue, "C", "CANCEL");
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaPERIOD_NAME_0_SelectedRowData(object pSender)
        {
            SearchDB();
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(W_PERIOD_NAME.EditValue));
            ildDEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(W_PERIOD_NAME.EditValue));
        }

        private void ilaBUDGET_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'ADD'", "N");
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUDGET_CAPACITY", DBNull.Value, "N");
        }

        private void ilaACCOUNT_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_CONTROL_YN", "N");
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBUDGET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_TYPE", "Value1 = 'ADD'", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_DEPT_CODE_FR", null);
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT.SetLookupParamValue("W_CHECK_CAPACITY", "C");
            ildDEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", iDate.ISMonth_1st(W_PERIOD_NAME.EditValue));
            ildDEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", iDate.ISMonth_Last(W_PERIOD_NAME.EditValue));
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_BUDGET_CONTROL_YN", "N");
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCAUSE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_CAUSE", "Value1 = 'ADD'", "Y");
        }

        #endregion

        #region ----- Adapter Event -----
        
        private void idaBUDGET_ADD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void idaBUDGET_ADD_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        private void idaBUDGET_ADD_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {

        }

        #endregion




    }
}