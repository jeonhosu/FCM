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

namespace FCMF0615
{
    public partial class FCMF0615 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0615()
        {
            InitializeComponent();
        }

        public FCMF0615(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            idaSLIP_HEADER_LIST.Fill();
            igrSLIP_LIST_IF.Focus();
        }

        private void Search_DB_DETAIL(object pSLIP_HEADER_ID)
        {
            if (iString.ISNull(pSLIP_HEADER_ID) != string.Empty)
            {
                itbSLIP.SelectedIndex = 1;
                itbSLIP.SelectedTab.Focus();
                idaSLIP_HEADER.SetSelectParamValue("W_HEADER_ID", pSLIP_HEADER_ID);

                idaSLIP_HEADER.Fill();

                idaSLIP_LINE.OraSelectData.AcceptChanges();
                idaSLIP_LINE.Refillable = true;
                idaSLIP_HEADER.OraSelectData.AcceptChanges();
                idaSLIP_HEADER.Refillable = true;
            }
        }

        #endregion;

        #region ----- Initialize Event -----

        private Boolean Check_SlipHeader_Added()
        {
            Boolean Row_Added_Status = false;
            for (int r = 0; r < idaSLIP_HEADER.SelectRows.Count; r++)
            {
                if (idaSLIP_HEADER.SelectRows[r].RowState == DataRowState.Added)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10261"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return (Row_Added_Status);
        }

        private void Init_Total_GL_Amount()
        {
            decimal vDR_Amount = Convert.ToDecimal(0);
            if (igrSLIP_LINE.RowCount > 0)
            {
                for (int r = 0; r < igrSLIP_LINE.RowCount; r++)
                {
                    vDR_Amount = vDR_Amount + iString.ISDecimaltoZero(igrSLIP_LINE.GetCellValue(r, igrSLIP_LINE.GetColumnToIndex("GL_AMOUNT")));
                }
            }
            TOTAL_AMOUNT.EditValue = iString.ISDecimaltoZero(vDR_Amount);
        }

        private void Init_Control_Management_Value()
        {
            igrSLIP_LINE.SetCellValue("MANAGEMENT1", null);
            igrSLIP_LINE.SetCellValue("MANAGEMENT1_DESC", null);
            igrSLIP_LINE.SetCellValue("MANAGEMENT2", null);
            igrSLIP_LINE.SetCellValue("MANAGEMENT2_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER1", null);
            igrSLIP_LINE.SetCellValue("REFER1_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER2", null);
            igrSLIP_LINE.SetCellValue("REFER2_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER3", null);
            igrSLIP_LINE.SetCellValue("REFER3_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER4", null);
            igrSLIP_LINE.SetCellValue("REFER4_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER5", null);
            igrSLIP_LINE.SetCellValue("REFER5_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER6", null);
            igrSLIP_LINE.SetCellValue("REFER6_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER7", null);
            igrSLIP_LINE.SetCellValue("REFER7_DESC", null);
            igrSLIP_LINE.SetCellValue("REFER8", null);
            igrSLIP_LINE.SetCellValue("REFER8_DESC", null);
        }

        private void Init_Set_Item_Prompt(DataRow pDataRow)
        {
            if (pDataRow == null)
            {
                return;
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            MANAGEMENT1.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["MANAGEMENT1_YN"], "F") == "F".ToString())
            //{
            MANAGEMENT1.Nullable = true;
            MANAGEMENT1.ReadOnly = true;
            MANAGEMENT1.Insertable = false;
            MANAGEMENT1.Updatable = false;
            MANAGEMENT1.TabStop = false;
            MANAGEMENT1.Refresh();
            //}
            //else
            //{
            //MANAGEMENT1.Nullable = true;
            //MANAGEMENT1.ReadOnly = false;
            //MANAGEMENT1.Insertable = true;
            //MANAGEMENT1.Updatable = true;
            //MANAGEMENT1.TabStop = true;
            if (iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]) == "NUMBER".ToString())
            {
                MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]) == "RATE".ToString())
            {
                MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                MANAGEMENT1.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["MANAGEMENT1_DATA_TYPE"]) == "DATE".ToString())
            {
                MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                MANAGEMENT1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["MANAGEMENT1_YN"], "N") == "Y".ToString())
                {
                    MANAGEMENT1.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["MANAGEMENT1_YN"], "N") == "Y".ToString())
            {
                MANAGEMENT1.ReadOnly = false;
            }
            MANAGEMENT1.Refresh();
            //}
            MANAGEMENT2.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["MANAGEMENT2_YN"], "F") == "F".ToString())
            //{
            MANAGEMENT2.Nullable = true;
            MANAGEMENT2.ReadOnly = true;
            MANAGEMENT2.Insertable = false;
            MANAGEMENT2.Updatable = false;
            MANAGEMENT2.TabStop = false;
            MANAGEMENT2.Refresh();
            //}
            //else
            //{
            //    MANAGEMENT2.Nullable = true;
            //    MANAGEMENT2.ReadOnly = false;
            //    MANAGEMENT2.Insertable = true;
            //    MANAGEMENT2.Updatable = true;
            //    MANAGEMENT2.TabStop = true;
            if (iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]) == "NUMBER".ToString())
            {
                MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]) == "RATE".ToString())
            {
                MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                MANAGEMENT2.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["MANAGEMENT2_DATA_TYPE"]) == "DATE".ToString())
            {
                MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                MANAGEMENT2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["MANAGEMENT2_YN"], "N") == "Y".ToString())
                {
                    MANAGEMENT2.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["MANAGEMENT2_YN"], "N") == "Y".ToString())
            {
                MANAGEMENT2.ReadOnly = false;
            }
            MANAGEMENT2.Refresh();
            //}
            REFER1.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER1_YN"], "F") == "F".ToString())
            //{
            REFER1.Nullable = true;
            REFER1.ReadOnly = true;
            REFER1.Insertable = false;
            REFER1.Updatable = false;
            REFER1.TabStop = false;
            REFER1.Refresh();
            //}
            //else
            //{
            //    REFER1.Nullable = true;
            //    REFER1.ReadOnly = false;
            //    REFER1.Insertable = true;
            //    REFER1.Updatable = true;
            //    REFER1.TabStop = true;
            if (iString.ISNull(pDataRow["REFER1_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER1_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER1.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER1.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER1_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER1.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER1.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER1_YN"], "N") == "Y".ToString())
                {
                    REFER1.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER1_YN"], "N") == "Y".ToString())
            {
                REFER1.ReadOnly = false;
            }
            REFER1.Refresh();
            //}
            REFER2.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER2_YN"], "F") == "F".ToString())
            //{
            REFER2.Nullable = true;
            REFER2.ReadOnly = true;
            REFER2.Insertable = false;
            REFER2.Updatable = false;
            REFER2.TabStop = false;
            REFER2.Refresh();
            //}
            //else
            //{
            //    REFER2.Nullable = true;
            //    REFER2.ReadOnly = false;
            //    REFER2.Insertable = true;
            //    REFER2.Updatable = true;
            //    REFER2.TabStop = true;
            if (iString.ISNull(pDataRow["REFER2_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER2_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER2.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER2.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER2_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER2.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER2.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER2_YN"], "N") == "Y".ToString())
                {
                    REFER2.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER2_YN"], "N") == "Y".ToString())
            {
                REFER2.ReadOnly = false;
            }
            REFER2.Refresh();
            //}
            REFER3.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER3_YN"], "F") == "F".ToString())
            //{
            REFER3.Nullable = true;
            REFER3.ReadOnly = true;
            REFER3.Insertable = false;
            REFER3.Updatable = false;
            REFER3.TabStop = false;
            REFER3.Refresh();
            //}
            //else
            //{
            //    REFER3.Nullable = true;
            //    REFER3.ReadOnly = false;
            //    REFER3.Insertable = true;
            //    REFER3.Updatable = true;
            //    REFER3.TabStop = true;
            if (iString.ISNull(pDataRow["REFER3_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER3_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER3.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER3.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER3_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER3.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER3.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER3_YN"], "N") == "Y".ToString())
                {
                    REFER3.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER3_YN"], "N") == "Y".ToString())
            {
                REFER3.ReadOnly = false;
            }
            REFER3.Refresh();
            //}
            REFER4.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER4_YN"], "F") == "F".ToString())
            //{
            REFER4.Nullable = true;
            REFER4.ReadOnly = true;
            REFER4.Insertable = false;
            REFER4.Updatable = false;
            REFER4.TabStop = false;
            REFER4.Refresh();
            //}
            //else
            //{
            //    REFER4.Nullable = true;
            //    REFER4.ReadOnly = false;
            //    REFER4.Insertable = true;
            //    REFER4.Updatable = true;
            //    REFER4.TabStop = true;
            if (iString.ISNull(pDataRow["REFER4_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER4_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER4.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER4.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER4_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER4.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER4.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER4_YN"], "N") == "Y".ToString())
                {
                    REFER4.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER4_YN"], "N") == "Y".ToString())
            {
                REFER4.ReadOnly = false;
            }
            REFER4.Refresh();
            //}
            REFER5.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER5_YN"], "F") == "F".ToString())
            //{
            REFER5.Nullable = true;
            REFER5.ReadOnly = true;
            REFER5.Insertable = false;
            REFER5.Updatable = false;
            REFER5.TabStop = false;
            REFER5.Refresh();
            //}
            //else
            //{
            //    REFER5.Nullable = true;
            //    REFER5.ReadOnly = false;
            //    REFER5.Insertable = true;
            //    REFER5.Updatable = true;
            //    REFER5.TabStop = true;
            if (iString.ISNull(pDataRow["REFER5_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER5_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER5.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER5.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER5_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER5.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER5.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER5_YN"], "N") == "Y".ToString())
                {
                    REFER5.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER5_YN"], "N") == "Y".ToString())
            {
                REFER5.ReadOnly = false;
            }
            REFER5.Refresh();
            //}
            REFER6.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER6_YN"], "F") == "F".ToString())
            //{
            REFER6.Nullable = true;
            REFER6.ReadOnly = true;
            REFER6.Insertable = false;
            REFER6.Updatable = false;
            REFER6.TabStop = false;
            REFER6.Refresh();
            //}
            //else
            //{
            //    REFER6.Nullable = true;
            //    REFER6.ReadOnly = false;
            //    REFER6.Insertable = true;
            //    REFER6.Updatable = true;
            //    REFER6.TabStop = true;
            if (iString.ISNull(pDataRow["REFER6_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER6_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER6.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER6.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER6_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER6.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER6.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER6_YN"], "N") == "Y".ToString())
                {
                    REFER6.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER6_YN"], "N") == "Y".ToString())
            {
                REFER6.ReadOnly = false;
            }
            REFER6.Refresh();
            //}
            REFER7.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER7_YN"], "F") == "F".ToString())
            //{
            REFER7.Nullable = true;
            REFER7.ReadOnly = true;
            REFER7.Insertable = false;
            REFER7.Updatable = false;
            REFER7.TabStop = false;
            REFER7.Refresh();
            //}
            //else
            //{
            //    REFER7.Nullable = true;
            //    REFER7.ReadOnly = false;
            //    REFER7.Insertable = true;
            //    REFER7.Updatable = true;
            REFER7.TabStop = true;
            if (iString.ISNull(pDataRow["REFER7_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER7_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER7.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER7.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER7_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER7.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER7.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER7_YN"], "N") == "Y".ToString())
                {
                    REFER7.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER7_YN"], "N") == "Y".ToString())
            {
                REFER7.ReadOnly = false;
            }
            REFER7.Refresh();
            //}
            REFER8.NumberDecimalDigits = 0;
            //if (iString.ISNull(pDataRow["REFER8_YN"], "F") == "F".ToString())
            //{
            REFER8.Nullable = true;
            REFER8.ReadOnly = true;
            REFER8.Insertable = false;
            REFER8.Updatable = false;
            REFER8.TabStop = false;
            REFER8.Refresh();
            //}
            //else
            //{
            //    REFER8.Nullable = true;
            //    REFER8.ReadOnly = false;
            //    REFER8.Insertable = true;
            //    REFER8.Updatable = true;
            //    REFER8.TabStop = true;
            if (iString.ISNull(pDataRow["REFER8_DATA_TYPE"]) == "NUMBER".ToString())
            {
                REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
            }
            else if (iString.ISNull(pDataRow["REFER8_DATA_TYPE"]) == "RATE".ToString())
            {
                REFER8.EditAdvType = ISUtil.Enum.EditAdvType.NumberEdit;
                REFER8.NumberDecimalDigits = 4;
            }
            else if (iString.ISNull(pDataRow["REFER8_DATA_TYPE"]) == "DATE".ToString())
            {
                REFER8.EditAdvType = ISUtil.Enum.EditAdvType.DateTimeEdit;
            }
            else
            {
                REFER8.EditAdvType = ISUtil.Enum.EditAdvType.TextEdit;
                if (iString.ISNull(pDataRow["REFER8_YN"], "N") == "Y".ToString())
                {
                    REFER8.Nullable = false;
                }
            }
            if (iString.ISNull(pDataRow["REFER8_YN"], "N") == "Y".ToString())
            {
                REFER8.ReadOnly = false;
            }
            REFER8.Refresh();
            //}
            ///////////////////////////////////////////////////////////////////////////////////////////////////            
            //if (iString.ISNull(pDataRow["MANAGEMENT1_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    MANAGEMENT1.LookupAdapter = ilaMANAGEMENT1;
            //}
            //else
            //{
            //    MANAGEMENT1.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["MANAGEMENT2_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    MANAGEMENT2.LookupAdapter = ilaMANAGEMENT2;
            //}
            //else
            //{
            //    MANAGEMENT2.LookupAdapter = null;
            //}
            //if (iString.ISNull(pDataRow["REFER1_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER1.LookupAdapter = ilaREFER1;
            //}
            //else
            //{
            //    REFER1.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER2_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER2.LookupAdapter = ilaREFER2;
            //}
            //else
            //{
            //    REFER2.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER3_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER3.LookupAdapter = ilaREFER3;
            //}
            //else
            //{
            //    REFER3.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER4_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER4.LookupAdapter = ilaREFER4;
            //}
            //else
            //{
            //    REFER4.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER5_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER5.LookupAdapter = ilaREFER5;
            //}
            //else
            //{
            //    REFER5.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER6_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER6.LookupAdapter = ilaREFER6;
            //}
            //else
            //{
            //    REFER6.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER7_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER7.LookupAdapter = ilaREFER7;
            //}
            //else
            //{
            //    REFER7.LookupAdapter = null;
            //}

            //if (iString.ISNull(pDataRow["REFER8_LOOKUP_YN"], "N") == "Y".ToString())
            //{
            //    REFER8.LookupAdapter = ilaREFER8;
            //}
            //else
            //{
            //    REFER8.LookupAdapter = null;
            //}
        }

        private void Init_Set_Item_Need(DataRow pDataRow)
        {// 관리항목 필수여부 세팅.
            if (pDataRow == null)
            {
                return;
            }

            object mDATA_VALUE;
            object mDATA_TYPE;
            object mDR_CR_YN = "N";
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            //--1
            mDATA_VALUE = MANAGEMENT1.EditValue;
            MANAGEMENT1.Nullable = true;
            mDATA_TYPE = pDataRow["MANAGEMENT1_DATA_TYPE"];
            mDR_CR_YN = pDataRow["MANAGEMENT1_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT1_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                MANAGEMENT1.Nullable = false;
            }
            MANAGEMENT1.EditValue = mDATA_VALUE;
            MANAGEMENT1.Refresh();
            //--2
            mDATA_VALUE = MANAGEMENT2.EditValue;
            MANAGEMENT2.Nullable = true;
            mDATA_TYPE = pDataRow["MANAGEMENT2_DATA_TYPE"];
            mDR_CR_YN = pDataRow["MANAGEMENT2_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["MANAGEMENT2_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                MANAGEMENT2.Nullable = false;
            }
            MANAGEMENT2.Refresh();
            MANAGEMENT2.EditValue = mDATA_VALUE;
            //--3
            mDATA_VALUE = REFER1.EditValue;
            REFER1.Nullable = true;
            mDATA_TYPE = pDataRow["REFER1_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER1_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER1_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER1.Nullable = false;
            }
            REFER1.Refresh();
            REFER1.EditValue = mDATA_VALUE;
            //--4
            mDATA_VALUE = REFER2.EditValue;
            REFER2.Nullable = true;
            mDATA_TYPE = pDataRow["REFER2_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER2_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER2_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER2.Nullable = false;
            }
            REFER2.Refresh();
            REFER2.EditValue = mDATA_VALUE;
            //--5
            mDATA_VALUE = REFER3.EditValue;
            REFER3.Nullable = true;
            mDATA_TYPE = pDataRow["REFER3_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER3_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER3_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER3.Nullable = false;
            }
            REFER3.Refresh();
            REFER3.EditValue = mDATA_VALUE;
            //--6
            mDATA_VALUE = REFER4.EditValue;
            REFER4.Nullable = true;
            mDATA_TYPE = pDataRow["REFER4_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER4_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER4_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER4.Nullable = false;
            }
            REFER4.Refresh();
            REFER4.EditValue = mDATA_VALUE;
            //--7
            mDATA_VALUE = REFER5.EditValue;
            REFER5.Nullable = true;
            mDATA_TYPE = pDataRow["REFER5_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER5_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER5_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER5.Nullable = false;
            }
            REFER5.Refresh();
            REFER5.EditValue = mDATA_VALUE;
            //--8
            mDATA_VALUE = REFER6.EditValue;
            REFER6.Nullable = true;
            mDATA_TYPE = pDataRow["REFER6_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER6_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER6_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER6.Nullable = false;
            }
            REFER6.Refresh();
            REFER6.EditValue = mDATA_VALUE;
            //--9
            mDATA_VALUE = REFER7.EditValue;
            REFER7.Nullable = true;
            mDATA_TYPE = pDataRow["REFER7_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER7_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = pDataRow["REFER7_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER7.Nullable = false;
            }
            REFER7.Refresh();
            REFER7.EditValue = mDATA_VALUE;
            //--10
            mDATA_VALUE = REFER8.EditValue;
            REFER8.Nullable = true;
            mDATA_TYPE = pDataRow["REFER8_DATA_TYPE"];
            mDR_CR_YN = pDataRow["REFER8_YN"];
            //if (iString.ISNull(pACCOUNT_DR_CR) == "1")
            //{
            //    mDR_CR_YN = igrSLIP_LINE.GetCellValue("REFER8_DR_YN"];
            //}
            //else if (iString.ISNull(pACCOUNT_DR_CR) == "2")
            //{
            //    mDR_CR_YN = igrSLIP_LINE.GetCellValue("REFER8_CR_YN"];
            //}
            if (iString.ISNull(mDATA_TYPE) == "VARCHAR2" && iString.ISNull(mDR_CR_YN) == "Y")
            {
                REFER8.Nullable = false;
            }
            REFER8.Refresh();
            REFER8.EditValue = mDATA_VALUE;
        }

        private void Init_Default_Value()
        {
            int mPreviousRowPosition = idaSLIP_LINE.CurrentRowPosition() - 1;
            object mPrevious_Code;
            object mPrevious_Name;
            string mData_Type;
            string mLookup_Type;

            if (mPreviousRowPosition > -1
                && iString.ISNull(REMARK.EditValue) == string.Empty
                && iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"]) != string.Empty)
            {//REMARK.
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REMARK"];
                REMARK.EditValue = mPrevious_Name;
            }

            //1
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT1.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT1.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_LOOKUP_TYPE"]))
            {//MANAGEMENT1_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT1_DESC"];

                MANAGEMENT1.EditValue = mPrevious_Code;
                MANAGEMENT1_DESC.EditValue = mPrevious_Name;
            }
            //2
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT2_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(MANAGEMENT2.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    MANAGEMENT2.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_LOOKUP_TYPE"]))
            {//MANAGEMENT2_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["MANAGEMENT2_DESC"];

                MANAGEMENT2.EditValue = mPrevious_Code;
                MANAGEMENT2_DESC.EditValue = mPrevious_Name;
            }
            //3
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER1_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER1.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER1.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_LOOKUP_TYPE"]))
            {//REFER1_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER1_DESC"];

                REFER1.EditValue = mPrevious_Code;
                REFER1_DESC.EditValue = mPrevious_Name;
            }
            //4
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER2_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER2.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER2.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_LOOKUP_TYPE"]))
            {//REFER2_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER2_DESC"];

                REFER2.EditValue = mPrevious_Code;
                REFER2_DESC.EditValue = mPrevious_Name;
            }
            //5
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER3_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER3.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER3.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_LOOKUP_TYPE"]))
            {//REFER3_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER3_DESC"];

                REFER3.EditValue = mPrevious_Code;
                REFER3_DESC.EditValue = mPrevious_Name;
            }
            //6
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER4_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER4.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER4.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_LOOKUP_TYPE"]))
            {//REFER4_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER4_DESC"];

                REFER4.EditValue = mPrevious_Code;
                REFER4_DESC.EditValue = mPrevious_Name;
            }
            //7
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER5_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER5.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER5.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_LOOKUP_TYPE"]))
            {//REFER5_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER5_DESC"];

                REFER5.EditValue = mPrevious_Code;
                REFER5_DESC.EditValue = mPrevious_Name;
            }
            //8
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER6_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER6.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER6.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_LOOKUP_TYPE"]))
            {//REFER6_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER6_DESC"];

                REFER6.EditValue = mPrevious_Code;
                REFER6_DESC.EditValue = mPrevious_Name;
            }
            //9
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER7_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER7.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER7.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_LOOKUP_TYPE"]))
            {//REFER7_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER7_DESC"];

                REFER7.EditValue = mPrevious_Code;
                REFER7_DESC.EditValue = mPrevious_Name;
            }
            //10
            mData_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_DATA_TYPE"));
            mLookup_Type = iString.ISNull(igrSLIP_LINE.GetCellValue("REFER8_LOOKUP_TYPE"));
            if (mData_Type == "NUMBER".ToString())
            {
            }
            else if (mData_Type == "RATE".ToString())
            {
            }
            else if (mData_Type == "DATE".ToString())
            {
                if (iString.ISNull(REFER8.EditValue) == string.Empty && iString.ISNull(SLIP_DATE.EditValue) != string.Empty)
                {
                    REFER8.EditValue = Convert.ToDateTime(SLIP_DATE.EditValue).ToShortDateString();
                }
            }
            if (mPreviousRowPosition > -1
                && mLookup_Type != string.Empty
                && mLookup_Type == iString.ISNull(idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_LOOKUP_TYPE"]))
            {//REFER8_LOOKUP_TYPE
                mPrevious_Code = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8"];
                mPrevious_Name = idaSLIP_LINE.CurrentRows[mPreviousRowPosition]["REFER8_DESC"];

                REFER8.EditValue = mPrevious_Code;
                REFER8_DESC.EditValue = mPrevious_Name;
            }
        }

        private bool Validate_Date(object pValue)
        {
            bool mValidate = false;

            if (iString.ISNull(pValue).Length != 10)
            {
                return mValidate;
            }
            try
            {
                DateTime mDate = Convert.ToDateTime(pValue);
            }
            catch
            {
                return mValidate;
            }

            mValidate = true;
            return mValidate;
        }

        // 부가세 관련 설정 제어 - 세액/공급가액(세액 * 10)
        private void Init_VAT_Amount()
        {
            object mVAT_ENABLED_FLAG = igrSLIP_LINE.GetCellValue("VAT_ENABLED_FLAG");
            if (iString.ISNull(mVAT_ENABLED_FLAG, "N") != "Y")
            {
                return;
            }

            Decimal mGL_AMOUNT = iString.ISDecimaltoZero(GL_AMOUNT.EditValue);
            REFER1.EditValue = mGL_AMOUNT * 10; //공급가액 설정.
        }

        //부서 
        private void Init_Dept()
        {
            if (iString.ISNull(igrSLIP_LINE.GetCellValue("MANAGEMENT1_LOOKUP_TYPE")) == "DEPT" && iString.ISNull(MANAGEMENT1.EditValue) == String.Empty)
            {
                MANAGEMENT1_DESC.EditValue = BUDGET_DEPT_NAME_L.EditValue;
                MANAGEMENT1.EditValue = BUDGET_DEPT_CODE_L.EditValue;
            }
        }

        //관리항목 LOOKUP 선택시 처리.
        private void Init_SELECT_LOOKUP(object pManagement_Type)
        {
            string mMANAGEMENT = iString.ISNull(pManagement_Type);
        }

        #endregion

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
                    if (idaSLIP_HEADER_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_LIST.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaSLIP_HEADER_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_LIST.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaSLIP_HEADER_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_LIST.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaSLIP_HEADER_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaSLIP_HEADER_LIST.IsFocused)
                    {
                        idaSLIP_HEADER_LIST.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Events -----

        private void igrSLIP_LIST_IF_CellDoubleClick(object pSender)
        {
            if (igrSLIP_LIST_IF.Row > 0)
            {
                Search_DB_DETAIL(igrSLIP_LIST_IF.GetCellValue("HEADER_ID"));
            }
        }

        private void idaSLIP_LINE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Total_GL_Amount();
        }

        private void idaSLIP_LINE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Set_Item_Prompt(pBindingManager.DataRow);
        }

        private void FCMF0615_Shown(object sender, EventArgs e)
        {
            SLIP_DATE_FR_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            SLIP_DATE_TO_0.EditValue = iDate.ISGetDate();
        }

        #endregion;

        #region ----- Lookup Event -----

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

    }
}