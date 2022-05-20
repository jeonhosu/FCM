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

namespace EAPF0222
{
    public partial class EAPF0222 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public EAPF0222()
        {
            InitializeComponent();
        }

        public EAPF0222(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Set_Insert_Default()
        {
            DateTime vAPPLY_DATE = DateTime.Today;
            igrEXCHANGE_RATE.SetCellValue("APPLY_DATE", vAPPLY_DATE);
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    idaEXCHANGE_RATE.Fill();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaEXCHANGE_RATE.IsFocused)
                    {
                        idaEXCHANGE_RATE.AddOver();
                        Set_Insert_Default();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaEXCHANGE_RATE.IsFocused)
                    {
                        idaEXCHANGE_RATE.AddUnder();
                        Set_Insert_Default();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaEXCHANGE_RATE.Update();
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaEXCHANGE_RATE.IsFocused)
                    {
                        idaEXCHANGE_RATE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaEXCHANGE_RATE.IsFocused)
                    {
                        idaEXCHANGE_RATE.Delete();
                    }
                }
            }
        }

        #endregion;

        private void irbSORT_DATE_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            SORT_RULE_0.EditValue = Convert.ToInt32(iStatus.RadioCheckedString);

        }

        private void idaEXCHANGE_RATE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["FROM_CURRENCY"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=From Currency(From 통화)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["TO_CURRENCY"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=To Currency(To 통화)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["APPLY_PERIOD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Apply Date(기간)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SELLING_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Sell Rate(Sell 환율)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["SELLING_RATE"]) < iString.ISNumtoZero(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BUYING_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Buy Rate(Buy 환율)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["BUYING_RATE"]) < iString.ISNumtoZero(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BASE_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Use Rate(Use 환율)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["BASE_RATE"]) < iString.ISNumtoZero(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaEXCHANGE_RATE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void EAPF0222_Load(object sender, EventArgs e)
        {
            idaEXCHANGE_RATE.FillSchema();
            //irbSORT_DATE.RadioCheckedString = "1";
            irbSORT_DATE.Checked = true;
        }

        private void STD_DATE_FR_0_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            //if (iString.ISNull(STD_DATE_TO_0.EditValue) == string.Empty)
            //{
            //    STD_DATE_TO_0.EditValue = STD_DATE_FR_0.EditValue;
            //}
        }

        private void igrEXCHANGE_RATE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == igrEXCHANGE_RATE.GetColumnToIndex("BASE_RATE"))
            {
                if (iString.ISNumtoZero(igrEXCHANGE_RATE.GetCellValue("SELLING_RATE"), 0) == 0)
                {
                    igrEXCHANGE_RATE.SetCellValue("SELLING_RATE", e.NewValue);
                }
                if (iString.ISNumtoZero(igrEXCHANGE_RATE.GetCellValue("BUYING_RATE"), 0) == 0)
                {
                    igrEXCHANGE_RATE.SetCellValue("BUYING_RATE", e.NewValue);
                }
            }
        }

        private void ILA_DATE2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DATE2.SetLookupParamValue("W_START_PERIOD", V_DATE_FR.EditValue);
        }

        //private void igrEXCHANGE_RATE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        //{
        //    if (e.ColIndex == igrEXCHANGE_RATE.GetColumnToIndex("BASE_RATE"))
        //    {
        //        if (Convert.ToInt32(e.NewValue) < Convert.ToInt32(0))
        //        {   MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            e.Cancel = true;                
        //            return;
        //        }
        //    }
        //    if (e.ColIndex == igrEXCHANGE_RATE.GetColumnToIndex("SELLING_RATE"))
        //    {
        //        if (Convert.ToInt32(e.NewValue) < Convert.ToInt32(0))
        //        {
        //            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            e.Cancel = true;
        //            return;
        //        }
        //    }
        //    if (e.ColIndex == igrEXCHANGE_RATE.GetColumnToIndex("BUYING_RATE"))
        //    {
        //        if (Convert.ToInt32(e.NewValue) < Convert.ToInt32(0))
        //        {
        //            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            e.Cancel = true;
        //            return;
        //        }
        //    }
        //}
    }
}