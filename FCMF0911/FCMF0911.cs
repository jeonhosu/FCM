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

namespace FCMF0911
{
    public partial class FCMF0911 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0911()
        {
            InitializeComponent();
        }

        public FCMF0911(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            idaCLOSING_ACCOUNT.Fill();
            idaAUTO_JOURNAL.Fill();
            idaMISS_ACCOUNT.Fill();

            if (itbCLOSING_ACCOUNT.SelectedIndex == 0)
            {                
                igrCLOSING_ACCOUNT.Focus();
            }
            else if (itbCLOSING_ACCOUNT.SelectedIndex == 1)
            {                
                igrAUTO_JOURNAL.Focus();
            }
            else if (itbCLOSING_ACCOUNT.SelectedIndex == 2)
            {
                igrMISS_ACCOUNT.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
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
                    if (idaCLOSING_ACCOUNT.IsFocused)
                    {
                        idaCLOSING_ACCOUNT.AddOver();
                    }
                    else if (idaAUTO_JOURNAL.IsFocused)
                    {
                        idaAUTO_JOURNAL.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaCLOSING_ACCOUNT.IsFocused)
                    {
                        idaCLOSING_ACCOUNT.AddUnder();
                    }
                    else if (idaAUTO_JOURNAL.IsFocused)
                    {
                        idaAUTO_JOURNAL.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaCLOSING_ACCOUNT.IsFocused)
                    {
                        idaCLOSING_ACCOUNT.Update();
                    }
                    else if (idaAUTO_JOURNAL.IsFocused)
                    {
                        idaAUTO_JOURNAL.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaCLOSING_ACCOUNT.IsFocused)
                    {
                        idaCLOSING_ACCOUNT.Cancel();
                    }
                    else if (idaAUTO_JOURNAL.IsFocused)
                    {
                        idaAUTO_JOURNAL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaCLOSING_ACCOUNT.IsFocused)
                    {
                        idaCLOSING_ACCOUNT.Delete();
                    }
                    else if (idaAUTO_JOURNAL.IsFocused)
                    {
                        idaAUTO_JOURNAL.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void FCMF0911_Load(object sender, EventArgs e)
        {
            idaCLOSING_ACCOUNT.FillSchema();
            idaAUTO_JOURNAL.FillSchema();
        }

        private void FCMF0911_Shown(object sender, EventArgs e)
        {

        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaCLOSING_GROUP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CLOSING_GROUP", "Y");
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCLOSING_ACCOUNT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CLOSING_ACCOUNT_TYPE", "Y");
        }

        private void ilaCLOSING_GROUP_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CLOSING_GROUP", "Y");
        }

        private void ilaACCOUNT_CONTROL_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaCLOSING_ACCOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ENDING_ACCOUNT_CODE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Closing Code(결산 계정코드)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrCLOSING_ACCOUNT.CurrentCellMoveTo(igrCLOSING_ACCOUNT.GetColumnToIndex("CLOSING_GROUP_NAME"));
                return;
            }
            if (iString.ISNull(e.Row["ENDING_ACCOUNT_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Closing Code(결산 계정명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrCLOSING_ACCOUNT.CurrentCellMoveTo(igrCLOSING_ACCOUNT.GetColumnToIndex("ENDING_ACCOUNT_DESC"));
                return;
            } 
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrCLOSING_ACCOUNT.CurrentCellMoveTo(igrCLOSING_ACCOUNT.GetColumnToIndex("EFFECTIVE_DATE_FR"));
                return;
            }
        }

        private void idaCLOSING_ACCOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            { 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10307"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;            
            }
        }

        private void idaAUTO_JOURNAL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["CLOSING_GROUP"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Closing Group(결산그룹)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrAUTO_JOURNAL.CurrentCellMoveTo(igrAUTO_JOURNAL.GetColumnToIndex("CLOSING_GROUP_NAME"));
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrAUTO_JOURNAL.CurrentCellMoveTo(igrAUTO_JOURNAL.GetColumnToIndex("ACCOUNT_CODE"));
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrAUTO_JOURNAL.CurrentCellMoveTo(igrAUTO_JOURNAL.GetColumnToIndex("ACCOUNT_CODE"));
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_DR_CR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                igrAUTO_JOURNAL.CurrentCellMoveTo(igrAUTO_JOURNAL.GetColumnToIndex("ACCOUNT_DR_CR"));
                return;
            }
        }

        #endregion

    }
}