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

namespace FCMF0609
{
    public partial class FCMF0609 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0609()
        {
            InitializeComponent();
        }

        public FCMF0609(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {            
            idaBUDGET_CONTROL.Fill();
            igrBUDGET_CONTROL.Focus();
        }

        private void SetCommonParameter(object pGroup_Code, object pENABLED_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pENABLED_YN);
        }

        private void SetCommonParameter_W(object pGroupCode, object pWhere, object pEnabled_YN)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", pGroupCode);
            ildCOMMON_W.SetLookupParamValue("W_WHERE", pWhere);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }


        private void Insert_Budget_Control()
        {
            igrBUDGET_CONTROL.SetCellValue("SLIP_YN", "Y");
            igrBUDGET_CONTROL.SetCellValue("BASE_YN", "Y");
            igrBUDGET_CONTROL.SetCellValue("ADD_YN", "Y");
            igrBUDGET_CONTROL.SetCellValue("MOVE_YN", "Y");
            igrBUDGET_CONTROL.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
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
                    if (idaBUDGET_CONTROL.IsFocused)
                    {
                        idaBUDGET_CONTROL.AddOver();
                        Insert_Budget_Control();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaBUDGET_CONTROL.IsFocused)
                    {
                        idaBUDGET_CONTROL.AddUnder();
                        Insert_Budget_Control();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaBUDGET_CONTROL.IsFocused)
                    {
                        idaBUDGET_CONTROL.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBUDGET_CONTROL.IsFocused)
                    {
                        idaBUDGET_CONTROL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBUDGET_CONTROL.IsFocused)
                    {
                        idaBUDGET_CONTROL.Delete();
                    }
                }
            }
        }

        #endregion;  
        
        #region ----- Form Event -----

        private void FCMF0609_Load(object sender, EventArgs e)
        {
            idaBUDGET_CONTROL.FillSchema();
        }

        private void FCMF0609_Shown(object sender, EventArgs e)
        {
            W_ENABLED_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_CAPACITY", "Value1 = 'Y'", "N");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildDEPT.SetLookupParamValue("W_EFFECTIVE_DATE_FR", DateTime.Today);
            ildDEPT.SetLookupParamValue("W_EFFECTIVE_DATE_TO", DateTime.Today);            
        }

        private void ilaAPPROVE_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter_W("BUDGET_CAPACITY", "Value1 = 'Y'", "Y");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaBUDGET_CONTROL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10019"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CAPACITY_LEVEL"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Capacity Leve(권한)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Person Name(사원)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Effective Date from(적용 시작일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaBUDGET_CONTROL_PreDelete(ISPreDeleteEventArgs e)
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