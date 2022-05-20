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

namespace FCMF0311
{
    public partial class FCMF0311 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0311()
        {
            InitializeComponent();
        }

        public FCMF0311(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            Set_Tab_Focus();
        }

        private void Insert_Asset_Category()
        {
            ENABLED_FLAG.CheckBoxValue = "Y";
            EFFECTIVE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            AST_CATEGORY_CODE.Focus();
        }

        private void Insert_Asset_Class()
        {
            ENABLED_FLAG_2.CheckBoxValue = "Y";
            EFFECTIVE_DATE_FR_2.EditValue = iDate.ISMonth_1st(DateTime.Today);
            AST_CLASS_CODE.Focus();
        }

        private void Insert_Asset_Item()
        {
            ENABLED_FLAG_3.CheckBoxValue = "Y";
            EFFECTIVE_DATE_FR_3.EditValue = iDate.ISMonth_1st(DateTime.Today);
            AST_ITEM_CODE.Focus();
        }

        private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_Tab_Focus()
        {
            if (TB_ASSET_CATEGORY.SelectedTab.TabIndex == TP_ASSET_CATEGORY.TabIndex)
            {
                IDA_ASSET_CATEGORY.Fill();
                igrASSET_CATEGORY.Focus();
            }
            else if (TB_ASSET_CATEGORY.SelectedTab.TabIndex == TP_ASSET_CLASS.TabIndex)
            {
                IDA_ASSET_CLASS.Fill();
                igrASSET_CLASS.Focus();
            }
            else if (TB_ASSET_CATEGORY.SelectedTab.TabIndex == TP_ASSET_ITEM.TabIndex)
            {
                IDA_ASSET_ITEM.Fill();
                igrASSET_ITEM.Focus();
            }
        }


        private void Get_DPR_Rate(object pDPR_TYPE, object pUSEFUL_LIFE)
        {
            IDC_DPR_RATE.SetCommandParamValue("W_DPR_TYPE", pDPR_TYPE);
            IDC_DPR_RATE.SetCommandParamValue("W_USEFUL_LIFE", pUSEFUL_LIFE);
            IDC_DPR_RATE.ExecuteNonQuery();
            IGR_ASSET_CATEGORY_DPR_RATE.SetCellValue("DPR_RATE", IDC_DPR_RATE.GetCommandParamValue("O_DPR_RATE"));
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
                    if (IDA_ASSET_CATEGORY.IsFocused)
                    {
                        IDA_ASSET_CATEGORY.AddOver();
                        Insert_Asset_Category();
                    }
                    else if (IDA_ASSET_CATEGORY_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_CATEGORY_DPR_RATE.AddOver();
                        IGR_ASSET_CATEGORY_DPR_RATE.Focus();
                    }
                    else if (IDA_ASSET_CLASS.IsFocused)
                    {
                        IDA_ASSET_CLASS.AddOver();
                        Insert_Asset_Class();
                    }
                    else if (IDA_ASSET_ITEM.IsFocused)
                    {
                        IDA_ASSET_ITEM.AddOver();
                        Insert_Asset_Item();
                    }  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ASSET_CATEGORY.IsFocused)
                    {
                        IDA_ASSET_CATEGORY.AddUnder();
                        Insert_Asset_Category();
                    }
                    else if (IDA_ASSET_CATEGORY_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_CATEGORY_DPR_RATE.AddUnder();
                        IGR_ASSET_CATEGORY_DPR_RATE.Focus();
                    }
                    else if (IDA_ASSET_CLASS.IsFocused)
                    {
                        IDA_ASSET_CLASS.AddUnder();
                        Insert_Asset_Class();
                    }
                    else if (IDA_ASSET_ITEM.IsFocused)
                    {
                        IDA_ASSET_ITEM.AddUnder();
                        Insert_Asset_Item();
                    }  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_ASSET_CATEGORY.IsFocused || IDA_ASSET_CATEGORY_DPR_RATE.IsFocused || IDA_ASSET_CATEGORY_ACCOUNT.IsFocused)
                    {
                        IDA_ASSET_CATEGORY.Update();
                    }
                    else if (IDA_ASSET_CLASS.IsFocused)
                    {
                        IDA_ASSET_CLASS.Update();
                    }
                    else if (IDA_ASSET_ITEM.IsFocused)
                    {
                        IDA_ASSET_ITEM.Update();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ASSET_CATEGORY.IsFocused)
                    {
                        IDA_ASSET_CATEGORY.Cancel();
                        IDA_ASSET_CATEGORY_DPR_RATE.Cancel();
                    }
                    else if (IDA_ASSET_CATEGORY_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_CATEGORY_DPR_RATE.Cancel();
                    }
                    else if (IDA_ASSET_CLASS.IsFocused)
                    {
                        IDA_ASSET_CLASS.Cancel();
                    }
                    else if (IDA_ASSET_ITEM.IsFocused)
                    {
                        IDA_ASSET_ITEM.Cancel();
                    }  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ASSET_CATEGORY.IsFocused)
                    {
                        IDA_ASSET_CATEGORY.Delete();
                    }
                    else if (IDA_ASSET_CATEGORY_DPR_RATE.IsFocused)
                    {
                        IDA_ASSET_CATEGORY_DPR_RATE.Delete();
                    }
                    else if (IDA_ASSET_CLASS.IsFocused)
                    {
                        IDA_ASSET_CLASS.Delete();
                    }
                    else if (IDA_ASSET_ITEM.IsFocused)
                    {
                        IDA_ASSET_ITEM.Delete();
                    }  
                }
            }
        }

        #endregion;

        #region ----- Component Event -----

        private void FCMF0311_Load(object sender, EventArgs e)
        {
            IDA_ASSET_CATEGORY.FillSchema();
            //IDA_ASSET_CLASS.FillSchema();
            //IDA_ASSET_ITEM.FillSchema();
        }

        private void FCMF0311_Shown(object sender, EventArgs e)
        {
            ENABLED_1_FLAG_0.CheckBoxValue = "Y";
            ENABLED_2_FLAG_0.CheckBoxValue = "Y";
            ENABLED_3_FLAG_0.CheckBoxValue = "Y";
        }

        private void itbASSET_CATEGORY_Click(object sender, EventArgs e)
        {
            Set_Tab_Focus();
        }

        private void ASSET_CATEGORY_DPR_RATE_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            int vIDX_USEFUL_LIFE = IGR_ASSET_CATEGORY_DPR_RATE.GetColumnToIndex("USEFUL_LIFE");
            if (e.ColIndex == vIDX_USEFUL_LIFE)
            {
                Get_DPR_Rate(IGR_ASSET_CATEGORY_DPR_RATE.GetCellValue("DPR_TYPE"), e.CellValue);
            }
        }

        private void IGR_ASSET_CATEGORY_ACCOUNT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (IGR_ASSET_CATEGORY_ACCOUNT.GetColumnToIndex("ENABLED_FLAG") == e.ColIndex)
            {
                IGR_ASSET_CATEGORY_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaASSET_CATEGORY_1_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaASSET_CLASS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", ASSET_CATEGORY_2_ID_0.EditValue);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 2);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaASSET_CATEGORY_2_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ILA_ASSET_CATEGORY_3_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaEXPENSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("EXPENSE_TYPE", "Y");
        }

        private void ilaASSET_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_TYPE", "Y");
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaUPPER_ASSET_CATEGORY_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        
        private void ilaUPPER_ASSET_CLASS_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 2);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
               
        private void ilaASSET_CLASS_3_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", DBNull.Value);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 2);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

         private void ilaASSET_ITEM_3_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildASSET_CATEGORY.SetLookupParamValue("W_UPPER_AST_CATEGORY_ID", ASSET_CLASS_3_ID_0.EditValue);
            ildASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 3);
            ildASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaDPR_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("DPR_TYPE", "Y");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaDPR_TYPE_SelectedRowData(object pSender)
        {
            Get_DPR_Rate(IGR_ASSET_CATEGORY_DPR_RATE.GetCellValue("DPR_TYPE"), IGR_ASSET_CATEGORY_DPR_RATE.GetCellValue("USEFUL_LIFE"));
        }

        private void ilaACCOUNT_CONTROL_SelectedRowData(object pSender)
        {
            if(iString.ISNull(IGR_ASSET_CATEGORY_ACCOUNT.GetCellValue("ENABLED_FLAG"), "N") == "N")
            {
                IGR_ASSET_CATEGORY_ACCOUNT.SetCellValue("ENABLED_FLAG", "Y");
                IGR_ASSET_CATEGORY_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            }
        }

        private void ILA_USEFUL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("USEFUL_TYPE", "Y");
        }

        #endregion

        #region ----- Adapeter Event -----

        private void idaASSET_CATEGORY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["AST_CATEGORY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10093"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AST_CATEGORY_NAME"]) == string.Empty)
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

        private void idaASSET_CLASS_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["AST_CLASS_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10099"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AST_CLASS_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10100"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["UPPER_AST_CATEGORY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10101"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
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

        private void idaASSET_CLASS_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaASSET_ITEM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["AST_ITEM_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Asset Item Code(소분류 코드)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AST_ITEM_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Asset Item Name(소분류명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["UPPER_AST_CLASS_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Asset Class(자산 중분류)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
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

        private void idaASSET_ITEM_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_ASSET_CATEGORY_ACCOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ENABLED_FLAG"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10085"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_ASSET_CATEGORY_ACCOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_ASSET_CATEGORY_DPR_RATE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DPR_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10097"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USEFUL_LIFE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10098"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}