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

namespace FCMF0113
{
    public partial class FCMF0113_COPY : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0113_COPY()
        {
            InitializeComponent();
        }

        public FCMF0113_COPY(Form pMainForm, ISAppInterface pAppInterface, object pForm_Type_ID, object pForm_Type_Name)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            FS_FORM_TYPE_ID.EditValue = pForm_Type_ID;
            FS_FORM_TYPE_NAME.EditValue = pForm_Type_Name;
        }

        #endregion;

        #region ----- Private Methods ----

        public object Get_FS_Form_Type_ID
        {
            get
            {
                return FS_FORM_TYPE_ID.EditValue;
            }
        }

        public object Get_New_FS_Form_Type_ID
        {
            get
            {
                return NEW_FS_FORM_TYPE_ID.EditValue;
            }
        }

        #endregion;

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    
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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                }
            }
        }

        #endregion;

        #region ----- Method -----

        private void SetCommonParameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_FS_FORM_TYPE.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_FS_FORM_TYPE.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        #endregion;

        #region ----- Form Events -----

        private void FCMF0113_COPY_COPY_Shown(object sender, EventArgs e)
        {
            
        }

        private void BTN_FS_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void bCREATE_FORM_MST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
         
        #endregion;

    }
}