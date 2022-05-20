using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace EAPF0107
{
    public partial class EAPF0107 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public EAPF0107()
        {
            InitializeComponent();
        }

        public EAPF0107(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            //if (iString.ISNull(AUTHORITY_GROUP_NAME.EditValue) == string.Empty)
            //{// ±ÇÇÑ.
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10120"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    AUTHORITY_GROUP_NAME.Focus();
            //    return;
            //}

            IDA_AUTHORITY_GROUP_H.Fill();

            IGR_AUTHORITY_GROUP_H.Focus();
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
                    if (IDA_AUTHORITY_GROUP_H.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_H.AddOver();

                        IGR_AUTHORITY_GROUP_H.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
                        IGR_AUTHORITY_GROUP_H.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    else if (IDA_AUTHORITY_GROUP_L.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_L.AddOver();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_AUTHORITY_GROUP_H.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_H.AddUnder();

                        IGR_AUTHORITY_GROUP_H.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
                        IGR_AUTHORITY_GROUP_H.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    else if (IDA_AUTHORITY_GROUP_L.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_L.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_AUTHORITY_GROUP_H.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_AUTHORITY_GROUP_H.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_L.Cancel();
                        IDA_AUTHORITY_GROUP_H.Cancel();
                    }
                    else if (IDA_AUTHORITY_GROUP_L.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_L.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_AUTHORITY_GROUP_H.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_H.Delete();
                    }
                    else if (IDA_AUTHORITY_GROUP_L.IsFocused)
                    {
                        IDA_AUTHORITY_GROUP_L.Delete();
                    }
                }
            }
        }

        #endregion;

        private void EAPF0107_Load(object sender, EventArgs e)
        {
            IDA_AUTHORITY_GROUP_H.FillSchema();
        }
    }
}