using System;
using System.Data;
using System.Text;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;

using InfoSummit.Win.ControlAdv;

namespace FCMF0391
{
    public partial class Zebra_Position : Office2007Form
    {
        #region ----- Variables -----

        private string mMessageInfo = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public Zebra_Position(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- View Set Value Method ----

        private void ViewSetValue()
        {
        }

        #endregion;

        #region ----- Private Method ----


        #endregion;

        #region ----- Button Event -----

        private void btnSave_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                //idaZebra.SetInsertParamValue("P_BARCODE_PRINT_SET_CODE", "FC_ZEBRA");
                //idaZebra.SetUpdateParamValue("P_BARCODE_PRINT_SET_CODE", "FC_ZEBRA");
                idaZebra.Update();

                mMessageInfo = string.Format("Save OK, Label Printer Positon Set");
                isAppInterfaceAdv1.OnAppMessage(mMessageInfo);
                System.Windows.Forms.Application.DoEvents();

                MessageBoxAdv.Show(mMessageInfo, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Close();
            }
            catch (System.Exception ex)
            {
                mMessageInfo = ex.Message;
                isAppInterfaceAdv1.OnAppMessage(mMessageInfo);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void btnClose_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion;

        #region ----- FORM Event -----

        private void Zebra_Position_Load(object sender, EventArgs e)
        {
            try
            {
                //idaZebra.SetSelectParamValue("W_BARCODE_PRINT_SET_CODE", "FC_ZEBRA");
                idaZebra.Fill();
            }
            catch (System.Exception ex)
            {
                mMessageInfo = ex.Message;
                isAppInterfaceAdv1.OnAppMessage(mMessageInfo);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Zebra_Position_Shown(object sender, EventArgs e)
        {
            ViewSetValue();
        }

        #endregion;

        #region ----- AppMainButton Event -----

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
                    if (idaZebra.IsFocused)
                    {
                        idaZebra.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
            }
        }

        #endregion;
    }
}