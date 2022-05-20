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

namespace FCMF0273
{
    public partial class FCMF0273 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iString = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0273(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            object vObject1 = GL_DATE_FR_0.EditValue;
            if (iString.ISNull(vObject1) == string.Empty)
            {
                //시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object vObject2 = GL_DATE_TO_0.EditValue;
            if (iString.ISNull(vObject2) == string.Empty)
            {
                //종료일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Convert.ToDateTime(GL_DATE_FR_0.EditValue) > Convert.ToDateTime(GL_DATE_TO_0.EditValue))
            {
                //종료일은 시작일 이후이어야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10345"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                GL_DATE_FR_0.Focus();
                return;
            }

            object vObject3 = MANAGEMENT_NAME_0.EditValue;
            if (iString.ISNull(vObject3) == string.Empty)
            {
                //관리항목은 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10417"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object vObject4 = ACCOUNT_CODE_FR_0.EditValue;
            object vObject5 = ACCOUNT_CODE_TO_0.EditValue;
            if (iString.ISNull(vObject4) == string.Empty || iString.ISNull(vObject5) == string.Empty)
            {
                //계정과목은 필수입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int vACCOUNT_CODE_FR_0 = ConvertInteger(vObject4);
            int vACCOUNT_CODE_TO_0 = ConvertInteger(vObject5);
            if (vACCOUNT_CODE_FR_0 > vACCOUNT_CODE_TO_0)
            {
                //종료계정은 시작계정 이후의 계정이어야 합니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10414"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ACCOUNT_CODE_FR_0.Focus();
                return;
            }

            if (TAB_MAIN.SelectedTab.TabIndex == 1)
            {
                idaLIST_ACCOUNT.Fill();
                igrLIST_ACCOUNT.Focus();
            }
            else if (TAB_MAIN.SelectedTab.TabIndex == 2)
            {
                IDA_ALL_MANAGEMENT_LEDGER.Fill();
                IGR_ALL_MANAGEMENT_LEDGER.Focus();
            }
        }

        #endregion;

        #region ----- Convert decimal  Method ----

        private int ConvertInteger(object pObject)
        {
            bool vIsConvert = false;
            int vConvertInteger = 0;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        string vString = pObject as string;
                        vConvertInteger = int.Parse(vString);
                    }
                }

            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertInteger;
        }

        #endregion;

        #region ----- MDi ToolBar Button Event -----

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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaLIST_ACCOUNT.IsFocused == true)
                    {
                        idaLIST_ACCOUNT.Cancel();
                    }
                    else if (idaUP_MANAGEMENT_LEDGER.IsFocused == true)
                    {
                        idaUP_MANAGEMENT_LEDGER.Cancel();
                    }
                    else if (idaDET_MANAGEMENT_LEDGER.IsFocused == true)
                    {
                        idaDET_MANAGEMENT_LEDGER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0273_Load(object sender, EventArgs e)
        {
            
        }
        
        private void FCMF0273_Shown(object sender, EventArgs e)
        {
            GL_DATE_FR_0.EditValue = System.DateTime.Today;
            GL_DATE_TO_0.EditValue = System.DateTime.Today;
        }

        private void TAB_MAIN_Click(object sender, EventArgs e)
        {
            SearchDB(); 
        }
        #endregion

        #region ----- Adapter Event -----

        #endregion

        #region ----- Grid Event -----

        private void igrDET_CUSTOMER_LEDGER_CellDoubleClick(object pSender)
        {
            if (igrDET_MANAGEMENT_LEDGER.RowIndex > -1)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(igrDET_MANAGEMENT_LEDGER.GetCellValue("SLIP_HEADER_ID"));
                if (vSLIP_HEADER_ID > Convert.ToInt32(0))
                {
                    System.Windows.Forms.Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    FCMF0205.FCMF0205 vFCMF0205 = new FCMF0205.FCMF0205(this.MdiParent, isAppInterfaceAdv1.AppInterface, vSLIP_HEADER_ID);
                    vFCMF0205.Show();

                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.UseWaitCursor = false;
                }
            }
        }

        private void IGR_ALL_MANAGEMENT_LEDGER_CellDoubleClick(object pSender)
        {
            if (IGR_ALL_MANAGEMENT_LEDGER.RowIndex > -1)
            {
                int vSLIP_HEADER_ID = iString.ISNumtoZero(IGR_ALL_MANAGEMENT_LEDGER.GetCellValue("SLIP_HEADER_ID"));
                if (vSLIP_HEADER_ID > Convert.ToInt32(0))
                {
                    System.Windows.Forms.Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    FCMF0205.FCMF0205 vFCMF0205 = new FCMF0205.FCMF0205(this.MdiParent, isAppInterfaceAdv1.AppInterface, vSLIP_HEADER_ID);
                    vFCMF0205.Show();

                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.UseWaitCursor = false;
                }
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaACCOUNT_CODE_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL_0.SetLookupParamValue("W_ACCOUNT_SET_ID", null);
            ildACCOUNT_CONTROL_0.SetLookupParamValue("W_ACCOUNT_CODE", null);
        }

        private void ilaACCOUNT_CODE_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL_0.SetLookupParamValue("W_ACCOUNT_SET_ID", null);
            ildACCOUNT_CONTROL_0.SetLookupParamValue("W_ACCOUNT_CODE", ACCOUNT_CODE_FR_0.EditValue);
        }

        private void ilaMANAGEMENT_0_SelectedRowData(object pSender)
        {
            LIST_MANAGEMENT_GUBUN_CODE.EditValue = null;
            LIST_MANAGEMENT_GUBUN_NAME.EditValue = null;
        }

        #endregion


    }
}