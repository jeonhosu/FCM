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

namespace FCMF0112
{
    public partial class FCMF0112 : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public FCMF0112()
        {
            InitializeComponent();
        }

        public FCMF0112(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void Search_DB()
        {
            idaDOCUMENT_HEADER.Fill();
            igrDOCUMENT_HEADER.Focus();
        }

        private void SetCommonParameter(string pGroup_Codee, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Codee);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Insert_Document_Header()
        {
            igrDOCUMENT_HEADER.SetCellValue("ENABLED_FLAG", "Y");
            igrDOCUMENT_HEADER.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            igrDOCUMENT_HEADER.CurrentCellActivate(1);
        }

        private void Insert_Document_Line()
        {
            igrDOCUMENT_LINE.SetCellValue("CONTRA_ACCOUNT_YN", "N");

            igrDOCUMENT_LINE.CurrentCellMoveTo(igrDOCUMENT_LINE.GetColumnToIndex("DOCUMENT_LINE_TYPE"));
            igrDOCUMENT_LINE.CurrentCellActivate(igrDOCUMENT_LINE.GetColumnToIndex("DOCUMENT_LINE_TYPE"));
            igrDOCUMENT_LINE.Focus();
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
                    if (idaDOCUMENT_HEADER.IsFocused)
                    {
                        idaDOCUMENT_HEADER.AddOver();
                        Insert_Document_Header();
                    }
                    else if (idaDOCUMENT_LINE.IsFocused)
                    {
                        idaDOCUMENT_LINE.AddOver();
                        Insert_Document_Line();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaDOCUMENT_HEADER.IsFocused)
                    {
                        idaDOCUMENT_HEADER.AddUnder();
                        Insert_Document_Header();
                    }
                    else if (idaDOCUMENT_LINE.IsFocused)
                    {
                        idaDOCUMENT_LINE.AddUnder();
                        Insert_Document_Line();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {                    
                    idaDOCUMENT_HEADER.Update();                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaDOCUMENT_HEADER.IsFocused)
                    {
                        idaDOCUMENT_LINE.Cancel();
                        idaDOCUMENT_HEADER.Cancel();
                    }
                    else if (idaDOCUMENT_LINE.IsFocused)
                    {
                        idaDOCUMENT_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaDOCUMENT_HEADER.IsFocused)
                    {
                        for (int i = 0; i < igrDOCUMENT_LINE.RowCount; i++)
                        {
                            idaDOCUMENT_LINE.CurrentRows[i].Delete();
                        }
                        idaDOCUMENT_HEADER.Delete();
                    }
                    else if (idaDOCUMENT_LINE.IsFocused)
                    {
                        idaDOCUMENT_LINE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ---- Form Event -----
        private void FCMF0112_Load(object sender, EventArgs e)
        {
            idaDOCUMENT_HEADER.FillSchema();

        }
        #endregion

        #region ----- Lookup Event -----

        private void ilaSLIP_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("SLIP_TYPE", "N");
        }

        private void ilaJOB_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildJOB_CODE_0.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaSLIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("SLIP_TYPE", "Y");
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ilaACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_CONTROL_I_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ildACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_DR_CR_I_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaDOCUMENT_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["SLIP_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10116"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["JOB_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10155"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["JOB_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10155"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void idaDOCUMENT_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:= Data(해당 데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDOCUMENT_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DOCUMENT_LINE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10163"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CONTROL_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10123"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_DR_CR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10122"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LINE_SEQ"]) == string.Empty || iString.ISNull(e.Row["LINE_SEQ"]) == "0")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10415"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}