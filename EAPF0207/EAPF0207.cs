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
using ISCommonUtil;

namespace EAPF0207
{
    public partial class EAPF0207 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime idate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public EAPF0207()
        {
            InitializeComponent();
        }

        public EAPF0207(System.Windows.Forms.Form pMainForm, InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchFromDataAdapter()
        {
            IDA_PAYMENT_TERM.Fill();
        }

        #endregion;

        #region -- Default Value Setting --

        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            IGR_PAYMENT_TERM_LIST.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            IGR_PAYMENT_TERM_LIST.SetCellValue("ENABLED_FLAG", "Y");

            IGR_PAYMENT_TERM_LIST.Focus();
        }

        #endregion

        #region ----- Events -----

        private void EAPF0207_Load(object sender, EventArgs e)
        {
            IDA_PAYMENT_TERM.FillSchema();                        
        }
         
        private void isAppInterfaceAdv1_AppMainButtonClick(InfoSummit.Win.ControlAdv.ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchFromDataAdapter();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_PAYMENT_TERM.IsFocused == true)
                    {
                        IDA_PAYMENT_TERM.AddOver();
                        GRID_DefaultValue();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_PAYMENT_TERM.IsFocused == true)
                    {
                        IDA_PAYMENT_TERM.AddUnder();
                        GRID_DefaultValue();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_PAYMENT_TERM.Update();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_PAYMENT_TERM.IsFocused == true)
                    {
                        IDA_PAYMENT_TERM.Cancel();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_PAYMENT_TERM.IsFocused == true)
                    {
                        IDA_PAYMENT_TERM.Delete();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }


        #endregion;

        #region ----- Lookup Event -----
        
        private void ILA__CUT_OFF_DAY_RefreshLookupData(object pSender, InfoSummit.Win.ControlAdv.ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DAY_NUM");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }        

        #endregion

        #region ----- Adapter Event -----

        private void IDA_PAYMENT_TERM_PreRowUpdate(InfoSummit.Win.ControlAdv.ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PAYMENT_TERM_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Payment Term Code"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DESCRIPTION"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Payment Term Desc"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CASH_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Cash Type"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SPECIFIED_DAY"]) != string.Empty && iConv.ISNumtoZero(e.Row["SPECIFIED_DAY"]) > iConv.ISNumtoZero(31))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10158"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Effective Date From"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
        
        private void isDataAdapter1_PreDelete(InfoSummit.Win.ControlAdv.ISPreDeleteEventArgs e)
        {            
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Master Data"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}