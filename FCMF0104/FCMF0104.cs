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

namespace FCMF0104
{ 
    public partial class FCMF0104 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public FCMF0104(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property Method ------
        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void SEARCH_DB()
        {
            if (itbBANK.SelectedTab.TabIndex == 1)
            {
                idaBANK_GROUP.Fill();
                igrBANK_GROUP.Focus();
            }
            else if (itbBANK.SelectedTab.TabIndex == 2)
            {
                idaBANK_SITE.Fill();
                BANK_GROUP_NAME.Focus();
            }
            
        }

        private void Insert_Bank_Group()
        {
            igrBANK_GROUP.SetCellValue("ENABLED_FLAG", "Y");
            igrBANK_GROUP.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISGetDate());
            igrBANK_GROUP.Focus();
        }

        private void Insert_Bank_Site()
        {
            ENABLED_FLAG.CheckBoxValue = "Y";
            START_DATE.EditValue = iDate.ISGetDate();
            EFFECTIVE_DATE_FR.EditValue = iDate.ISGetDate();

            BANK_GROUP_NAME.Focus();
        }

        private void Insert_Bank_Account()
        {
            igrBANK_ACCOUNT.SetCellValue("ENABLED_FLAG", "Y");
            igrBANK_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISGetDate());

            igrBANK_ACCOUNT.CurrentCellMoveTo(igrBANK_ACCOUNT.GetColumnToIndex("BANK_ACCOUNT_CODE"));
            igrBANK_ACCOUNT.Focus();
        }
        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Button Click -----
        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaBANK_GROUP.IsFocused)
                    {
                        idaBANK_GROUP.AddOver();
                        Insert_Bank_Group();
                    }
                    else if (idaBANK_SITE.IsFocused)
                    {
                        idaBANK_SITE.AddOver();
                        Insert_Bank_Site();
                    }
                    else if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.AddOver();
                        Insert_Bank_Account();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaBANK_GROUP.IsFocused)
                    {
                        idaBANK_GROUP.AddUnder();
                        Insert_Bank_Group();
                    }
                    else if (idaBANK_SITE.IsFocused)
                    {
                        idaBANK_SITE.AddUnder();
                        Insert_Bank_Site();
                    }
                    else if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.AddUnder();
                        Insert_Bank_Account();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaBANK_GROUP.IsFocused)
                    {
                        idaBANK_GROUP.Update();
                    }
                    else if (idaBANK_SITE.IsFocused || idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_SITE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBANK_GROUP.IsFocused)
                    {
                        idaBANK_GROUP.Cancel();
                    }
                    else if (idaBANK_SITE.IsFocused)
                    {
                        idaBANK_SITE.Cancel();
                    }
                    else if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBANK_GROUP.IsFocused)
                    {
                        idaBANK_GROUP.Delete();
                    }
                    else if (idaBANK_SITE.IsFocused)
                    {
                        idaBANK_SITE.Delete();
                    }
                    else if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.Delete();
                    }
                }
            }
        }
        #endregion
        
        #region ----- Form Event -----
        private void FCMF0104_Load(object sender, EventArgs e)
        {
            idaBANK_GROUP.FillSchema();
            idaBANK_SITE.FillSchema();
        }
        #endregion

        #region ----- Adapter Event -----

        private void idaBANK_SITE_ExcuteKeySearch(object pSender)
        {
            SEARCH_DB();
            BANK_CODE.Focus();
        }
        

        private void idaBANK_GROUP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BANK_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Group Code(은행 그룹코드)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BANK_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Group Name(은행 그룹명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BANK_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Type(은행구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaBANK_GROUP_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaBANK_SITE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BANK_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Code(은행코드)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BANK_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Name(은행명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }            
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_TO"] != DBNull.Value)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {// 시작일자 ~ 종료일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 기간 검증 오류
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaBANK_SITE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaBANK_ACCOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["BANK_ACCOUNT_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Account Name(은행 계좌명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BANK_ACCOUNT_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Account Number(은행 계좌번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            //if (iString.ISNull(e.Row["OWNER_NAME"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Owner Name(예금주)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["ACCOUNT_TYPE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Account Type(계좌 종류)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Currency Code(통화)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
            //    e.Cancel = true;
            //    return;
            //}
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_TO"] != DBNull.Value)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {// 시작일자 ~ 종료일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 기간 검증 오류
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaBANK_ACCOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ----- Lookup Code -----
        private void ilaBANK_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BANK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDC_METHOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DC_METHOD");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaACCOUNT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaCURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        #endregion



    }
}