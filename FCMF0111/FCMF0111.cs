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

namespace FCMF0111
{
    public partial class FCMF0111 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public FCMF0111(Form pMainFom, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainFom;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Method -----

        private void SEARCH_DB()
        {
            IDA_DEPT.Fill();
            igrDEPT.Focus();
        }

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void Default_Set_Value()
        {
            idcDV_ACCOUNT_BOOK.ExecuteNonQuery();
            SOB_NAME_0.EditValue = idcDV_ACCOUNT_BOOK.GetCommandParamValue("O_SOB_DESCRIPTION");
            DEPT_LEVEL_0.EditValue = idcDV_ACCOUNT_BOOK.GetCommandParamValue("O_DEPT_LEVEL");
        }

        private void Insert_Department()
        {
            igrDEPT.SetCellValue("ENABLED_FLAG", "Y");
            igrDEPT.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
        }
        #endregion

        #region ----- main Button Click ------
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
                    if (IDA_DEPT.IsFocused)
                    {
                        IDA_DEPT.AddOver();
                        Insert_Department();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_DEPT.IsFocused)
                    {
                        IDA_DEPT.AddUnder();
                        Insert_Department();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_DEPT.IsFocused)
                    {
                        IDA_DEPT.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DEPT.IsFocused)
                    {
                        IDA_DEPT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DEPT.IsFocused)
                    {
                        IDA_DEPT.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void FCMF0111_Load(object sender, EventArgs e)
        {
            IDA_DEPT.FillSchema();

            Default_Set_Value();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }
        #endregion

        #region ---- Adapter Event -----

        private void IDA_DEPT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {   
            if (iString.ISNull(e.Row["DEPT_CODE"]) == string.Empty)
            {// 부서코드
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10019"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEPT_NAME"]) == string.Empty)
            {// 부서명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEPT_LEVEL"]) == string.Empty)
            {// 부서 레벨
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10021"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["DEPT_LEVEL"]) > Convert.ToInt32(1))
            {// 부서 레벨이 0이 아닐경우 상위부서는 반드시 선택해야 합니다.
                if (iString.ISNull(e.Row["UPPER_DEPT_ID"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10132"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }

            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {// 시작일자 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_TO"]) != string.Empty)
            {// 종료일자 
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void IDA_DEPT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ---- Lookup Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaDEPT_UPPER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_UPPER.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        #endregion
    }
}