using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;

using InfoSummit.Win.ControlAdv;

namespace FCMF0764
{
    public partial class FCMF0764 : Office2007Form
    {
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();
        private ISCommonUtil.ISFunction.ISConvert iString = new ISCommonUtil.ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public FCMF0764()
        {
            InitializeComponent();
        }

        public FCMF0764(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    int vIndexTab = isTAB.SelectedIndex;

                    if (vIndexTab == 0)
                    {
                        Search_LIST_FORWARD_AMT_MST();
                    }
                    else if (vIndexTab == 1)
                    {
                        Search_LIST_FORWARD_AMT_DET();
                    }
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
                    int vIndexTab = isTAB.SelectedIndex;

                    if (vIndexTab == 0)
                    {
                        if (idaLIST_FORWARD_AMT_MST.IsFocused)
                        {
                            idaLIST_FORWARD_AMT_MST.Cancel();
                        }
                    }
                    else if (vIndexTab == 1)
                    {
                        if (idaLIST_FORWARD_AMT_DET.IsFocused)
                        {
                            idaLIST_FORWARD_AMT_DET.Cancel();
                        }
                    }
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

        #region ----- Form Event -----

        private void FCMF0764_Load(object sender, EventArgs e)
        {
            FORWARD_YEAR_0.EditValue = iDate.ISYear(System.DateTime.Today);
        }

        private void FCMF0764_Shown(object sender, EventArgs e)
        {

        }

        #endregion

        #region ----- Methods ----

        private bool Check_Blank_Year()
        {
            bool vIsOK = true;

            object vObject1 = FORWARD_YEAR_0.EditValue;
            if (iString.ISNull(vObject1) == string.Empty)
            {
                vIsOK = false;

                //년도는 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return vIsOK;
        }

        private void Search_LIST_FORWARD_AMT_MST()
        {
            try
            {
                if (Check_Blank_Year() == true)
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    isTAB.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();

                    idaLIST_FORWARD_AMT_MST.Fill();

                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    isTAB.Cursor = System.Windows.Forms.Cursors.Default;
                    tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                    tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                    igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                    igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    FORWARD_YEAR_0.Focus();
                }
            }
            catch (System.Exception ex)
            {
                MessageBoxAdv.Show(ex.Message);
            }
        }

        private void Search_LIST_FORWARD_AMT_DET()
        {
            try
            {
                if (Check_Blank_Year() == true)
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    isTAB.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();

                    idaLIST_FORWARD_AMT_DET.Fill();

                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    isTAB.Cursor = System.Windows.Forms.Cursors.Default;
                    tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                    tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                    igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                    igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    FORWARD_YEAR_0.Focus();
                }
            }
            catch (System.Exception ex)
            {
                MessageBoxAdv.Show(ex.Message);
            }
        }

        #endregion;

        #region ----- Button Event -----

        private void bCREATE_FORWARD_AMT_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            System.Windows.Forms.DialogResult ChoiceValue;

            if (Check_Blank_Year() == false)
            {
                return;
            }

            //차기이월을 위한 기초자료를 생성하시겠습니까?[FCM_10434]
            ChoiceValue = MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10434"), "Question", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxDefaultButton.Button2);
            if (ChoiceValue == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            try
            {
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                isTAB.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                idcCREATE_FORWARD_AMT.ExecuteNonQuery();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                isTAB.Cursor = System.Windows.Forms.Cursors.Default;
                tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();

                object vObject_O_MESSAGE = idcCREATE_FORWARD_AMT.GetCommandParamValue("O_MESSAGE");
                string vO_MESSAGE = iString.ISNull(vObject_O_MESSAGE);
                MessageBoxAdv.Show(vO_MESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void bCREATE_CRJ_SLIP_1_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            System.Windows.Forms.DialogResult ChoiceValue;

            if (Check_Blank_Year() == false)
            {
                return;
            }

            //차기이월 작업을 실행 하시겠습니까?[FCM_10435]
            ChoiceValue = MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10435"), "Question", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxDefaultButton.Button2);
            if (ChoiceValue == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            try
            {
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                isTAB.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                idcCREATE_FORWARD_SLIP.ExecuteNonQuery();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                isTAB.Cursor = System.Windows.Forms.Cursors.Default;
                tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();

                object vObject_O_MESSAGE = idcCREATE_FORWARD_SLIP.GetCommandParamValue("O_MESSAGE");
                string vO_MESSAGE = iString.ISNull(vObject_O_MESSAGE);
                MessageBoxAdv.Show(vO_MESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void bDELETE_CRJ_SLIP_1_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            System.Windows.Forms.DialogResult ChoiceValue;

            if (Check_Blank_Year() == false)
            {
                return;
            }

            //차기이월된 자료를 삭제 하시겠습니까?[FCM_10436]
            ChoiceValue = MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10436"), "Question", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxDefaultButton.Button2);
            if (ChoiceValue == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            try
            {
                int vCountRow = igrLIST_FORWARD_AMT_DET.RowCount;
                if (vCountRow < 1)
                {
                    //2번째 탭(계정별관리항목별이월금액)에서 자료를 조회후 작업 바랍니다.[FCM_10437]
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10437"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                isTAB.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                idcDELETE_FORWARD_SLIP.ExecuteNonQuery();

                this.Cursor = System.Windows.Forms.Cursors.Default;
                isTAB.Cursor = System.Windows.Forms.Cursors.Default;
                tabLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                tabLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                igrLIST_FORWARD_AMT_MST.Cursor = System.Windows.Forms.Cursors.Default;
                igrLIST_FORWARD_AMT_DET.Cursor = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();

                object vObject_O_MESSAGE = idcDELETE_FORWARD_SLIP.GetCommandParamValue("O_MESSAGE");
                string vO_MESSAGE = iString.ISNull(vObject_O_MESSAGE);
                MessageBoxAdv.Show(vO_MESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Lookup Event -----


        #endregion;
    }
}