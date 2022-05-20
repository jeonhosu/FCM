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

namespace FCMF0120
{
    public partial class FCMF0120 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public FCMF0120(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property / Method -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorporation()
        {

        }

        private void Init_Fiscal_Period_Insert()
        {
            igrFISCAL_PERIOD.SetCellValue("PERIOD_STATUS", DV_PERIOD_STATUS.EditValue);
            igrFISCAL_PERIOD.SetCellValue("PERIOD_STATUS_NAME", DV_PERIOD_STATUS_NAME.EditValue);
        }

        private void Init_Fiscal_Year_Insert()
        {
            FISCAL_COUNT.EditValue = 0;
            FISCAL_YEAR.Focus();
        }

        private void SEARCH_DB()
        {
            idaFISCAL_CALENDAR.Fill();
            FISCAL_CALENDAR_ID.Focus();
        }

        #endregion
         
        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- Application_MainButtonClick -----
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
                    if (idaFISCAL_CALENDAR.IsFocused)
                    {
                        idaFISCAL_CALENDAR.AddOver();
                        FISCAL_CALENDAR_ID.EditValue = 0;
                        FISCAL_CALENDAR_ID.Focus();
                    }
                    else if (idaFISCAL_YEAR.IsFocused)
                    {
                        idaFISCAL_YEAR.AddOver();
                        Init_Fiscal_Year_Insert();
                    }
                    else if (idaFISCAL_PERIOD.IsFocused)
                    {
                        idaFISCAL_PERIOD.AddOver();
                        Init_Fiscal_Period_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaFISCAL_CALENDAR.IsFocused)
                    {
                        idaFISCAL_CALENDAR.AddUnder();
                        FISCAL_CALENDAR_ID.EditValue = 0;
                        FISCAL_CALENDAR_ID.Focus();
                    }
                    else if (idaFISCAL_YEAR.IsFocused)
                    {
                        idaFISCAL_YEAR.AddUnder();
                        Init_Fiscal_Year_Insert();
                    }
                    else if (idaFISCAL_PERIOD.IsFocused)
                    {
                        idaFISCAL_PERIOD.AddUnder();
                        Init_Fiscal_Period_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaFISCAL_CALENDAR.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaFISCAL_CALENDAR.IsFocused)
                    {
                        idaFISCAL_CALENDAR.Cancel();
                    }
                    else if (idaFISCAL_YEAR.IsFocused)
                    {
                        idaFISCAL_YEAR.Cancel();
                    }
                    else if (idaFISCAL_PERIOD.IsFocused)
                    {
                        idaFISCAL_PERIOD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaFISCAL_CALENDAR.IsFocused)
                    {
                        idaFISCAL_CALENDAR.Delete();
                    }
                    else if (idaFISCAL_YEAR.IsFocused)
                    {
                        idaFISCAL_YEAR.Delete();
                    }
                    else if (idaFISCAL_PERIOD.IsFocused)
                    {
                        idaFISCAL_PERIOD.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void FCMF0120_Load(object sender, EventArgs e)
        {
            idaFISCAL_CALENDAR.FillSchema();
            idaFISCAL_YEAR.FillSchema();
            idaFISCAL_PERIOD.FillSchema();
            idcDV_PERIOD_STATUS.ExecuteNonQuery();

            BTN_PREVIOUS.BringToFront();
            BTN_NEXT.BringToFront();
        }

        private void btnPERIOD_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(FISCAL_CALENDAR_ID.EditValue) == string.Empty)
            {// 회계달력ID
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(FISCAL_CALENDAR_ID))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                return;
            }

            DialogResult dlgResult;
            FCMF0120_YEAR vFCMF0120_YEAR = new FCMF0120_YEAR(isAppInterfaceAdv1.AppInterface, 
                                                            FISCAL_CALENDAR_ID.EditValue, 
                                                            FISCAL_CALENDAR_CODE.EditValue, 
                                                            FISCAL_CALENDAR_NAME.EditValue);
            dlgResult = vFCMF0120_YEAR.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                SEARCH_DB();
            }
            vFCMF0120_YEAR.Dispose();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents(); 
        }

        private void FISCAL_YEAR_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            START_DATE.EditValue = iDate.ISGetDate(string.Format("{0}-01-01", FISCAL_YEAR.EditValue));
            END_DATE.EditValue = iDate.ISGetDate(string.Format("{0}-12-31", FISCAL_YEAR.EditValue));
        }

        private void BTN_PREVIOUS_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaFISCAL_YEAR.MoveNext(FISCAL_YEAR.Name);
        }

        private void BTN_NEXT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaFISCAL_YEAR.MovePrevious(FISCAL_YEAR.Name);
        } 

        #endregion

        #region ----- Adapter Event -----

        private void idaFISCAL_CALENDAR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["FISCAL_CALENDAR_ID"]) == string.Empty)
            {// 회계달력ID
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(FISCAL_CALENDAR_ID))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return; 
            }
            if (iString.ISNull(e.Row["FISCAL_CALENDAR_NAME"]) == string.Empty)
            {// 회계달력명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(FISCAL_CALENDAR_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return; 
            }
        }

        private void idaFISCAL_CALENDAR_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaFISCAL_YEAR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {            
            if (iString.ISDecimaltoZero(e.Row["FISCAL_COUNT"],0) == 0)
            {// 회계기수
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10079"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FISCAL_YEAR"]) == string.Empty)
            {// 회계년도
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(FISCAL_YEAR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["START_DATE"]) == string.Empty)
            {// 회계년도 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(START_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["END_DATE"]) == string.Empty)
            {// 회계년도 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", string.Format("&&VALUE:={0}", Get_Edit_Prompt(END_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Convert.ToDateTime(e.Row["START_DATE"]) > Convert.ToDateTime(e.Row["END_DATE"]))
            {// 회계년도 기간설정
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["YEAR_STATUS"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(YEAR_STATUS_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaFISCAL_YEAR_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaFISCAL_PERIOD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERIOD_NAME"]) == string.Empty)
            {// 회계기간명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Period Name(기간명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["TEMP_PERIOD_STATUS"]) == string.Empty)
            {// 기간 상태
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Temp. Period Status(임시전표 기간상태)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PERIOD_STATUS"]) == string.Empty)
            {// 기간 상태
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Period Status(기간상태)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CLOSING_PERIOD_STATUS"]) == string.Empty)
            {// 기간 상태
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Closing Period Status(결산전표 기간상태)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["START_DATE"]) == string.Empty)
            {// 회계기간 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Period Start Date(회계기간 시작일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["END_DATE"]) == string.Empty)
            {// 회계기간 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Period End Date(회계기간 종료일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Convert.ToDateTime(e.Row["START_DATE"]) > Convert.ToDateTime(e.Row["END_DATE"]))
            {// 회계기간 기간설정
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["QUARTER_NUM"],0) == 0)
            {// 분기
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Quarter Num(분기)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["HALF_NUM"], 0) == 0)
            {// 반기
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Half Num(반기)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaFISCAL_PERIOD_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaFISCAL_YEAR_ExcuteKeySearch(object pSender)
        {
            SEARCH_DB();
        }

        private void idaFISCAL_CALENDAR_ExcuteKeySearch(object pSender)
        {
            SEARCH_DB();
        }

        private void igrFISCAL_PERIOD_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == igrFISCAL_PERIOD.GetColumnToIndex("PERIOD_NAME"))
            {// 회계기간 입력.

            }
        }
        #endregion

        #region ----- Lookup Event -----

        #endregion

    }
}