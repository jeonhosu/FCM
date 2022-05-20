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

namespace FCMF0110
{
    public partial class FCMF0110_TL : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public FCMF0110_TL(Form pMainForm, ISAppInterface pAppInterface,
                            object pACCOUNT_SET_ID, object pACCOUNT_SET_CODE, object pACCOUNT_SET_NAME, object pACCOUNT_LEVEL)
        {
            InitializeComponent();

            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            ACCOUNT_SET_ID.EditValue = pACCOUNT_SET_ID;
            ACCOUNT_SET_CODE.EditValue = pACCOUNT_SET_CODE;
            ACCOUNT_SET_NAME.EditValue = pACCOUNT_SET_NAME;
            ACCOUNT_LEVEL.EditValue = pACCOUNT_LEVEL;
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

        private void Init_Account_Group_Insert()
        {
            Init_Account_Group_Column();
                        
            IGR_ACCOUNT_GROUP_TL.CurrentCellMoveTo(IGR_ACCOUNT_GROUP_TL.GetColumnToIndex("LANG_DESCRIPTION"));
            IGR_ACCOUNT_GROUP_TL.Focus();
        }

        private void Init_Account_Group_Column()
        {
            int mStart_Column;

            int mMin_Column;
            int mMax_Column = 22;
            mMin_Column = (Convert.ToInt32(ACCOUNT_LEVEL.EditValue)) * 2;

            for (mStart_Column = 2; mStart_Column < mMin_Column; mStart_Column++)
            {
                IGR_ACCOUNT_CONTROL_DFV.GridAdvExColElement[mStart_Column].Visible = (int)1;
            }

            for (mStart_Column = mMin_Column; mStart_Column < mMax_Column; mStart_Column++)
            {
                IGR_ACCOUNT_CONTROL_DFV.GridAdvExColElement[mStart_Column].Visible = (int)0;
                for (int R = 0; R < IGR_ACCOUNT_CONTROL_DFV.RowCount; R++)
                {
                    if (iString.ISNull(IGR_ACCOUNT_CONTROL_DFV.GetCellValue(R, mStart_Column)) != string.Empty)
                    {
                        IGR_ACCOUNT_CONTROL_DFV.SetCellValue(R, mStart_Column, String.Empty);
                    }
                }
            }
            IGR_ACCOUNT_CONTROL_DFV.ResetDraw = true;

            mMax_Column = 12;
            mMin_Column = (Convert.ToInt32(ACCOUNT_LEVEL.EditValue)) + 1;
            for (mStart_Column = 2; mStart_Column < mMin_Column; mStart_Column++)
            {
                IGR_ACCOUNT_GROUP_TL.GridAdvExColElement[mStart_Column].Visible = (int)1;
            }

            for (mStart_Column = mMin_Column; mStart_Column < mMax_Column; mStart_Column++)
            {
                IGR_ACCOUNT_GROUP_TL.GridAdvExColElement[mStart_Column].Visible = (int)0;
                for (int R = 0; R < IGR_ACCOUNT_GROUP_TL.RowCount; R++)
                {
                    if (iString.ISNull(IGR_ACCOUNT_CONTROL_DFV.GetCellValue(R, mStart_Column)) != string.Empty)
                    {
                        IGR_ACCOUNT_GROUP_TL.SetCellValue(R, mStart_Column, String.Empty);
                    }
                }
            }
            IGR_ACCOUNT_GROUP_TL.ResetDraw = true;
        }
        
        private void SEARCH_DB()
        {
            IDA_ACCOUNT_CONTROL_DFV.Fill();
            Init_Account_Group_Column();
        }

        #endregion
        
        #region ----- Application_MainButtonClick -----

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

        #endregion

        #region ----- Form Event ------

        private void FCMF0110_TL_Load(object sender, EventArgs e)
        {
            IDA_ACCOUNT_CONTROL_DFV.FillSchema();           
        }

        private void FCMF0110_TL_Shown(object sender, EventArgs e)
        {
            Init_Account_Group_Column();

            Application.UseWaitCursor =  false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void BTN_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ACCOUNT_CONTROL_TL.AddUnder();
            Init_Account_Group_Insert();
        }

        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ACCOUNT_CONTROL_DFV.Update();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            IDA_ACCOUNT_CONTROL_TL.Cancel();   
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion

        #region ----- Adapter Event -----
        
        private void IDA_ACCOUNT_CONTROL_TL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (ACCOUNT_LEVEL.EditValue == null)
            {// 계정세트레벨
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Account Level(계정 세트 레벨)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_DESC"]) == string.Empty)
            {// 계정명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Account Desc(계정명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (2 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT1_DESC"]) == string.Empty)
                {// 세그먼트명1
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment1 Desc(세그먼트1명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (3 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT2_DESC"]) == string.Empty)
                {// 세그먼트명2
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment2 Desc(세그먼트2명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (4 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT3_DESC"]) == string.Empty)
                {// 세그먼트명3
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment3 Desc(세그먼트3명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (5 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT4_DESC"]) == string.Empty)
                {// 세그먼트명4
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment4 Desc(세그먼트4명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (6 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT5_DESC"]) == string.Empty)
                {// 세그먼트명5
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment5 Desc(세그먼트5명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (7 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT6_DESC"]) == string.Empty)
                {// 세그먼트명6
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment6 Desc(세그먼트6명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (8 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT7_DESC"]) == string.Empty)
                {// 세그먼트명7
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment7 Desc(세그먼트7명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (9 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT8_DESC"]) == string.Empty)
                {// 세그먼트명8
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment8 Desc(세그먼트8명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (10 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT9_DESC"]) == string.Empty)
                {// 세그먼트명9
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment9 Desc(세그먼트9명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (11 <= Convert.ToInt32(ACCOUNT_LEVEL.EditValue))
            {
                if (iString.ISNull(e.Row["SEGMENT10_DESC"]) == string.Empty)
                {// 세그먼트명10
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Segment10 Desc(세그먼트10명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        #endregion
        
        #region ----- Lookup Event -----
        
        private void ILA_LANG_GROUP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_LANG.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ILD_LANG.SetLookupParamValue("W_LOOKUP_TYPE", "SYSTEM_TERRITORY");
        }

        private void ILA_LANG_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_LANG.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ILD_LANG.SetLookupParamValue("W_LOOKUP_TYPE", "SYSTEM_TERRITORY");
        }

        #endregion


    }
}