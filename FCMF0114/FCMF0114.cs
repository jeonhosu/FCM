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


/*
 * 
 * Project      : (주)그린테크놀로지 ERP
 * Module       : Financial(회계관리)
 * Program Name : FCMF0114
 * Description  : 재무제표보고서양식설정관리
 *
 * relevant program  : 
 * 
 * Program History :
 * 
 ------------------------------------------------------------------------------
   Date         Worker                  Description
------------------------------------------------------------------------------
 * 2015-04-06   Im Dong Eon(임동언)     최초 생성
 * 
 * 
 * 
 */


namespace FCMF0114
{
    public partial class FCMF0114 : Office2007Form
    {
        #region ----- Variables -----

            private ISFunction.ISConvert iString = new ISFunction.ISConvert();
            private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

            object v_Search_yn = 'Y';  //update전 해지관련 정보 체크 제어

        #endregion;


        #region ----- Constructor -----

        public FCMF0114(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

            private void Search()
            {

                //보고서양식은 필수사항입니다.
                if (iString.ISNull(W_FS_TYPE_CD.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10156"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_FS_TYPE.Focus();
                    return;
                }

                //조회조건의 필요한 값이 모두 채워졌을 경우 조건에 부합되는 자료를 조회한다.
                if (tabMain.SelectedIndex == 0)
                {
                    idaLIST_1tabMaster.Fill();
                    igrLIST_1tabMaster.Focus();
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
                    if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)        //검색
                    {
                        Search();
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)  //위에 새레코드 추가
                    {
                        if (idaLIST_1tabMaster.IsFocused)
                        {
                            if (iString.ISNull(W_FS_LEVEL.EditValue) == string.Empty)
                            {
                                //작업할 보고서양식을 선택 후 작업바랍니다.
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10546"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                W_FS_TYPE.Focus();
                                return;
                            }
                            else
                            {

                                idaLIST_1tabMaster.AddOver();

                                FS_TYPE.EditValue = W_FS_TYPE_CD.EditValue;

                                //합계잔액시산표 : 1005, 제조원가명세서 : 1004, 손익계산서 : 1003, 재무상태표 : 1002
                                if (iString.ISNull(W_FS_TYPE_CD.EditValue) == "1005" || iString.ISNull(W_FS_TYPE_CD.EditValue) == "1004" || iString.ISNull(W_FS_TYPE_CD.EditValue) == "1003")
                                {
                                    //인쇄위치
                                    FS_PRT_POS.EditValue = "04";
                                    FS_PRT_POS_NM.EditValue = "무관";
                                }

                                //출력구분
                                FS_RET_GB.EditValue = "03";
                                FS_RET_GB_NM.EditValue = "계정 + 세목";

                                LAST_LEVEL_YN.CheckBoxValue = "N";  //최하위여부
                                VIEW_YN.CheckBoxValue = "Y";        //출력여부
                                FORM_FRAME_YN.CheckBoxValue = "N";  //보고서틀유지여부

                                ITEM_CODE.Focus();
                            }

                        }
                        else if (idaLIST_1tabDetail.IsFocused)
                        {
                            idaLIST_1tabDetail.AddOver();
                        }
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder) //아래에 새레코드 추가
                    {
                        if (idaLIST_1tabMaster.IsFocused)
                        {
                            if (iString.ISNull(W_FS_LEVEL.EditValue) == string.Empty)
                            {
                                //작업할 보고서양식을 선택 후 작업바랍니다.
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10546"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                W_FS_TYPE.Focus();
                                return;
                            }
                            else
                            {

                                idaLIST_1tabMaster.AddUnder();

                                FS_TYPE.EditValue = W_FS_TYPE_CD.EditValue;

                                if (iString.ISNull(W_FS_TYPE_CD.EditValue) == "1005" || iString.ISNull(W_FS_TYPE_CD.EditValue) == "1004" || iString.ISNull(W_FS_TYPE_CD.EditValue) == "1003")
                                {
                                    FS_PRT_POS.EditValue = "04";
                                    FS_PRT_POS_NM.EditValue = "무관";
                                }

                                //출력구분
                                FS_RET_GB.EditValue = "03";
                                FS_RET_GB_NM.EditValue = "계정 + 세목";

                                LAST_LEVEL_YN.CheckBoxValue = "N";  //최하위여부
                                VIEW_YN.CheckBoxValue = "Y";        //출력여부
                                FORM_FRAME_YN.CheckBoxValue = "N";  //보고서틀유지여부

                                ITEM_CODE.Focus();
                            }

                        }
                        else if (idaLIST_1tabDetail.IsFocused)
                        {
                            idaLIST_1tabDetail.AddUnder();
                        }
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)   //저장
                    {

                        if (idaLIST_1tabDetail.AddedRowCount > 0 || idaLIST_1tabDetail.ModifiedRowCount > 0)
                        {
                            idaLIST_1tabDetail.Update();
                        }

                        //저장 전 신규로 추가한 자료가 있는지를 파악하여 있다면 저장 후 재조회하기 위함이다.
                        int Add_row_cnt = idaLIST_1tabMaster.AddedRowCount;

                        idaLIST_1tabMaster.Update();

                        if (Add_row_cnt > 0 && iString.ISNull(v_Search_yn).Equals("Y"))
                        {
                            Search();
                        }

                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)   //취소
                    {
                        if (tabMain.SelectedIndex == 0)
                        {
                            //M-D구조에서 취소기능 구현 시 디테일이 위에 있어야 M-D구조에서 디테일에 자료가 존재한다는 메세지를 게거할 수 있다.
                            idaLIST_1tabDetail.Cancel();
                            idaLIST_1tabMaster.Cancel();
                        }
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)   //삭제
                    {

                        //선택한 자료를 삭제하시겠습니까?
                        DialogResult mChoiceValue;

                        mChoiceValue = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

                        if (idaLIST_1tabMaster.IsFocused)
                        {
                            if (mChoiceValue == DialogResult.No)
                            {   //삭제취소
                                igrLIST_1tabMaster.Focus();
                                return;
                            }
                            else
                            {

                                for (int r = 0; r < igrLIST_1tabDetail.RowCount; r++)
                                {

                                    idaLIST_1tabDetail.Delete();
                                }

                                idaLIST_1tabMaster.Delete();
                             
                                idaLIST_1tabMaster.Update();
                                idaLIST_1tabDetail.Update();

                                Search();
                            }

                        }
                        else if (idaLIST_1tabDetail.IsFocused)
                        {
                            if (mChoiceValue == DialogResult.No)
                            {   //삭제취소
                                igrLIST_1tabDetail.Focus();
                                return;
                            }
                            else
                            {

                                idaLIST_1tabDetail.Delete();
                             
                                idaLIST_1tabDetail.Update();

                                Search();
                            }
                        }

                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)    //인쇄
                    {
                        //XLPrinting("PRINT");
                    }
                    else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)   //엑셀
                    {
                        //XLPrinting("FILE");
                    }
                }
            }

        #endregion;

        #region ----- Form Event -----

            private void FCMF0114_Load(object sender, EventArgs e)
            {
                idaLIST_1tabMaster.FillSchema();
                idaLIST_1tabDetail.FillSchema();
            }

        #endregion



        #region ----- Grid Event -----



        #endregion




        #region ----- Adapter Lookup Event -----

            private void ilaFS_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
            {
                ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FS_TYPE");
                ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "N");
            }

            private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
            {
                ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_DR_CR");
                ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
            }

            private void ilaFS_PRT_POS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
            {
                ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FS_PRT_POS");
                ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
            }

            private void ilaCALC_SIGN_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
            {
                ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "CALC_SIGN");
                ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
            }

            private void ilaFS_RET_GB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
            {
                ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FS_RET_GB");
                ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "N");
            }

            private void ilaFS_TAR_DATA_GB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
            {
                ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FS_TAR_DATA_GB");
                ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "N");
            }


            private void idaLIST_1tabMaster_PreRowUpdate(ISPreRowUpdateEventArgs e)
            {

                //항목코드는 필수입니다.
                if (iString.ISNull(e.Row["ITEM_CODE"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10157"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    ITEM_CODE.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //항목명은 필수입니다.
                if (iString.ISNull(e.Row["ITEM_NAME"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10158"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    ITEM_NAME.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //출력명은 필수입니다.
                if (iString.ISNull(e.Row["VIEW_NAME"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10547"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    VIEW_NAME.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //정렬순서는 필수입니다.
                if (iString.ISNull(e.Row["SORT_SEQ"]) == string.Empty || iString.ISNull(e.Row["SORT_SEQ"]) == "0")
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10539"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    SORT_SEQ.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //계산레벨은 필수입니다.
                if (iString.ISNull(e.Row["FS_LEVEL"]) == string.Empty || iString.ISNull(e.Row["FS_LEVEL_NM"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10540"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    FS_LEVEL_NM.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //인쇄위치는 필수입니다.
                if (iString.ISNull(e.Row["FS_PRT_POS"]) == string.Empty || iString.ISNull(e.Row["FS_PRT_POS_NM"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10541"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    FS_PRT_POS_NM.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }


                //출력구분은 필수입니다.
                if (iString.ISNull(e.Row["FS_RET_GB"]) == string.Empty || iString.ISNull(e.Row["FS_RET_GB_NM"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10550"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    FS_RET_GB_NM.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }


                //update전 정합성을 통과했을 경우는 정상적 처리를 위해 아래 변수값을 제어한다.
                v_Search_yn = 'Y';

            }

            private void idaLIST_1tabDetail_PreRowUpdate(ISPreRowUpdateEventArgs e)
            {
                //항목명은 필수입니다.
                if (iString.ISNull(e.Row["ITEM_NAME"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10158"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    igrLIST_1tabDetail.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //연산부호는 필수입니다.
                if (iString.ISNull(e.Row["CALC_SIGN_NM"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10543"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    igrLIST_1tabDetail.Focus();
                    e.Cancel = true;
                    v_Search_yn = 'N';
                    return;
                }

                //update전 정합성을 통과했을 경우는 정상적 처리를 위해 아래 변수값을 제어한다.
                v_Search_yn = 'Y';
            }

            private void ITEM_NAME_EditValueChanged(object pSender)
            {
                VIEW_NAME.EditValue = ITEM_NAME.EditValue;
            }

        #endregion




    }
}