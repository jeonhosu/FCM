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

using System.IO;
using Syncfusion.GridExcelConverter;

namespace FCMF0316
{
    public partial class FCMF0316_DPR : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
          
        #endregion;

        #region ----- Constructor -----

        public FCMF0316_DPR()
        {
            InitializeComponent();
        }

        public FCMF0316_DPR(Form pMainForm, ISAppInterface pAppInterface, object pSALE_HEADER_ID, 
                                object pASSET_ID, object pASSET_CODE, object pASSET_DESC, 
                                object pDPR_TYPE, object pDPR_TYPE_NAME)
        {
            InitializeComponent(); 
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_SALE_HEADER_ID.EditValue = pSALE_HEADER_ID;
            V_ASSET_ID.EditValue = pASSET_ID;
            V_ASSET_CODE.EditValue = pASSET_CODE;
            V_ASSET_DESC.EditValue = pASSET_DESC;

            V_DPR_TYPE.EditValue = pDPR_TYPE;
            V_DPR_TYPE_DESC.EditValue = pDPR_TYPE_NAME;
        }

        #endregion;

        #region ----- Private Methods -----
             
        private void Search_DB()
        {
            IDA_ASSET_SALE_DPR.Fill();     
        }
        #endregion;

        #region ----- Initialize Event -----
           
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
         
        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                     
                }
                //전표 행 위치 보정 위해 주석 
                //else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                //{
                //    if (idaSLIP_LINE.IsFocused)
                //    {
                //        if (Check_Sub_Panel() == false)
                //        {
                //            return;
                //        } 

                //        idaSLIP_LINE.AddOver();
                //        InsertSlipLine();
                //    }
                //    else
                //    {
                //        if (Check_SlipHeader_Added() == true)
                //        {
                //            return;
                //        }
                //        else
                //        {
                //            idaSLIP_HEADER.SetSelectParamValue("W_SLIP_HEADER_ID", 0);
                //            idaSLIP_HEADER.Fill();

                //            if (Check_Sub_Panel() == false)
                //            {
                //                return;
                //            } 

                //            idaSLIP_HEADER.AddOver();
                //            idaSLIP_LINE.AddOver();
                //            InsertSlipHeader();
                //            InsertSlipLine();

                //            SLIP_DATE.Focus();
                //        }
                //    }
                //}
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    ;
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

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0316_ACCOUNT_Load(object sender, EventArgs e)
        {             
              
        }

        private void FCMF0316_ACCOUNT_Shown(object sender, EventArgs e)
        {
             
        }

        private void BTN_INQUIRY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Search_DB();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ASSET_SALE_DPR.Cancel();
            DialogResult = DialogResult.Cancel;
            this.Close();
        }
         
        #endregion

        #region ----- Lookup Event ----- 
         
        #endregion       

        #region ----- Adapter Event -----
          
        #endregion


    }
}