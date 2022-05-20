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

namespace EAPF0216
{
    public partial class EAPF0216 : Office2007Form
    {
        ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        private void EAPF0216_Load(object sender, EventArgs e)
        {
            idaPLAN_EXCHANGE_RATE.FillSchema();
        }

        #region ----- Variables -----

        #endregion;

        #region ----- Constructor -----

        public EAPF0216(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            DateTime vDateTime = (DateTime) DateTime.Today;
            int vYear = vDateTime.Year;
            int vMonth = vDateTime.Month;
            string vYearMonth = string.Format("{0}-{1:D2}", vYear, vMonth);
            iedPERIOD_NAME.EditValue = vYearMonth;   

        }

        #endregion;

        #region ----- Private Methods ----

        private void Header_Setting()
        {
            //DateTime vDateTime = (DateTime)iedPERIOD_NAME.EditValue;
            //int vYear = vDateTime.Year;
            //int vMonth = vDateTime.Month;
            //string vYearMonth = string.Format("{0}-{1:D2}", vYear, vMonth);
            //isGridAdvEx1.SetCellValue("PERIOD_NAME", vYearMonth);   \
            
        }


        //private void GRID_DefaultValue()
        //{
        //    //idcLOCAL_DATE.ExecuteNonQuery();
        //    //isGridAdvEx1.SetCellValue("EFFECTIVE_DATE_FR", DateCellQueryInfoEventArgs  ) ;
        //    //iedDELIVERY_DATE.("EFFECTIVE_DATE_FR", DateTime.Today);
            
        //}

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    idaPLAN_EXCHANGE_RATE.Fill();

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaPLAN_EXCHANGE_RATE.IsFocused == true)
                    {
                        idaPLAN_EXCHANGE_RATE.AddOver();
                        Header_Setting();
                        //GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaPLAN_EXCHANGE_RATE.IsFocused == true)
                    {
                        idaPLAN_EXCHANGE_RATE.AddUnder();
                        Header_Setting();
                        //GRID_DefaultValue();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaPLAN_EXCHANGE_RATE.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPLAN_EXCHANGE_RATE.IsFocused == true)
                    {
                        idaPLAN_EXCHANGE_RATE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPLAN_EXCHANGE_RATE.IsFocused == true)
                    {
                        idaPLAN_EXCHANGE_RATE.Delete();
                    }
                }
            }
        }

     

        #endregion;

        private void ilaCalendar_SelectedRowData(object pSender)
        {
            igrPLAN_EXCHANGE_RATE.SetCellValue("APPLY_START_DATE", iDate.ISMonth_1st(igrPLAN_EXCHANGE_RATE.GetCellValue("PERIOD_NAME")));
            igrPLAN_EXCHANGE_RATE.SetCellValue("APPLY_END_DATE", iDate.ISMonth_Last(igrPLAN_EXCHANGE_RATE.GetCellValue("PERIOD_NAME")));
        }


        //#region -- Default Value Setting --
        
        //#endregion

       


    }

}