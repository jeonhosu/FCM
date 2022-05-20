using System;

namespace FCMF0391
{
    public class Z4Mplus
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGrid = null;

        private System.Drawing.Printing.PrintDocument mPrintDoc = null;
        private System.Windows.Forms.PrintDialog mPrintDialog = null;
        private System.Windows.Forms.PrintPreviewDialog mPrintPreviewDialog;

        private string mMessageError = string.Empty;

        private string mManageNo = string.Empty;
        private string mpAssetName = string.Empty;
        private string mpItemSpec = string.Empty;
        private string mpManageF = string.Empty;
        private string mpManageS = string.Empty;
        private string mpAcquireDate = string.Empty;
        private string mpUseDept = string.Empty;

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public Z4Mplus(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, System.Windows.Forms.PrintDialog pPrintDialog, System.Windows.Forms.PrintPreviewDialog pPrintPreviewDialog)
        {
            mAppInterface = pAppInterface;
            mGrid = pGrid;
            mPrintDialog = pPrintDialog;
            mPrintPreviewDialog = pPrintPreviewDialog;
            mPrintDoc = new System.Drawing.Printing.PrintDocument();
            mPrintDoc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocument_PrintPage);
        }

        #endregion;

        #region ----- Print Page Event -----

        private void PrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintDocument(e);
        }

        #endregion;

        #region ----- Dispose Method -----

        public void Dispose()
        {
            mPrintDoc.Dispose();
            mPrintPreviewDialog.Dispose();
            mPrintDialog.Dispose();
        }

        #endregion;

        #region ----- PRINTING Method -----

        public void PRINTING()
        {
            try
            {
                System.Windows.Forms.DialogResult vResult = mPrintDialog.ShowDialog();

                string vPrintName = mPrintDialog.PrinterSettings.PrinterName;
                short vCopies = mPrintDialog.PrinterSettings.Copies;    //인쇄 대화상자 표시

                if (vResult == System.Windows.Forms.DialogResult.OK)
                {
                    mPrintDoc.PrinterSettings.PrinterName = vPrintName; //선택한 프린터 기종
                    mPrintDoc.PrinterSettings.Copies = vCopies;         //인쇄매수

                    mAppInterface.OnAppMessageEvent(vPrintName);
                    System.Windows.Forms.Application.DoEvents();

                    int vIndexCheckBox = mGrid.GetColumnToIndex("SELECT_CHECK_YN");
                    int vTotalRow = mGrid.RowCount;

                    for (int nRow = 0; nRow < vTotalRow; nRow++)
                    {
                        if ((string)mGrid.GetCellValue(nRow, vIndexCheckBox) == "Y") //체크한 항목에 한해 출력하기 위한 조건문
                        {
                            mGrid.CurrentCellMoveTo(nRow, 0);
                            mGrid.Focus();
                            mGrid.CurrentCellActivate(nRow, 0);

                            GetValue(nRow);
                                                        
                            mPrintDoc.Print();                            

                            //미리보기
                            //mPrintPreviewDialog.ClientSize = new System.Drawing.Size(500, 320);
                            //mPrintPreviewDialog.PrintPreviewControl.Zoom = 100.0F / 100.0F;
                            //mPrintPreviewDialog.Document = mPrintDoc;
                            //mPrintPreviewDialog.ShowDialog();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- PRINTING Method -----

        private void PrintDocument(System.Drawing.Printing.PrintPageEventArgs e) //미리보기를 위한 객체
        {
            string vTextPrint = string.Empty;
            System.Drawing.Font vPrintFont = null;
            
            try
            {
                //관리번호
                vTextPrint = mManageNo;
                vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);  // 폰트체 : EAN 13, 폰트 사이즈, 폰트스타일 지정
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 115, 65);  // 글씨 출력 위치 x = 30, y = 0

                //자산명 + 규격
                vTextPrint = mpAssetName + mpItemSpec; //Arial Black
                vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 115, 105);

                //관리자(정)
                vTextPrint = mpManageF;
                vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 115, 145);

                //관리자(부)
                vTextPrint = mpManageS;
                vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 285, 145);

                //취득일자
                vTextPrint = mpAcquireDate;
                vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 115, 185);

                //사용부서
                vTextPrint = mpUseDept;

                /*
                //사용부서 초기 설정
                vTextPrint = mpUseDept;
                vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 285, 185);
                */

                //사용 부서명이 긴 경우(예 : 임원(경영전략본부)), 문자 길이를 체크하여 두 줄로 출력되도록 구현한 부분임.
                if (vTextPrint.Length > 1 && vTextPrint.Substring(0, 2) == "임원" && vTextPrint.Substring(2, 1) == "(")
                {
                    vPrintFont = new System.Drawing.Font("EAN 13", 8, System.Drawing.FontStyle.Regular);
                    e.Graphics.DrawString(vTextPrint.Substring(2), vPrintFont, System.Drawing.Brushes.Black, 285, 180);    //본부명
                    e.Graphics.DrawString(vTextPrint.Substring(0, 2), vPrintFont, System.Drawing.Brushes.Black, 285, 194); //임원
                }
                else if (vTextPrint.Length > 2 && vTextPrint.Substring(0, 3) == "본부장" && vTextPrint.Substring(3, 1) == "(")
                {
                    vPrintFont = new System.Drawing.Font("EAN 13", 8, System.Drawing.FontStyle.Regular);
                    e.Graphics.DrawString(vTextPrint.Substring(3), vPrintFont, System.Drawing.Brushes.Black, 285, 180);    //본부명
                    e.Graphics.DrawString(vTextPrint.Substring(0, 3), vPrintFont, System.Drawing.Brushes.Black, 285, 194); //본부장
                }
                else if (vTextPrint == "중국부래주재원")
                {
                    vPrintFont = new System.Drawing.Font("EAN 13", 9, System.Drawing.FontStyle.Regular);
                    e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 285, 185); 
                }
                else
                {
                    vPrintFont = new System.Drawing.Font("EAN 13", 10, System.Drawing.FontStyle.Regular);
                    e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, 285, 185);                
                }
               
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Get Value Method -----

        private void GetValue(int pRow)
        {
            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                //관리번호
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("ASSET_CODE"));
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mManageNo = string.Format("{0}", vConvertString);
                }
                else
                {
                    mManageNo = "";
                }

                //자산명
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("ASSET_DESC"));
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mpAssetName = string.Format("{0}", vConvertString);
                }
                else
                {
                    mpAssetName = "";
                }

                //규격
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("ITEM_SPEC"));
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mpItemSpec = string.Format(", (규격){0}", vConvertString);
                }
                else
                {
                    mpItemSpec = "";
                }

                //관리자(정)
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("FIRST_USER_NAME"));
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mpManageF = string.Format("{0}", vConvertString);
                }
                else
                {
                    mpManageF = "";
                }

                //관리자(부)
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("SECOND_USER_NAME"));
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mpManageS = string.Format("{0}", vConvertString);
                }
                else
                {
                    mpManageS = "";
                }

                //취득일자
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("ACQUIRE_DATE"));
                IsConvert = IsConvertDate(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mpAcquireDate = string.Format("{0}", vConvertString);
                }
                else
                {
                    mpAcquireDate = "";
                }

                //사용부서
                vObject = mGrid.GetCellValue(pRow, mGrid.GetColumnToIndex("USE_DEPT_NAME"));
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    mpUseDept = string.Format("{0}", vConvertString);
                }
                else
                {
                    mpUseDept = "";
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }            
        }

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {
            bool vIsConvert = false;
            pConvertString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        pConvertString = pObject as string;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(object pObject, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        pConvertDecimal = vIsConvertNum;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out string pConvertDateTimeShort)
        {
            bool vIsConvert = false;
            pConvertDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is System.DateTime;
                    if (vIsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime.ToString("yyyy-MM-dd", null);
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        #endregion;
    }
}
