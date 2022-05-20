using System;

namespace FCMF0513
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineFIRST1 = 10; //출력 시작 행 기준 위치

        private int mPrintingLineSTART1 = 10; //라인 출력시 엑셀 시작 행 위치 지정

        private int mPrintingLineFIRST2 = 4; //출력 시작 행 기준 위치

        private int mPrintingLineSTART2 = 4; //라인 출력시 엑셀 시작 행 위치 지정

        private int mMaxLinePrinting = 67;

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치
        private int mIncrementCopyMAX = 67;  //복사되어질 행의 범위

        private int mCopyColumnSTART = 1; //복사되어  진 행 누적 수
        private int mCopyColumnEND = 46;  //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

        private string mtmpString1 = string.Empty;
        private string mtmpString2 = string.Empty;

        private string mMessageValue1 = string.Empty; //소계[EAPP_10046]
        private string mMessageValue2 = string.Empty; //총가용자금[FCM_10222]

        private string mMessageValue3 = string.Empty; //입금[FCM_10212]
        private string mMessageValue4 = string.Empty; //출금[FCM_10213]
        private string mMessageValue5 = string.Empty; //이체[FCM_10230]
        private string mMessageValue6 = string.Empty; //입금합계[FCM_10214]
        private string mMessageValue7 = string.Empty; //출금합계[FCM_10215]

        private bool mIsPrinted = false; //자금현황을 출력 했는지?

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;
        }

        #endregion;

        #region ----- XL File Open -----

        public bool XLFileOpen()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mPrinting.XLOpenFile(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
        }

        #endregion;

        #region ----- MaxIncrement Methods ----

        private int MaxIncrement(string pPathBase, string pSaveFileName)
        {
            int vMaxNumber = 0;
            System.IO.DirectoryInfo vFolder = new System.IO.DirectoryInfo(pPathBase);
            string vPattern = string.Format("{0}*", pSaveFileName);
            System.IO.FileInfo[] vFiles = vFolder.GetFiles(vPattern);

            foreach (System.IO.FileInfo vFile in vFiles)
            {
                string vFileNameExt = vFile.Name;
                int vCutStart = vFileNameExt.LastIndexOf(".");
                string vFileName = vFileNameExt.Substring(0, vCutStart);

                int vCutRight = 3;
                int vSkip = vFileName.Length - vCutRight;
                string vTextNumber = vFileName.Substring(vSkip, vCutRight);
                int vNumber = int.Parse(vTextNumber);

                if (vNumber > vMaxNumber)
                {
                    vMaxNumber = vNumber;
                }
            }

            return vMaxNumber;
        }

        #endregion;

        #region ----- Line Clear All Methods ----

        private void XlLineClear(int pPrintingLine)
        {
            try
            {
                mPrinting.XLActiveSheet("Destination");

                int vStartRow = pPrintingLine + 1;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mCopyLineSUM - 3;
                int vEndCol = mCopyColumnEND - 1;

                if (pPrintingLine > vEndRow)
                {
                    return;
                }

                if (vStartRow > vEndRow)
                {
                    vStartRow = vEndRow; //시작하는 행이 계산후, 끝나는 행 보다 값이 커지므로, 끝나는 행 값을 줌
                }

                if (vStartRow == vEndRow)
                {
                    mPrinting.XL_LineClear(vStartRow, vStartCol, vEndCol);
                }
                else
                {
                    mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }
        }

        #endregion;

        #region ----- Cell Merge Methods ----

        private void CellMerge(int pXLine, int[] pXLColumn, string pMessageValue)
        {
            try
            {
                int vXLine = pXLine - 1;

                mPrinting.XLActiveSheet("Destination");

                object vObject = null;
                mPrinting.XLSetCell(vXLine, pXLColumn[1], vObject);

                int vStartRow = vXLine;
                int vStartCol = pXLColumn[0];
                int vEndRow = vXLine;
                int vEndCol = pXLColumn[2] - 1;

                mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);

                mPrinting.XLSetCell(vXLine, pXLColumn[0], pMessageValue);

                mPrinting.XL_LineDraw_TopBottom(vXLine, vStartCol, (mCopyColumnEND - 1), 2);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[8];
            pXLColumn = new int[8];

            pGDColumn[0] = pGrid.GetColumnToIndex("TR_MANAGE_NAME");       //구분
            pGDColumn[1] = pGrid.GetColumnToIndex("BANK_NAME");            //은행
            pGDColumn[2] = pGrid.GetColumnToIndex("BEGIN_AMOUNT");         //전일잔액
            pGDColumn[3] = pGrid.GetColumnToIndex("DR_AMOUNT");            //당일입금
            pGDColumn[4] = pGrid.GetColumnToIndex("CR_AMOUNT");            //당일출금
            pGDColumn[5] = pGrid.GetColumnToIndex("REMAIN_AMOUNT");        //당일잔액
            pGDColumn[6] = pGrid.GetColumnToIndex("CURRENCY_CODE");        //통화
            pGDColumn[7] = pGrid.GetColumnToIndex("REMAIN_CURR_AMOUNT");   //당일잔액[외화잔액]

            pXLColumn[0] = 2;    //구분
            pXLColumn[1] = 6;    //과목
            pXLColumn[2] = 16;   //전일잔액
            pXLColumn[3] = 22;   //당일입금
            pXLColumn[4] = 27;   //당일출금
            pXLColumn[5] = 32;   //당일잔액
            pXLColumn[6] = 38;   //통화
            pXLColumn[7] = 40;   //당일잔액[외화잔액]
        }

        private void SetArray2(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[5];
            pXLColumn = new int[5];


            pDBColumn[0] = "ACCOUNT_DR_CR_NAME";  //구분
            pDBColumn[1] = "BANK_NAME";           //은행명
            pDBColumn[2] = "REMARK";              //내역
            pDBColumn[3] = "DEPOSIT_AMOUNT";      //현금/예금
            pDBColumn[4] = "BILL_AMOUNT";         //어음

            pXLColumn[0] = 2;    //구분
            pXLColumn[1] = 6;    //은행명
            pXLColumn[2] = 15;   //내역
            pXLColumn[3] = 29;   //현금/예금
            pXLColumn[4] = 36;   //어음
        }

        #endregion;

        #region ----- Convert decimal  Method ----

        private decimal ConvertNumber(string pStringNumber)
        {
            decimal vConvertDecimal = 0m;

            try
            {
                bool isNull = string.IsNullOrEmpty(pStringNumber);
                if (isNull != true)
                {
                    vConvertDecimal = decimal.Parse(pStringNumber);
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertDecimal;
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
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime.ToShortDateString();
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

        #region ----- XlLine1 Methods -----

        private int XlLine1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //[BANK_NAME]은행
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    if (mtmpString2 != vConvertString)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                        mtmpString2 = vConvertString;

                        int vDrawLineColumnSTART = vXLColumnIndex;
                        int vDrawLineColumnEND = pXLColumn[2] - 1;

                        mPrinting.XL_LineDraw_Top(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 1);
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[TR_MANAGE_NAME]구분
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    if (mtmpString1 != vConvertString)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                        mtmpString1 = vConvertString;

                        int vDrawLineColumnSTART = mCopyColumnSTART + 1;
                        int vDrawLineColumnEND = mCopyColumnEND - 1;

                        mPrinting.XL_LineDraw_Top(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);

                        if (mMessageValue1 == vConvertString)
                        {
                            vDrawLineColumnSTART = pXLColumn[0];
                            vDrawLineColumnEND = pXLColumn[2] - 1;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex, null);
                            mPrinting.XLCellMerge(vXLine, vDrawLineColumnSTART, vXLine, vDrawLineColumnEND, false);
                            mPrinting.XLSetCell(vXLine, vDrawLineColumnSTART, vConvertString);

                            //vDrawLineColumnSTART = pXLColumn[0];
                            //vDrawLineColumnEND = mCopyColumnEND - 1;

                            //mPrinting.XLCellAlignmentHorizontal(vXLine, vXLColumnIndex, vXLine, vXLColumnIndex, "C");
                            //mPrinting.XL_LineDraw_TopBottom(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);

                            //vDrawLineColumnSTART = pXLColumn[0];
                            //vDrawLineColumnEND = pXLColumn[1] -1;
                            //mPrinting.XL_LineClearRIGHT(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);
                        }
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[BEGIN_AMOUNT]전일잔액
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[DR_AMOUNT]당일입금
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[CR_AMOUNT]당일출금
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[REMAIN_AMOUNT]당일잔액
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[CURRENCY_CODE]통화
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[REMAIN_CURR_AMOUNT]당일잔액[외화잔액]
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }


            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #region ----- XlLine2 Methods -----

        private int XlLine2(System.Data.DataRow pRow, int pXLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //[BANK_NAME]은행명
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[pDBColumn[1]];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    if (mtmpString2 != vConvertString)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                        mtmpString2 = vConvertString;

                        int vDrawLineColumnSTART = vXLColumnIndex;
                        int vDrawLineColumnEND = pXLColumn[2] - 1;

                        mPrinting.XL_LineDraw_Top(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 1);
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[ACCOUNT_DR_CR_NAME]구분
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[pDBColumn[0]];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    if (mtmpString1 != vConvertString)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                        mtmpString1 = vConvertString;

                        int vDrawLineColumnSTART = mCopyColumnSTART + 1;
                        int vDrawLineColumnEND = mCopyColumnEND - 1;

                        mPrinting.XL_LineDraw_Top(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);

                        if (mMessageValue6 == vConvertString)
                        {
                            vDrawLineColumnSTART = pXLColumn[0];
                            vDrawLineColumnEND = pXLColumn[3] - 1;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex, null);
                            mPrinting.XLCellMerge(vXLine, vDrawLineColumnSTART, vXLine, vDrawLineColumnEND, false);
                            mPrinting.XLSetCell(vXLine, vDrawLineColumnSTART, vConvertString);

                            //vDrawLineColumnSTART = pXLColumn[0];
                            //vDrawLineColumnEND = pXLColumn[1] - 1;
                            //mPrinting.XL_LineClearRIGHT(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);

                            //vDrawLineColumnSTART = pXLColumn[1];
                            //vDrawLineColumnEND = pXLColumn[2] - 1;
                            //mPrinting.XL_LineClearRIGHT(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);

                            mPrinting.XLCellAlignmentHorizontal(vXLine, vXLColumnIndex, vXLine, vXLColumnIndex, "C");
                            mPrinting.XL_LineDraw_TopBottom(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);

                            //vDrawLineColumnSTART = pXLColumn[0];
                            //vDrawLineColumnEND = pXLColumn[1] - 1;
                            //mPrinting.XL_LineClearTOP(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);
                        }
                        else if (mMessageValue7 == vConvertString)
                        {
                            vDrawLineColumnSTART = pXLColumn[0];
                            vDrawLineColumnEND = pXLColumn[3] - 1;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex, null);
                            mPrinting.XLCellMerge(vXLine, vDrawLineColumnSTART, vXLine, vDrawLineColumnEND, false);
                            mPrinting.XLSetCell(vXLine, vDrawLineColumnSTART, vConvertString);

                            //vDrawLineColumnSTART = pXLColumn[0];
                            //vDrawLineColumnEND = pXLColumn[1] - 1;
                            //mPrinting.XL_LineClearRIGHT(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);

                            //vDrawLineColumnSTART = pXLColumn[1];
                            //vDrawLineColumnEND = pXLColumn[2] - 1;
                            //mPrinting.XL_LineClearRIGHT(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);

                            mPrinting.XLCellMerge(vXLine, vDrawLineColumnSTART, vXLine, vDrawLineColumnEND, false);
                            mPrinting.XLCellAlignmentHorizontal(vXLine, vXLColumnIndex, vXLine, vXLColumnIndex, "C");
                            mPrinting.XL_LineDraw_TopBottom(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);

                            //vDrawLineColumnSTART = pXLColumn[0];
                            //vDrawLineColumnEND = pXLColumn[1] - 1;
                            //mPrinting.XL_LineClearTOP(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND);
                        }
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[REMARK]내역
                vXLColumnIndex = pXLColumn[2];
                vObject = pRow[pDBColumn[2]];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[DEPOSIT_AMOUNT]현금/예금
                vXLColumnIndex = pXLColumn[3];
                vObject = pRow[pDBColumn[3]];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[BILL_AMOUNT]어음
                vXLColumnIndex = pXLColumn[4];
                vObject = pRow[pDBColumn[4]];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }


            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, string pDate)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
            mIsPrinted = false;

            int[] vGDColumn;
            int[] vXLColumn;
            string[] vDBColumn;

            int vPrintingLine = mPrintingLineSTART1;

            try
            {
                int vBy = 57; // 52; // 32;
                int vTotal = 0;
                int vCountDBRow = 0;

                int vTotalRow = pGrid.RowCount;
                if (pAdapter.OraSelectData != null)
                {
                    vCountDBRow = pAdapter.OraSelectData.Rows.Count;
                }

                if (vCountDBRow > 0)
                {
                    vTotal = vTotalRow + vCountDBRow + 4 + 2; //4 : 두번째 Sheet[SourceTab2]에 헤더, 타이틀, 페이지번호 까지의 행수, 2 : SourceTab3의 금일필요사항
                }
                else
                {
                    vTotal = vTotalRow + vCountDBRow;
                }
                
                mPageTotalNumber = vTotal / vBy;
                mPageTotalNumber = (vTotal % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);

                #region ----- First Write ----
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    SetArray1(pGrid, out vGDColumn, out vXLColumn);

                    mPrinting.XLActiveSheet("SourceTab1");
                    mPrinting.XLSetCell(4, 2, pDate);

                    mMessageValue1 = mMessageAdapter.ReturnText("EAPP_10046");  //소계[EAPP_10046]
                    mMessageValue2 = mMessageAdapter.ReturnText("FCM_10222");   //총가용자금[FCM_10222]

                    int vSheet1LIneMAX = 51; // 46; //32; //SourceTab1 실제 출력되는 행수
                    if (vTotalRow > vSheet1LIneMAX)
                    {
                        mIncrementCopyMAX = 67; // 62; //42;
                    }
                    else
                    {
                        mIncrementCopyMAX = vTotalRow + 9;
                    }
                    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XlLine1(pGrid, vRow, vPrintingLine, vGDColumn, vXLColumn);
                        mIsPrinted = true;

                        if (vTotalRow == vCountRow)
                        {
                            CellMerge(vPrintingLine, vXLColumn, mMessageValue2);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineFIRST1 - 1);
                                mtmpString1 = string.Empty;
                                mtmpString2 = string.Empty;
                            }
                        }
                    }
                }
                #endregion;

                #region ----- Second Write ----
                vTotalRow = vCountDBRow;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    mMessageValue3 = mMessageAdapter.ReturnText("FCM_10212"); //입금[FCM_10212]
                    mMessageValue4 = mMessageAdapter.ReturnText("FCM_10213"); //출금[FCM_10213]
                    mMessageValue5 = mMessageAdapter.ReturnText("FCM_10230"); //이체[FCM_10230]
                    mMessageValue6 = mMessageAdapter.ReturnText("FCM_10214"); //입금합계[FCM_10214]
                    mMessageValue7 = mMessageAdapter.ReturnText("FCM_10215"); //출금합계[FCM_10215]

                    SetArray2(out vDBColumn, out vXLColumn);

                    if (mIsPrinted == false)
                    {
                        int vSheet2LIneMAX = 63; // 58; // 38; //SourceTab2 실제 출력되는 행수
                        if (vTotalRow > vSheet2LIneMAX)
                        {
                            mIncrementCopyMAX = 67; //62; //42;
                        }
                        else
                        {
                            mIncrementCopyMAX = vTotalRow + 3;
                        }

                        //일자금 계획을 출력 안 했다면
                        mCopyLineSUM = SecondCopyAndPaste2(mPrinting, mCopyLineSUM);

                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART2 - 1);
                    }
                    else
                    {
                        if (mPageTotalNumber > 1)
                        {
                            int vRest = mMaxLinePrinting - mCopyLineSUM; //현재 작업 Sheet에 남은 행수 구하기
                            int vLineWrite = vTotalRow + 4 + 2; //4 : 출력할 행, 헤더, 타이틀, 페이지 까지 포함된 행수, 2 : SourceTab3의 금일필요사항
                            if (vLineWrite < vRest)
                            {
                                mCopyLineSUM++;
                                mIncrementCopyMAX = vTotalRow + 3;
                            }
                            else
                            {
                                mCopyLineSUM = mCopyLineSUM + vRest;

                                mCopyLineSUM++;
                                mIncrementCopyMAX = vTotalRow + 3;
                            }
                        }
                        else
                        {
                            mCopyLineSUM++;
                            mIncrementCopyMAX = vTotalRow + 3;
                        }
                        //일자금 계획을 출력 했다면
                        mCopyLineSUM = SecondCopyAndPaste2(mPrinting, mCopyLineSUM);

                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART2 - 1);
                    }

                    foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XlLine2(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            mCopyLineSUM = ThirdCopyAndPaste(mPrinting, mCopyLineSUM);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineFIRST2 - 1);
                                mtmpString1 = string.Empty;
                                mtmpString2 = string.Empty;
                            }
                        }
                    }
                }
                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            mPrintingLineSTART1 = vPrintingLine;

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            int vPrintingLineEND = mCopyLineSUM - 1;
            if (vPrintingLineEND < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = SecondCopyAndPaste2(mPrinting, mCopyLineSUM);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //첫번째 페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            pPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);


            if (mPageTotalNumber > 1)
            {
                mPageNumber++; //페이지 번호
                mPrinting.XLCellMerge(vCopyPrintingRowEnd, 23, vCopyPrintingRowEnd, 26, false);
                string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
                mPrinting.XLSetCell(vCopyPrintingRowEnd, 23, vPageNumberText); //페이지 번호, XLcell[행, 열]
            }

            return vCopySumPrintingLine;
        }

        private int SecondCopyAndPaste2(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            pPrinting.XLActiveSheet("SourceTab2");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            return vCopySumPrintingLine;
        }

        private int ThirdCopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;
            int vIncrementCopyMAX = 3;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + vIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            pPrinting.XLActiveSheet("SourceTab3");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, vIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //페이지 번호
            int vPageWriteLine = vCopyPrintingRowEnd - 1;
            mPrinting.XLCellMerge(vPageWriteLine, 23, vPageWriteLine, 26, false);
            string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            mPrinting.XLSetCell(vPageWriteLine, 23, vPageNumberText); //페이지 번호, XLcell[행, 열]

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}