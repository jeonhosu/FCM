using System;
using ISCommonUtil;

namespace FCMF0209
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // 쉬트명 정의.
        private string mTargetSheet = "PRINT";
        private string mSourceSheet1 = "SOURCE1";
        private string mSourceSheet2 = "SOURCE2";
        
        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.
        
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 12;
        private int mCopy_EndRow = 36;

        private int m1stLastRow = 49;       //첫장 최종 인쇄 라인.
        
        private int mPrintingLastRow = 37;  //최종 인쇄 라인 다음 라인.

        private int mPromptRow = 1;
        private int mCurrentRow = 2;       //현재 인쇄되는 row 위치.
        
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
        {// 파일명 뒤에 일련번호 증가.
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

        #region ----- Array Set 1 (총합계)----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("CUSTOMER_COUNT");
            pGDColumn[1] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("GL_AMOUNT_5");
            pGDColumn[3] = pGrid.GetColumnToIndex("GL_AMOUNT_4");
            pGDColumn[4] = pGrid.GetColumnToIndex("GL_AMOUNT_3");
            pGDColumn[5] = pGrid.GetColumnToIndex("GL_AMOUNT_2");
            pGDColumn[6] = pGrid.GetColumnToIndex("GL_AMOUNT_1");
            pGDColumn[7] = pGrid.GetColumnToIndex("VAT_AMOUNT_5");
            pGDColumn[8] = pGrid.GetColumnToIndex("VAT_AMOUNT_4");
            pGDColumn[9] = pGrid.GetColumnToIndex("VAT_AMOUNT_3");
            pGDColumn[10] = pGrid.GetColumnToIndex("VAT_AMOUNT_2");
            pGDColumn[11] = pGrid.GetColumnToIndex("VAT_AMOUNT_1");
                        
            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 19;
            pXLColumn[1] = 24;
            pXLColumn[2] = 28;
            pXLColumn[3] = 31;
            pXLColumn[4] = 34;
            pXLColumn[5] = 37;
            pXLColumn[6] = 40;
            pXLColumn[7] = 42;
            pXLColumn[8] = 45;
            pXLColumn[9] = 48;
            pXLColumn[10] = 51;
            pXLColumn[11] = 54;
        }

        #endregion;

        #region ----- Array Set 2 (명세) -----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[15];
            pXLColumn = new int[15];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = 0;
            pGDColumn[1] = pGrid.GetColumnToIndex("TAX_REG_NO");
            pGDColumn[2] = pGrid.GetColumnToIndex("CUSTOMER_DESC");
            pGDColumn[3] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("GL_AMOUNT_5");
            pGDColumn[5] = pGrid.GetColumnToIndex("GL_AMOUNT_4");
            pGDColumn[6] = pGrid.GetColumnToIndex("GL_AMOUNT_3");
            pGDColumn[7] = pGrid.GetColumnToIndex("GL_AMOUNT_2");
            pGDColumn[8] = pGrid.GetColumnToIndex("GL_AMOUNT_1");
            pGDColumn[9] = pGrid.GetColumnToIndex("VAT_AMOUNT_5");
            pGDColumn[10] = pGrid.GetColumnToIndex("VAT_AMOUNT_4");
            pGDColumn[11] = pGrid.GetColumnToIndex("VAT_AMOUNT_3");
            pGDColumn[12] = pGrid.GetColumnToIndex("VAT_AMOUNT_2");
            pGDColumn[13] = pGrid.GetColumnToIndex("VAT_AMOUNT_1");
            pGDColumn[14] = 0;


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 3;
            pXLColumn[1] = 6;
            pXLColumn[2] = 12;
            pXLColumn[3] = 22;
            pXLColumn[4] = 25;
            pXLColumn[5] = 28;
            pXLColumn[6] = 31;
            pXLColumn[7] = 34;
            pXLColumn[8] = 37;
            pXLColumn[9] = 39;
            pXLColumn[10] = 42;
            pXLColumn[11] = 45;
            pXLColumn[12] = 48;
            pXLColumn[13] = 51;
            pXLColumn[14] = 53;
        }

        #endregion;

        #region ----- Array Set 2  : Adapter 적용시 ----

        //private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{// 아답터의 table 값.
        //    pGDColumn = new int[10];
        //    pXLColumn = new int[10];

        //    pGDColumn[0] = pTable.Columns.IndexOf("PO_TYPE_NAME");
        //    pGDColumn[1] = pTable.Columns.IndexOf("DISPLAY_NAME");
        //    pGDColumn[2] = pTable.Columns.IndexOf("PO_DATE");
        //    pGDColumn[3] = pTable.Columns.IndexOf("PO_NO");
        //    pGDColumn[4] = pTable.Columns.IndexOf("SUPPLIER_SHORT_NAME");
        //    pGDColumn[5] = pTable.Columns.IndexOf("PRICE_TERM_NAME");
        //    pGDColumn[6] = pTable.Columns.IndexOf("PAYMENT_METHOD_NAME");
        //    pGDColumn[7] = pTable.Columns.IndexOf("PAYMENT_TERM_NAME");
        //    pGDColumn[8] = pTable.Columns.IndexOf("REMARK");
        //    pGDColumn[9] = pTable.Columns.IndexOf("STEP_DESCRIPTION");


        //    pXLColumn[0] = 9;   //PO_TYPE_NAME
        //    pXLColumn[1] = 25;  //DISPLAY_NAME
        //    pXLColumn[2] = 42;  //PO_DATE
        //    pXLColumn[3] = 54;  //PO_NO
        //    pXLColumn[4] = 9;   //SUPPLIER_SHORT_NAME
        //    pXLColumn[5] = 35;  //PRICE_TERM_NAME
        //    pXLColumn[6] = 14;  //PAYMENT_METHOD_NAME
        //    pXLColumn[7] = 41;  //PAYMENT_TERM_NAME
        //    pXLColumn[8] = 7;   //REMARK
        //    pXLColumn[9] = 49;  //금액
        //}

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {// 문자열 여부 체크 및 해당 값 리턴.
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
        {// 숫자 여부 체크 및 해당 값 리턴.
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

        private bool IsConvertDate(object pObject, out System.DateTime pConvertDateTimeShort)
        {// 날짜 여부 체크 및 해당 값 리턴.
            bool vIsConvert = false;
            pConvertDateTimeShort = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is System.DateTime;
                    if (vIsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime;
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

        #region ----- Header Write Method ----

        public void HeaderWrite(object pGL_DATE, object pACCOUNT_CODE, object pACCOUNT_DESC)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                //기준일자
                vXLine = 1;
                vXLColumn = 3;
                mPrinting.XLSetCell(vXLine, vXLColumn, pGL_DATE);

                // 계정코드
                vXLine = 2;
                vXLColumn = 3;
                mPrinting.XLSetCell(vXLine, vXLColumn, pACCOUNT_CODE);

                // 계정명
                vXLine = 2;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, pACCOUNT_DESC);                
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;
        
        #region ----- Excel Line Wirte Methods ----

        public int LineWrite(string pTerritory, InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;
            string vVisible_YN = "0";

            int vCurrentCol = 1;
            int vTotalRow = pGRID.RowCount;
            int vTotalCol = pGRID.ColCount;            
            decimal vNumberValue = 0;

            object vDecimalDigit = 0;            
            object vColumnType  = null;
            object vValue = null;
            object vPrintValue = null;

            try
            {
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.                

                if (vTotalRow > 0)
                {
                    #region ----- Write Page Copy(SourceSheet => TargetSheet) ----
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                    #endregion;
                    
                    for (int c = 0; c < vTotalCol; c++)
                    {// 프롬프트 표시.
                        vVisible_YN = iString.ISNull(pGRID.GridAdvExColElement[c].Visible, "0");
                        if (vVisible_YN == "1")
                        {
                            if (pTerritory == "TL1_KR")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL1_KR;
                            }
                            else if (pTerritory == "TL2_CN")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL2_CN;
                            }
                            else if (pTerritory == "TL3_VN")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL3_VN;
                            }
                            else if (pTerritory == "TL4_JP")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL4_JP;
                            }
                            else if (pTerritory == "TL5_XAA")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL5_XAA;
                            }

                            if (iString.ISNull(vValue) == string.Empty)
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].Default;
                            }
                            vCurrentCol = vCurrentCol + 1;
                            mPrinting.XLSetCell(mPromptRow, vCurrentCol, vValue);
                        }
                    }

                    mPrinting.XLCellAlignmentHorizontal(mPromptRow, 1, mPromptRow, vCurrentCol, "C");
                    mCopy_EndCol = vCurrentCol;  // copy 영역 지정.
                    vCurrentCol = 1;
                    for (int r = 0; r < vTotalRow; r++)
                    {//Row
                        for (int c = 0; c < vTotalCol; c++)
                        {//Col
                            vVisible_YN = iString.ISNull(pGRID.GridAdvExColElement[c].Visible, "0");
                            if (vVisible_YN == "1")
                            {
                                vCurrentCol = vCurrentCol + 1;
                                vValue = pGRID.GetCellValue(r, c);
                                vColumnType = pGRID.GridAdvExColElement[c].ColumnType;
                                vDecimalDigit = pGRID.GridAdvExColElement[c].DecimalDigits;
                                if (iString.ISNull(vColumnType) == "NumberEdit")
                                {
                                    try
                                    {
                                        vNumberValue = iString.ISDecimaltoZero(vValue);
                                        if (iString.ISNumtoZero(vDecimalDigit) > 0)
                                        {
                                            vPrintValue = string.Format("{0:###,###,###,###,###,###,###,###,###.####}", vNumberValue);
                                        }
                                        else
                                        {
                                            vPrintValue = string.Format("{0:###,###,###,###,###,###,###,###,###}", vNumberValue);
                                        }
                                    }
                                    catch
                                    {
                                        vPrintValue = vValue;
                                    }
                                    mPrinting.XLCellAlignmentHorizontal(mCurrentRow, vCurrentCol, mCurrentRow, vCurrentCol, "R");
                                }
                                else
                                {
                                    vPrintValue = vValue;
                                }
                                mPrinting.XLSetCell(mCurrentRow, vCurrentCol, vPrintValue);
                            }
                            vMessage = String.Format("Writing - [{0}/{1}]", r, vTotalRow);
                            mAppInterface.OnAppMessageEvent(vMessage);
                            System.Windows.Forms.Application.DoEvents();
                        }
                        vCurrentCol = 1;
                        mCurrentRow = mCurrentRow + 1;
                    }
                    mPrinting.XLColumnAutoFit(1, 1, mCurrentRow, mCopy_EndCol);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (mPageNumber == 1)
            {
                if (pPageRowCount == m1stLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - m1stLastRow - 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
            else
            {
                if (pPageRowCount == mPrintingLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - mPrintingLastRow - 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
        }

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNumber(string pActiveSheet, object pPageNumber)
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.
            //int vXLRow = 51; //엑셀에 내용이 표시되는 행 번호
            //int vXLCol = 51;
            //if (iString.ISDecimaltoZero(pPageNumber) == 1)
            //{//첫장 적용
            //    vXLRow = 51; //엑셀에 내용이 표시되는 행 번호
            //    vXLCol = 52;
            //}
            //else
            //{//첫장 외.
            //    vXLRow = 54; //엑셀에 내용이 표시되는 행 번호
            //    vXLCol = 51;
            //}

            //try
            //{ // 원본을 복사해서 타겟 에 복사해 넣음.(
            //    mPrinting.XLActiveSheet(pActiveSheet);
            //    mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            //}
            //catch (System.Exception ex)
            //{
            //    mMessageError = ex.Message;
            //    mAppInterface.OnAppMessageEvent(mMessageError);
            //    System.Windows.Forms.Application.DoEvents();
            //}
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            // page수 표시.
            mPageNumber = mPageNumber + 1;
            XLPageNumber(pActiveSheet, mPageNumber);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);            
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol); 
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

            return vPasteEndRow;


            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //pPrinting.XLActiveSheet("SourceTab1");
            //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLActiveSheet("Destination");
            //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            //mPageNumber++; //페이지 번호
            //// 페이지 번호 표시.
            ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            ////int vRowSTART = vCopyPrintingRowEnd - 2;
            ////int vRowEND = vCopyPrintingRowEnd - 2;
            ////int vColumnSTART = 30;
            ////int vColumnEND = 33;
            ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //페이지 번호, XLcell[행, 열]

            //return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            if (iString.ISNull(pSaveFileName) == string.Empty)
            {
                return;
            }

            //int vMaxNumber = MaxIncrement(pSavePath.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", pSavePath, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
