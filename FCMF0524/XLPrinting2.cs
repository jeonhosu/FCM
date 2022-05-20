// 프린트파일 // XLprint file


using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;
namespace FCMF0524
{
    public class XLPrinting2
    {
        #region ----- Variables -----
        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private XL.XLPrint mPrinting = null;
        private string mMessageError = string.Empty;
        // 쉬트명 정의.
        private string mDestination = "Destination";
        private string mSourceTab1 = "SourceTab1";
        private string mSourceTab2 = "SourceTab2";
        private int mPageNumber = 0;
        private bool mIsNewPage = false;
        private string mXLOpenFileName = string.Empty;
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;
        ///////////////////////////////////////////////////////////////////////////////////////
        //---------------------------------------------------------------------- Values -----//
        private int mCopy_StartCol = 1;     // 복사될 Column 시작값
        private int mCopy_StartRow = 1;     // 복사될 Row 시작값
        private int mCopy_EndCol = 94;      // 복사될 Column 최대값
        private int mCopy_EndRow = 35;      // 복사될 Row 최대값
        private int mStart_Row_1st = 4;      // 인쇄되는 row 위치(Page 1st)
        private int mEnd_Row_1st = 35;        // 종료되는 row 위치(Page 1st)
        private int mStart_Row_2nd = 1;       // 인쇄되는 row 위치(Page 2nd)
        private int mEnd_Row_2nd = 35;        // 종료되는 row 위치(Page 2nd)
        //---------------------------------------------------------------------- Values -----//
        ///////////////////////////////////////////////////////////////////////////////////////
        private int mMulti = 1;                // 곱셈
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
        public XLPrinting2(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
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
        #region ----- Convert DateTime Methods ----
        private object ConvertDateTime(object pObject)
        {
            object vObject = null;
            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
            return vObject;
        }
        private object ConvertDate(object pObject)
        {
            object vObject = null;
            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
            return vObject;
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
                int vCutRight = 2;
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
        #region ----- Excel Wirte [Header] Methods ----
        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pData_1st)
        {
            string vString = string.Empty;
            try
            {
                mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택
                //이송 번호
                vString = string.Format("{0}", pData_1st.CurrentRow["TRANSFER_NO"]);
                mPrinting.XLSetCell(10, 9, vString);
                // 이송 일자
                vString = string.Format("{0}", pData_1st.CurrentRow["TRANSFER_DATE"]);
                mPrinting.XLSetCell(10, 50, vString);
                // From 작업장
                vString = string.Format("{0}", pData_1st.CurrentRow["FROM_WORKCENTER_DESC"]);
                mPrinting.XLSetCell(11, 9, vString);
                // From 자원
                vString = string.Format("{0}", pData_1st.CurrentRow["FROM_RESOURCE_DESC"]);
                mPrinting.XLSetCell(12, 9, vString);
                // To 작업장
                vString = string.Format("{0}", pData_1st.CurrentRow["TO_WORKCENTER_DESC"]);
                mPrinting.XLSetCell(11, 50, vString);
                // To 자원
                vString = string.Format("{0}", pData_1st.CurrentRow["TO_RESOURCE_DESC"]);
                mPrinting.XLSetCell(12, 50, vString);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }
        #endregion
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
        private bool IsConvertNumber(string pStringNumber, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;
            try
            {
                if (pStringNumber != null)
                {
                    decimal vIsConvertNum = decimal.Parse(pStringNumber);
                    pConvertDecimal = vIsConvertNum;
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
        #region ----- XlLine Methods -----
        private int LineWrite(System.Data.DataRow pRow, int pCurrentLine)
        {
            int vXLine = pCurrentLine; //엑셀에 내용이 표시되는 행 번호
            object vObject;
            string vString = string.Empty;
            mPrinting.XLActiveSheet(mDestination); //셀에 문자를 넣기 위해 쉬트 선택
            try
            {

                //[NO]
                vObject = pRow["ROW_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 1, vString);


                //[거래처코드]
                vObject = pRow["VENDOR_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[거래처명칭]
                vObject = pRow["VENDOR_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vString);

                //[사업자번호]
                vObject = pRow["TAX_REG_NO"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[계정명]
                vObject = pRow["ACCOUNT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vString);
                //[내역 -비고 ]
                vObject = pRow["SLIP_REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 32, vString);

                //[통화]
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 46, vString);

                //[발생환율]
                vObject = pRow["EXCHANGE_RATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vString);
               
                //[외화금액]
                vObject = pRow["GL_CURR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 56, vString);

                //[원화금액]
                vObject = pRow["GL_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 63, vString);

                //[지급예정일]
                vObject = pRow["DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 70, vString);
                //[은행명]
                vObject = pRow["BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 75, vString);

                //[계좌번호]
                vObject = pRow["BANK_ACCOUNT_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 81, vString);

                //[예금주]
                vObject = pRow["ACCOUNT_HOLDER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 89, vString);

                //[합계라인 색상]
                vObject = pRow["COLOR_FLAG"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                    if (vString == "Y")
                    {
                        mPrinting.XLCellColorBrush(vXLine, 1, 94, System.Drawing.Color.LightGray);
                    }
                }
                else
                {
                    vString = string.Empty;
                }
                

                //-------------------//
                vXLine = vXLine + 1; // 다음 행에 출력될 그리드 증가 값
                //-------------------//
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pCurrentLine = vXLine;
            return pCurrentLine;
        }
        #endregion;
        #region ----- Excel Wirte [Line] Methods ----
        public int MainWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pData_1st
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pData_2nd
                            )
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
            int vPrintingLine = mStart_Row_1st;
           // HeaderWrite(pData_1st); // Header 부분 Print
            mCopyLineSUM = CopyAndPaste(mPrinting, mSourceTab1, mCopy_StartCol);
            try
            {
                int vTotalRow = pData_2nd.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pData_2nd.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        mStart_Row_1st = LineWrite(vRow, vPrintingLine);   // Line 부분 Print
                        if (vTotalRow != vCountRow)
                        {
                            IsNewPage(vPrintingLine);
                            vPrintingLine = vPrintingLine + 1;
                            if (mIsNewPage == true)
                            {
                                mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                                mMulti = mMulti + 1;
                                vPrintingLine = mStart_Row_2nd;

                            }
                        }
                    }
                }
                mPrinting.XL_LineClearALL(vPrintingLine + 1, 1, mCopyLineSUM, 94);
                mPrinting.XL_LineDraw(vPrintingLine, 1, 94, 2);


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
        private void IsNewPage(int pCurrentLine)
        {
            if (mEnd_Row_1st <= pCurrentLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceTab2, mCopyLineSUM);
                mEnd_Row_1st = mEnd_Row_2nd;
                pCurrentLine = mStart_Row_2nd;
            }
            else
            {
                mIsNewPage = false;
            }
        }
        #endregion;
        #region ----- Excel Copy&Paste Methods ----
        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, string pSourceTab1, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호
            int vCopySumPrintingLine = pCopySumPrintingLine;
            mPrinting.XLActiveSheet(pSourceTab1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);
            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;
            pPrinting.XLActiveSheet(mDestination);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }
        #endregion;
        // 복사 출력시
        #region ----- Printing Methods ----
        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }
        #endregion;
        // 엑셀 파일로 출력시
        #region ----- Save Methods ----
        public void Save(string pSaveFileName)
        {

            mPrinting.XLDeleteSheet("SourceTab1");
            mPrinting.XLDeleteSheet("SourceTab2");
            //System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            //int vMaxNumber = MaxIncrement(vSaveFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);
            //vSaveFileName = string.Format("{0}\\{1}.xls", vSaveFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }
        #endregion;
    }
}