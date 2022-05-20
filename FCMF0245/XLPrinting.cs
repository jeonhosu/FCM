// 프린트파일 // XLprint file


using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;
namespace FCMF0245
{
    public class XLPrinting
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
        private string mSourceTab3 = "SourceTab3";
        private string mSourceTab4 = "SourceTab4";
        private int mPageNumber = 0;
        private bool mIsNewPage = false;
        private string mXLOpenFileName = string.Empty;
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;
        ///////////////////////////////////////////////////////////////////////////////////////
        //---------------------------------------------------------------------- Values -----//
        private int mCopy_StartCol = 1;     // 복사될 Column 시작값
        private int mCopy_StartRow = 1;     // 복사될 Row 시작값
        private int mCopy_EndCol = 41;      // 복사될 Column 최대값
        private int mCopy_EndRow = 63;      // 복사될 Row 최대값
        private int mStart_Row_1st = 14;      // 인쇄되는 row 위치(Page 1st)
        private int mEnd_Row_1st = 61;        // 종료되는 row 위치(Page 1st)
        private int mStart_Row_2nd = 3;       // 인쇄되는 row 위치(Page 2nd)
        private int mEnd_Row_2nd = 61;        // 종료되는 row 위치(Page 2nd)
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
        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
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
            try
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
            catch
            {
                //
            }
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
        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pPrint_Appr_Person, string Person ,string Dept ,DateTime Date ,string Period)
        {
            string vString = string.Empty;
            object vObject = string.Empty;

            try
            {
                mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택

                //승인단계 인쇄.
                //승인 여부에 처리.
                //작성자
                if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_N"]).Equals("Y"))
                {
                    vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_N"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 26, vString);

                    //승인 이미지
                    mPrinting.XLActiveSheet("Sheet1");
                    object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 3);

                    mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택
                    object vRangeDestination = mPrinting.XLGetRange(8, 26, 9, 28);
                    mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                }

                //검토1
                if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A"]).Equals("Y"))
                {
                    vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 29, vString);

                    //승인 이미지
                    mPrinting.XLActiveSheet("Sheet1");
                    object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 3);

                    mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택
                    object vRangeDestination = mPrinting.XLGetRange(8, 29, 9, 31);
                    mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                }

                //검토2
                if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A1"]).Equals("Y"))
                {
                    vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A1"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 32, vString);

                    //승인 이미지
                    mPrinting.XLActiveSheet("Sheet1");
                    object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 3);

                    mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택
                    object vRangeDestination = mPrinting.XLGetRange(8, 32, 9, 34);
                    mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                }

                //확인
                if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A2"]).Equals("Y"))
                {
                    vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A2"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 35, vString);

                    //승인 이미지
                    mPrinting.XLActiveSheet("Sheet1");
                    object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 3);

                    mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택
                    object vRangeDestination = mPrinting.XLGetRange(8, 35, 9, 37);
                    mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                }

                //승인
                if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_B"]).Equals("Y"))
                {
                    vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_B"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 38, vString);

                    //승인 이미지
                    mPrinting.XLActiveSheet("Sheet1");
                    object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 3);

                    mPrinting.XLActiveSheet(mSourceTab1); //셀에 문자를 넣기 위해 쉬트 선택
                    object vRangeDestination = mPrinting.XLGetRange(8, 38, 9, 40);
                    mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                }


                // 작성자
                vString = string.Format("{0}", Person);
                mPrinting.XLSetCell(8,18, vString);
                mPrinting.XLSetCell(62, 5, vString);
                //작성부서
                vString = string.Format("{0}", Dept);
                mPrinting.XLSetCell(6, 18, vString);
                mPrinting.XLSetCell(1, 12, vString);
                //작성일자 
                vString = string.Format("{0}", Date);
                mPrinting.XLSetCell(62, 32, vString);
                //월마감
                vString = string.Format("{0}", Period);
                mPrinting.XLSetCell(1, 23, vString);

                //////////셀에 문자를 넣기 위해 쉬트 선택
                mPrinting.XLActiveSheet(mSourceTab2);
                // 작성자
                vString = string.Format("{0}", Person);
                mPrinting.XLSetCell(62, 5, vString);
                //작성일자 
                vString = string.Format("{0}", Date);
                mPrinting.XLSetCell(62, 32, vString);

                //////////셀에 문자를 넣기 위해 쉬트 선택
                mPrinting.XLActiveSheet(mSourceTab3); 
                // 작성자
                vString = string.Format("{0}", Person);
                mPrinting.XLSetCell(62, 5, vString);
                //작성일자 
                vString = string.Format("{0}", Date);
                mPrinting.XLSetCell(62, 32, vString);

                //////////셀에 문자를 넣기 위해 쉬트 선택
                mPrinting.XLActiveSheet(mSourceTab4); 
                // 작성자
                vString = string.Format("{0}", Person);
                mPrinting.XLSetCell(62, 5, vString);
                //작성일자 
                vString = string.Format("{0}", Date);
                mPrinting.XLSetCell(62, 32, vString);
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
                //[M월마감]
                vObject = pRow["M0_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);
                //[증감율1]
                vObject = pRow["RATE1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:###.00}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);
                //[M-1월마감]
                vObject = pRow["M1_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vString);
                //[증감율2]
                vObject = pRow["RATE2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:###.00}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);
                //[M-2월마감]
                vObject = pRow["M2_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);
                //-------------------//
                vXLine = vXLine + 2; // 다음 행에 출력될 그리드 증가 값
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
        private int LineWrite2(System.Data.DataRow pRow, int pCurrentLine, int pRowcount)
        {
            int vXLine = pCurrentLine; //엑셀에 내용이 표시되는 행 번호
            
            object vObject;
            string vString = string.Empty;
            decimal vConvertDecimal = 0;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mDestination); //셀에 문자를 넣기 위해 쉬트 선택
            try
            {

                //0 - 일련번호
                //vGDColumnIndex = pGDColumn[0];
                //vXLColumnIndex = pXLColumn[0];
                //vObject = Convert.ToDecimal(pGridRow) + 1;
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //vObject = pRowcount; 
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                    
                //    vString = string.Format("{0}", vObject );
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                mPrinting.XLSetCell(vXLine, 2, pRowcount);
                //[계정코드]
                vObject = pRow["ACCOUNT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);
                //[계정과목]
                vObject = pRow["ACCOUNT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vString);
                //[당월]
                vObject = pRow["M0_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vString);
                //[증감율]
                vObject = pRow["RATE1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###.00}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);
                //[M-1월]
                vObject = pRow["M1_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 24, vString);
                //[증감율2]
                vObject = pRow["RATE2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###.00}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 31, vString);
                //[M-1월]
                vObject = pRow["M2_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);
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
        private int LineWrite3(System.Data.DataRow pRow, int pCurrentLine)
        {
            int vXLine = pCurrentLine; //엑셀에 내용이 표시되는 행 번호
            object vObject;
            string vString = string.Empty;
            mPrinting.XLActiveSheet(mDestination); //셀에 문자를 넣기 위해 쉬트 선택
            try
            {
                //[구분 - 년월]
                vObject = pRow["PERIOD"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 2, vString);
                //[판관비]
                vObject = pRow["GA_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 5, vString);
                //[구성비율]
                vObject = pRow["RATE1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 11, vString);

                //[제조비용]
                vObject = pRow["MFG_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vString);

                //[구성비율2]
                vObject = pRow["RATE2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vString);
                //비용합계]
                vObject = pRow["TOT_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vString);
                //매출금액]
                vObject = pRow["SALE_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);
                //매출액대비비율]
                vObject = pRow["RATE3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 36, vString);
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
        private int LineWrite4(System.Data.DataRow pRow, int pCurrentLine)
        {
            int vXLine = pCurrentLine; //엑셀에 내용이 표시되는 행 번호
            object vObject;
            string vString = string.Empty;
            mPrinting.XLActiveSheet(mDestination); //셀에 문자를 넣기 위해 쉬트 선택
            try
            {
                //[거래처 ]
                vObject = pRow["VENDOR_FULL_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 2, vString);
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
                mPrinting.XLSetCell(vXLine, 8, vString);
                //[적요]
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 13, vString);
                //[공급가액]
                vObject = pRow["SUPPLY_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vString);
                
                //[부가세]
                vObject = pRow["VAT_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 28, vString);

                //[합계]
                vObject = pRow["TOTAL"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vString);

                //[승인상태]
                vObject = pRow["SLIP_STATUS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 38, vString);
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
        public int MainWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pCLOSE_SLIP_SUMMARY
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pCLOSE_SLIP_ACCOUNT
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pCLOSE_SLIP_MONTHLY
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pCLOSE_SLIP_LIST
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pPRINT_APPROVAL_PERSON
                            , string pPerson_Name
                            , string pDepartment
                            , DateTime pDate 
                            , string pPeriod
                            )
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
            int vPrintingLine = mStart_Row_1st;
            HeaderWrite(pPRINT_APPROVAL_PERSON, pPerson_Name, pDepartment ,pDate ,pPeriod); // Header 부분 Print
            mCopyLineSUM = CopyAndPaste(mPrinting, mSourceTab1, mCopy_StartCol);
            try
            {
                int vTotalRow = 0;
                vTotalRow = pCLOSE_SLIP_SUMMARY.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pCLOSE_SLIP_SUMMARY.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mStart_Row_1st = LineWrite(vRow, mStart_Row_1st);   // Line 부분 Print
                        if (vTotalRow != vCountRow)
                        {
                            IsNewPage(vPrintingLine, mSourceTab1);
                            vPrintingLine = vPrintingLine + 2;
                            if (mIsNewPage == true)
                            {
                                mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                                mMulti = mMulti + 1;
                                vPrintingLine = mStart_Row_1st;

                            }
                        }
                    }
                    //계정별
                    //mCopyLineSUM = 
                    
                    CopyAndPaste2(mPrinting, mSourceTab2, mStart_Row_1st + 1);

                    vTotalRow = pCLOSE_SLIP_ACCOUNT.CurrentRows.Count;
                    mStart_Row_1st = mStart_Row_1st + 3;
                    vPrintingLine = vPrintingLine + 5;
                    vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pCLOSE_SLIP_ACCOUNT.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        mStart_Row_1st = LineWrite2(vRow, mStart_Row_1st, vCountRow);   // Line 부분 Print
                        if (vTotalRow != vCountRow)
                        {
                            IsNewPage(vPrintingLine, mSourceTab2);
                            vPrintingLine = vPrintingLine + 1;
                            if (mIsNewPage == true)
                            {
                                mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                                mMulti = mMulti + 1;
                                vPrintingLine = mStart_Row_1st;

                            }
                        }
                    }

                    if(mCopyLineSUM - vPrintingLine - 3>= 15 )
                    {
                        //이어서 출력                        
                        mPrinting.XL_LineClearALL(vPrintingLine + 1, 2, mCopyLineSUM - 3, mCopy_EndCol);
                        mPrinting.XL_LineDraw(vPrintingLine, 2, mCopy_EndCol - 1, 2);
                        CopyAndPaste3(mPrinting, mSourceTab3, mStart_Row_1st + 1);
                    }
                    else
                    {
                        //새로운장에 찍어줌 
                        mPrinting.XL_LineClearALL(vPrintingLine + 1, 2, mCopyLineSUM - 3, mCopy_EndCol);
                        mPrinting.XL_LineDraw(vPrintingLine, 2, mCopy_EndCol - 1, 2);
                        mCopyLineSUM = CopyAndPaste(mPrinting, mSourceTab3, mCopyLineSUM);

                        mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                        mMulti = mMulti + 1;
                    }

                    //월 누적집계
                    vTotalRow = pCLOSE_SLIP_MONTHLY.CurrentRows.Count;

                    mStart_Row_1st = mStart_Row_1st + 3;
                    vPrintingLine = mStart_Row_1st; // vPrintingLine + 4;
                    vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pCLOSE_SLIP_MONTHLY.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        mStart_Row_1st = LineWrite3(vRow, mStart_Row_1st);   // Line 부분 Print
                        if (vTotalRow != vCountRow)
                        {
                            IsNewPage1(vPrintingLine, mSourceTab3);
                            vPrintingLine = vPrintingLine + 1;
                            if (mIsNewPage == true)
                            {
                                mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                                mMulti = mMulti + 1;
                                vPrintingLine = mStart_Row_1st;

                            }
                        }
                    }
                    int vline = 12 - vTotalRow;
                    if (mCopyLineSUM - vPrintingLine - vline - 3 >= 3)
                    {
                        //이어서 출력                        
                        CopyAndPaste4(mPrinting , mSourceTab4, mStart_Row_1st + vline + 1);
                    }
                    else
                    {
                        //새로운장에 찍어줌 
                        mCopyLineSUM = CopyAndPaste(mPrinting, mSourceTab4, mCopyLineSUM);
                        
                        mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                        mMulti = mMulti + 1;                     
                    }

                    //월 마감내역
                    vTotalRow = pCLOSE_SLIP_LIST.CurrentRows.Count;

                    mStart_Row_1st = mStart_Row_1st + vline + 3;
                    vPrintingLine = mStart_Row_1st; // vPrintingLine + 4;
                    vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pCLOSE_SLIP_LIST.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        mStart_Row_1st = LineWrite4(vRow, mStart_Row_1st);   // Line 부분 Print
                        if (vTotalRow != vCountRow)
                        {
                            IsNewPage(vPrintingLine, mSourceTab4);
                            vPrintingLine = vPrintingLine + 1;
                            if (mIsNewPage == true)
                            {
                                mStart_Row_1st = (mMulti * mCopy_EndRow) + mStart_Row_2nd;
                                mMulti = mMulti + 1;
                                vPrintingLine = mStart_Row_1st;

                            }
                        }
                    }
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
        private void IsNewPage(int pCurrentLine,string pSourceTab)
        {
            if (mEnd_Row_1st <= pCurrentLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, pSourceTab, mCopyLineSUM);
                mEnd_Row_1st = mCopyLineSUM - 3; // mEnd_Row_2nd;
                pCurrentLine = mStart_Row_2nd;
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage1(int pCurrentLine, string pSourceTab)
        {
            if (mEnd_Row_1st <= pCurrentLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste2(mPrinting, pSourceTab, mCopyLineSUM);
                mEnd_Row_1st = mCopyLineSUM - 4;
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

            mPrinting.XL_LineDraw(pCopySumPrintingLine - 3,2, mCopy_EndCol - 1, 2);
            
            return vCopySumPrintingLine;
        }

        private int CopyAndPaste2(XL.XLPrint pPrinting, string pSourceTab1, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호
            int vCopySumPrintingLine = pCopySumPrintingLine;
            //mPrinting.XLActiveSheet(pSourceTab1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow - pCopySumPrintingLine-1, mCopy_EndCol);
            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;
            pPrinting.XLActiveSheet(mDestination);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, mCopy_EndRow -2, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }

        private int CopyAndPaste3(XL.XLPrint pPrinting, string pSourceTab1, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호
            int vCopySumPrintingLine = pCopySumPrintingLine;
            mPrinting.XLActiveSheet(pSourceTab1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, 15, mCopy_EndCol);
            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;
            pPrinting.XLActiveSheet(mDestination);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART +15 , mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }

        private int CopyAndPaste4(XL.XLPrint pPrinting, string pSourceTab1, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호
            int vCopySumPrintingLine = pCopySumPrintingLine;
            mPrinting.XLActiveSheet(pSourceTab1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopyLineSUM - pCopySumPrintingLine-2, mCopy_EndCol);
            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;
            pPrinting.XLActiveSheet(mDestination);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, mCopyLineSUM -3, mCopy_EndCol);
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

        public void PrintPreview(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrintPreview();
        }
        #endregion;
        // 엑셀 파일로 출력시
        #region ----- Save Methods ----
        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);
            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        public void Delete()
        {
            mPrinting.XLDeleteSheet("SourceTab1");
            mPrinting.XLDeleteSheet("SourceTab2");
            mPrinting.XLDeleteSheet("SourceTab3");
            mPrinting.XLDeleteSheet("SourceTab4");
        }
        #endregion;
    }
}