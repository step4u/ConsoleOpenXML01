using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Formula.Functions;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using SixLabors.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ConsoleOpenXML01
{
    internal class Program
    {
        public static string Title = "영상 위변조 분석 결과서";
        public static string[] Sections = new string[]
        {
            "분석 영상 정보",
            "분석 영상 (이미지)",
            "분석 영상의 해시값",
            "분석 사항",
            "분석 방법",
            "딥러닝 기반 영상 위변조 검출 모델 결과 - (결과값)",
            "위변조 의심 영역 이미지",
            "파일 구조 기반 영상 위변조 검출 모델 결과",
            "종합 결론",
        };

        public static string[] SubSections = new string[]
        {
            "딥러닝",
            "파일구조 기반",
        };

        static void Main(string[] args)
        {
            Test02();
        }

        static string key = "Tldptmdkdl1!";
        static string filePath = @"D:\ReportBase.docx";
        static string newFilePath = @"D:\ReportBase2.docx";

        static void Test01()
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var nfs = new POIFSFileSystem(fs);
                var info = new EncryptionInfo(nfs);
                Decryptor dc = Decryptor.GetInstance(info);
                if (dc.VerifyPassword(key))
                {
                    using (var ds = dc.GetDataStream(nfs))
                    {
                        var doc = new XWPFDocument(ds);

                        SetPageBorders(ST_Border.@double, "000000", 4, 24, true);

                        XWPFParagraph titleParagraph = doc.CreateParagraph();
                        titleParagraph.Alignment = ParagraphAlignment.CENTER;
                        XWPFRun titleRun = titleParagraph.CreateRun();
                        titleRun.SetText(Title);
                        titleRun.FontFamily = "HY헤드라인M";
                        titleRun.FontSize = 28;
                        titleRun.IsBold = false;



                        using (FileStream outFs = new FileStream(newFilePath, FileMode.Create, FileAccess.Write))
                        {
                            doc.Write(outFs);
                        }

                        doc.Close();
                        doc.Dispose();
                    }
                }
            }
        }


        static XWPFDocument doc;
        static void Test02()
        {
            doc = new XWPFDocument();
            
            numId = SetListFormat();
            sNumId = SetListFormat(fmt: ST_NumberFormat.upperLetter);

            if (doc.Document.body.sectPr == null)
                doc.Document.body.sectPr = new CT_SectPr();

            doc.Document.body.sectPr.pgMar = new CT_PageMar
            {
                top = (ulong)(1.6 * 567),
                bottom = (ulong)(2.3 * 567),
                left = (ulong)(2.5 * 567),
                right = (ulong)(2.5 * 567),
                header = (ulong)(1.0 * 567),
                footer = (ulong)(0.7 * 567)
            };

            SetPageBorders(ST_Border.@double, "000000", 4, 24, true);

            InsertBlank();

            AddSection(content: "영상 위변조 분석 결과서", fontFamily: "HY헤드라인M", fontSize: 28, alignment: ParagraphAlignment.CENTER);

            InsertBlank();

            foreach (var section in Sections)
            {
                AddSection(listFmt: numId, content: section, isBold: true);

                int idx = Array.IndexOf(Sections, section);

                if (idx < Sections.Length - 1)
                    InsertBlank();
            }

            foreach (var sub in SubSections)
            {
                string content = "\t" + sub;
                AddSection(listFmt: sNumId, content: content, isBold: true);
            }

            using (FileStream outFs = new FileStream(newFilePath, FileMode.Create, FileAccess.Write))
            {
                doc.Write(outFs);
            }

            doc.Close();
            doc.Dispose();
        }

        static string numId = "0";
        static string sNumId = "0";
        static string SetListFormat(ST_NumberFormat fmt = ST_NumberFormat.@decimal, string numberId = "0", string start = "1", string def = "%1.", string lvl = "0")
        {
            // 목록 스타일 정의
            XWPFNumbering numbering = doc.CreateNumbering();

            // AbstractNum 생성 및 설정
            CT_AbstractNum abstractNum = new CT_AbstractNum();
            abstractNum.abstractNumId = "0";  // 고유 ID 설정

            // 수준 0(첫 번째 수준) 정의 추가
            CT_Lvl level = abstractNum.AddNewLvl();
            level.ilvl = lvl;  // 레벨 설정
            level.numFmt = new CT_NumFmt { val = fmt };  // 글머리 기호 형식
            level.lvlText = new CT_LevelText { val = def };  // 기호 정의

            // 목록의 시작 번호 설정 (수동으로 초기화)
            level.start = new CT_DecimalNumber { val = start };  // 시작 번호 설정

            // AbstractNum을 numbering에 추가
            XWPFAbstractNum xwpfAbstractNum = new XWPFAbstractNum(abstractNum, numbering);
            string abstractNumId = numbering.AddAbstractNum(xwpfAbstractNum);
            string numId = numbering.AddNum(abstractNumId);  // NumID 생성

            return numId;
        }

        static void AddSection(string listFmt = null, string content = "", bool isBold = false, string fontFamily = "굴림", int fontSize = 12, ParagraphAlignment alignment = ParagraphAlignment.LEFT)
        {
            XWPFParagraph para = doc.CreateParagraph();
            para.Alignment = alignment;

            if (listFmt != null)
                para.SetNumID(listFmt);

            XWPFRun run = para.CreateRun();
            run.SetText(content);
            run.FontFamily = fontFamily;
            run.FontSize = fontSize;
            run.IsBold = isBold;
        }

        static void InsertBlank(int count = 1)
        {
            for (int i = 0; i < count; i++)
            {
                doc.CreateParagraph();
            }
        }

        public static void SetPageBorders(ST_Border borderStyle, string color, uint size, uint space, bool applyToFirstPage = true)
        {
            CT_SectPr sectPr = doc.Document.body.sectPr;
            if (sectPr == null)
            {
                sectPr = doc.Document.body.AddNewSectPr();
            }

            if (sectPr.pgBorders == null)
            {
                sectPr.pgBorders = new CT_PageBorders();
            }

            CT_PageBorders pgBorders = sectPr.pgBorders;

            // 테두리 설정 (상, 하, 좌, 우)
            CT_Border[] borders = new CT_Border[4];
            for (int i = 0; i < 4; i++)
            {
                borders[i] = new CT_Border();
                borders[i].color = color;
                borders[i].sz = size;
                borders[i].space = space;
                borders[i].val = borderStyle;
            }

            pgBorders.top = borders[0];
            pgBorders.bottom = borders[1];
            pgBorders.left = borders[2];
            pgBorders.right = borders[3];

            //// 첫 페이지 적용 여부 설정
            //pgBorders.offsetFrom = ST_PageBorderOffset.page;
            //if (!applyToFirstPage)
            //{
            //    //pgBorders.firstPage = false;
            //}
        }

    }


    public enum BorderStyle
    {
        Single,
        Double,
        Dotted,
        Dashed
        // 필요에 따라 다른 스타일 추가
    }

}
