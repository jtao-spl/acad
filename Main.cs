using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using static acad.TextTools;
//using Excel = NetOffice.ExcelApi;
using System.IO;

namespace acad
{
    public class Main
    {
        [CommandMethod("SelectDim")]
        public void Select()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            //PromptSelectionResult psr = ed.SelectAll();
            // (enget (car (entsel)))    //获取图元数据

            TypedValue[] values = new TypedValue[] {
                new TypedValue((int)DxfCode.Operator,"<or"),
                new TypedValue((int)DxfCode.Start,"mtext"),
                new TypedValue((int)DxfCode.Start,"TOLERANCE"),
                new TypedValue((int)DxfCode.Start,"TEXT"),
                new TypedValue((int)DxfCode.Start,"DIMENSION"),
                new TypedValue((int)DxfCode.Operator,"or>"),

            };
            SelectionFilter sfilter = new SelectionFilter(values);
            PromptSelectionResult psr = ed.GetSelection(sfilter);
            if (psr.Status != PromptStatus.OK) return;
            SelectionSet sSet = psr.Value;
            //this.PrintProperty(sSet);
            this.generatePingCeTable(sSet);
            ed.WriteMessage("数据提取成功，请在系统中进行进一步操作。");
        }

        public void generatePingCeTable(SelectionSet sSet)
        {
            List<SizedElement> ses = new List<SizedElement>();
            List<SurfaceRoughness> srs = new List<SurfaceRoughness>();
            List<OtherRequirement> ors = new List<OtherRequirement>();
            List<GeometricalTolerance> gts = new List<GeometricalTolerance>();
            List<SafetyRequirement> secs = new List<SafetyRequirement>();

            ObjectId[] ids = sSet.GetObjectIds();

            Database db = HostApplicationServices.WorkingDatabase;

            for (int i = 0; i < ids.Length; i++)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    
                    //Entity ent = (Entity)ids[i].GetObject(OpenMode.ForRead, true);
                    Entity ent = (Entity)trans.GetObject(ids[i],OpenMode.ForRead, false);
                    string ent_type = ent.GetType().Name;
                    switch (ent_type)
                    {
                        case "MText":
                            MText mtext = (MText)ent;

                            if (mtext.Contents.Contains(TextSpecialSymbol.Degree))
                            {
                                SizedElement e = new SizedElement();
                                e.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.ANGLE;
                                e.baseSize = Convert.ToDecimal(mtext.Contents.Replace(TextSpecialSymbol.Degree, ""));
                                ses.Add(e);

                            }
                            else if (mtext.Text.Contains("∅"))
                            {
                                if (mtext.Text.StartsWith("∅"))
                                {
                                    SizedElement e = new SizedElement();
                                    e.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.DIAMETER;
                                    //格式："%%C{\\Ftxt,@extfont2|c134;15.6}"


                                    string origin = mtext.Text.Substring(1);
                                    if (origin.Contains('+') || origin.Contains('-') || origin.Contains('/'))
                                    {
                                        string[] split2 = origin.Split('+');
                                        e.baseSize = Convert.ToDecimal(split2[0]);
                                        string bound = split2[1].Replace(',', '.');
                                        string[] split3 = bound.Split('/');
                                        e.upperSize = Convert.ToDecimal(split3[0]);
                                        e.lowerSize = Convert.ToDecimal(split3[1]);

                                    }
                                    else
                                    {
                                        e.baseSize = Convert.ToDecimal(origin);
                                        e.upperSize = 0;
                                        e.lowerSize = 0;
                                      

                                    }
                                    ses.Add(e);
                                }
                                else
                                {
                                    //1×{\\Ftxt,@extfont2 | c134; 0.5}
                                    //"6×∅22+0,15/ 0"  text
                                    //"6×∅22+0,15/ 0"

                                    //{\\fSimSun|b0|i0|c134|p2;6×}∅22{\\H0.6x;\\S+0,15^ 0;}
                                    SizedElement e = new SizedElement();
                                    e.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.DIAMETER;
                                    string[] split1 = mtext.Text.Split('∅');

                                    decimal count = Convert.ToDecimal(split1[0].ExtractNumber());
                                    string origin = split1[1];
                                    if (origin.Contains('+') || origin.Contains('-') || origin.Contains('/'))
                                    {
                                        string[] split2 = origin.Split('+');
                                        e.baseSize = Convert.ToDecimal(split2[0]);
                                        string bound = split2[1].Replace(',', '.');
                                        string[] split3 = bound.Split('/');
                                        e.upperSize = Convert.ToDecimal(split3[0]);
                                        e.lowerSize = Convert.ToDecimal(split3[1]);

                                    }
                                    else
                                    {
                                        e.baseSize = Convert.ToDecimal(origin);
                                        e.upperSize = 0;
                                        e.lowerSize = 0;
                                    }
                                    ses.Add(e);
                                }
                                break;
                            }
                            else if (mtext.Text.StartsWith("Ra"))
                            {
                                SurfaceRoughness sr = new SurfaceRoughness();
                                sr.RoughnessType = "Ra";
                                sr.RoughnessValue = mtext.Text.Substring(2);
                                srs.Add(sr);

                            }
                            else if (mtext.Text.StartsWith("R"))
                            {
                                SizedElement e2 = new SizedElement();
                                e2.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.RADIAL;
                                e2.baseSize = Convert.ToDecimal(mtext.Text.Substring(1));
                                ses.Add(e2);

                            }
                            else if (mtext.Text.Contains("min") || mtext.Text.Contains("max"))
                            {
                                SurfaceRoughness sr = new SurfaceRoughness();
                                sr.RoughnessType = "Ra";
                                sr.RoughnessValue = mtext.Text;
                                srs.Add(sr);
                            }
                            else if (mtext.Text.Contains("技术要求"))
                            {
                                OtherRequirement oreq = new OtherRequirement();
                                oreq.requirement = mtext.Text;
                                ors.Add(oreq);
                            }
                            else
                            {
                                //OtherRequirement oreq = new OtherRequirement();
                                //oreq.requirement = mtext.Text;
                                //ors.Add(oreq);
                            }

                            break;
                        case "RotatedDimension"://转角标注
                                                //Dimtm Specifies the minimum tolerance limit for dimension text. 下限
                                                //Dimtp Specifies the maximum tolerance limit for dimension text. 上限
                            RotatedDimension rotatedDimension = (RotatedDimension)ent;

                            SizedElement ele = new SizedElement();
                            if (rotatedDimension.Prefix == "%%c")
                            {
                                ele.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.DIAMETER;//直径
                            }
                            else
                            {
                                ele.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.LINE;//直线
                            }
                            ele.baseSize = Convert.ToDecimal(rotatedDimension.Measurement);
                            if (rotatedDimension.Dimtol)
                            {
                                ele.upperSize = Convert.ToDecimal(rotatedDimension.Dimtp);
                                ele.lowerSize = Convert.ToDecimal(-rotatedDimension.Dimtm);

                            }
                            else
                            {
                                ele.upperSize = 0;
                                ele.lowerSize = 0;

                            }
                            ses.Add(ele);

                            //ed.WriteMessage("rotatedDimension" + Convert.ToDecimal(rotatedDimension.Measurement) + ", upper:" + rotatedDimension.Dimtp + ",lowder:" + rotatedDimension.Dimtm + "\n");
                            break;
                        case "RadialDimension"://角度标注
                            RadialDimension rdimension = (RadialDimension)ent;
                            SizedElement element1 = new SizedElement();
                            element1.baseSize = Convert.ToDecimal(rdimension.Measurement);
                            element1.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.ANGLE;
                            if (rdimension.Dimtol)
                            {
                                element1.upperSize = Convert.ToDecimal(rdimension.Dimtp);
                                element1.lowerSize = Convert.ToDecimal(-rdimension.Dimtm);
                            }
                            else
                            {
                                element1.upperSize = 0;
                                element1.lowerSize = 0;
                            }
                            ses.Add(element1);
                            break;
                        case "DiametricDimension":
                            DiametricDimension dDimension = (DiametricDimension)ent;
                            SizedElement element2 = new SizedElement();
                            element2.baseSize = Convert.ToDecimal(dDimension.Measurement);
                            element2.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.DIAMETER;
                            if (dDimension.Dimtol)
                            {
                                element2.upperSize = Convert.ToDecimal(dDimension.Dimtp);
                                element2.lowerSize = Convert.ToDecimal(-dDimension.Dimtm);
                            }
                            else
                            {
                                element2.upperSize = 0;
                                element2.lowerSize = 0;
                            }
                            ses.Add(element2);

                            break;
                        case "FeatureControlFrame":
                            FeatureControlFrame featureControlFrame = (FeatureControlFrame)ent;
                            GeometricalTolerance gt = new GeometricalTolerance();
                            string rawText = featureControlFrame.Text;
                            string[] rawTextArr = rawText.Split("%%v".ToCharArray());
                            string finalVal = "";
                            for (int m = 1; m < rawTextArr.Length; m++)
                            {
                                finalVal += rawTextArr[m];
                            }
                            gt.TonerancePrecision = finalVal;
                            gt.ToneranceType = featureControlFrame.Text.Substring(7, 1);
                            gts.Add(gt);

                            break;
                        default:
                            //ed.WriteMessage(ent.ToString() + "\n");
                            break;
                    }
                    ent.Dispose();
                    trans.Commit();
                    
                }
            }
            Element element = new Element();
            element.sizedElements = ses.ToArray();
            element.surfaceRoughnesses = srs.ToArray();
            element.otherRequirements = ors.ToArray();
            element.geometricalTolerances = gts.ToArray();
            element.safetyRequirements = secs.ToArray();
            //this.SaveElementToExcel(element);
            ComponentTool tool = new ComponentTool();
            int ComponentId = tool.CreateComponent();
            tool.CreateComponentSize(ComponentId, element);

            string fileName = Path.GetFileName(db.Filename);
            FileStream fs = new FileStream(db.Filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            BinaryReader br = new BinaryReader(fs);
            byte[] content = br.ReadBytes((int)fs.Length);
            br.Close();
            fs.Close();
            tool.SaveOriginFile(ComponentId, fileName, content);
            
        }

        //public void SaveElementToExcel(Element element)
        //{
        //    Database db = HostApplicationServices.WorkingDatabase;
        //    Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        //    string directoryName = Path.GetDirectoryName(db.Filename);
        //    string fileName = Path.GetFileNameWithoutExtension(db.Filename);
        //    PromptSaveFileOptions opt = new PromptSaveFileOptions("save excel file");
        //    opt.DialogCaption = "Save Excel";
        //    opt.Filter = "Excel 97-2003 工作簿(*.xls)|*.xls|Excel工作簿(*.xlsx)|*.xlsx";
        //    opt.FilterIndex = 1;
        //    opt.InitialDirectory = directoryName;
        //    opt.InitialFileName = fileName;
        //    PromptFileNameResult pfr = ed.GetFileNameForSave(opt);
        //    if (pfr.Status != PromptStatus.OK) return;
        //    fileName = pfr.StringResult;
        //    //Directory.SetAccessControl(Path.GetDirectoryName(fileName), new DirectorySecurity(directoryName, AccessControlSections.All));
        //    Excel.Application excelApp = new Excel.Application();
        //    Excel.Workbook book = excelApp.Workbooks.Add();
        //    Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.Add();
        //    ExcelTool.DrawElements(element, sheet, 3, 5);
        //    book.SaveAs(fileName);
        //    excelApp.Quit();
        //    excelApp.Dispose();
        //}


        //private void PrintProperty(SelectionSet sSet)
        //{
        //    ObjectId[] ids = sSet.GetObjectIds();
        //    Database db = HostApplicationServices.WorkingDatabase;
        //    Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        //    using (Transaction trans = db.TransactionManager.StartTransaction())
        //    {
        //        for (int i = 0; i < ids.Length; i++)
        //        {
        //            Entity ent = (Entity)ids[i].GetObject(OpenMode.ForRead);
        //            string ent_type = ent.GetType().Name;
        //            switch (ent_type)
        //            {
        //                case "MText":
        //                    MText mtext = (MText)ent;
        //                    ed.WriteMessage("mtext:" + mtext.Text + "\n");

        //                    Console.WriteLine(mtext.Text);
        //                    break;
        //                case "RotatedDimension":
        //                    //Dimtm Specifies the minimum tolerance limit for dimension text.
        //                    //Dimtp Specifies the maximum tolerance limit for dimension text.
        //                    RotatedDimension rotatedDimension = (RotatedDimension)ent;
        //                    ed.WriteMessage("rotatedDimension.Prefix==\\%\\%c, ");
        //                    ed.WriteMessage("rotatedDimension.Prefix:" + rotatedDimension.Prefix + "rotatedDimension:" + Convert.ToDecimal(rotatedDimension.Measurement) + ", upper:" + rotatedDimension.Dimtp + ",lowder:" + rotatedDimension.Dimtm + "\n");
        //                    //LinearDimension linearDimension = new LinearDimension(rotatedDimension.Measurement, rotatedDimension.Dimtm, rotatedDimension.Dimtp);
        //                    //linearDimension.Print();
        //                    break;
        //                case "RadialDimension"://角度标注
        //                    RadialDimension rdimension = (RadialDimension)ent;
        //                    ed.WriteMessage("RadialDimension.Measurement" + rdimension.Measurement);
        //                    break;
        //                case "FeatureControlFrame":
        //                    FeatureControlFrame featureControlFrame = (FeatureControlFrame)ent;
        //                    ed.WriteMessage(featureControlFrame.Text);
        //                    break;
        //                default:
        //                    ed.WriteMessage(ent.ToString() + "\n");
        //                    break;
        //            }


        //        }
        //        trans.Commit();
        //    }
        //}

    }
    //public class PingCeTable
    //{
    //    public ExamSample examSample;
    //    public ExamBaseInfo baseInfo;
    //    public ScoreRule scoreRule;
    //    public Element element;
    //    public List<Evalueation> evalueations;
    //    public void addSizedElement(MText mtext) { }
    //    public void addSizedElement(RotatedDimension rotatedDim) { }

    //}
    public class ExamSample
    {
        public int titleRowNo;//开始绘制的行数
        public int sampleStarColoumNo = 1;//开始绘制的列数
        public int sampleColoums;//占据的列数
        public int samplRowCount;//样例占据的行
    }
    public struct ExamBaseInfo //考核信息
    {
        public DateTime ExamDateTime; //考核日期  及考核時間
        public string ElementNo;//零件编号
        public string StudentClass; //班级
        public string StudentName;
        public string StudentNo;
        public string ExamProject; //考核项目
        public string ExamTeacher1;
        public string ExamTeacher2;
    }
    public struct ScoreRule //评测说明
    {
        public string ScoreType; //类型
        public string ScoreSymble;//符号
        public string ScoreStandard;//评测标准
    }
    public class SizedElement
    {
        public ELEMENT_SIZED_ELEMENT_SUB_TYPE sizeType { get; set; }
        public decimal baseSize { get; set; }
        public decimal upperSize { get; set; }
        public decimal lowerSize { get; set; }
        //public decimal totalElementScore { get; set; }//manual setting
        //public ToleraceLevel toleranceLevel { get; set; }
        //public decimal CalculateDelta()
        //{
        //    if (this.upperSize == 0 && this.lowerSize == 0)
        //    {
        //        decimal delta = AcdBaseTool.CalculateDeltaByToleranceLevelAndSizeField(this.toleranceLevel, this.baseSize.getSizeField());
        //        return delta;
        //    }
        //    return 0;
        //}

    }

    public struct GeometricalTolerance //形位公差
    {
        public string ToneranceType; //公差类型
        public string TonerancePrecision;//公差精度,可能带有直径符号
        //public decimal totalElementScore;//manual setting

    }
    public struct SurfaceRoughness//表面粗糙度
    {
        public string RoughnessType;
        public string RoughnessValue;
        //public decimal totalElementScore;//manual setting

    }
    public struct OtherRequirement
    {
        public string requirement;//其他要求
        //public decimal totalElementScore;
    }
    public struct SafetyRequirement
    {
        public string safetyRequirement;
        //public decimal totalElementScore;
    }
    public struct Element
    {

        public SizedElement[] sizedElements;
        public GeometricalTolerance[] geometricalTolerances;
        public SurfaceRoughness[] surfaceRoughnesses;
        public OtherRequirement[] otherRequirements;
        public SafetyRequirement[] safetyRequirements;
    }
    public class Evalueation
    {
        public decimal selfValue;
        public decimal selfScore;
        public decimal groupValue;
        public decimal groupScore;
        public decimal finalValue;
        public decimal finalScore;
    }
    public enum ToleraceLevel
    {
        FirstLevel = 0,
        Middle = 1,
        Rought = 2,
        MostRought = 3,
    }
    public enum ELEMENT_FIRST_TYPE
    {
        SIZED_ELEMENT,
        GEOMETRICAL_TOLERNACE,
        SURFACE_ROUGHNESS,
        OTHER
    }
    public enum ELEMENT_SIZED_ELEMENT_SUB_TYPE
    {
        LINE,
        DIAMETER,
        RADIAL,
        ANGLE,
    }
    public enum ELEMENT_GEO_TOLERANCE_SUB_TYPE
    {
        A
    }
    public enum ELEMENT_SURFACE_ROUGHNESS_SUB_TYPE
    {
        RA
    }

}
