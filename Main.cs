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
using System.Text.RegularExpressions;

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
                 new TypedValue((int)DxfCode.Start, "INSERT"),
                new TypedValue((int)DxfCode.Operator,"or>"),

            };
            SelectionFilter sfilter = new SelectionFilter(values);
            //PromptSelectionResult psr = ed.GetSelection(sfilter);
            PromptSelectionResult psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK) return;
            SelectionSet sSet = psr.Value;
            //this.PrintProperty(sSet);
            this.generatePingCeTable(sSet);
            ed.WriteMessage("数据提取成功，请在系统中进行进一步操作。");
        }
        public void PrintProperty(SelectionSet sSet)
        {
            ObjectId[] ids = sSet.GetObjectIds();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;

            for (int i = 0; i < ids.Length; i++)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {

                    //entity ent = (entity)ids[i].getobject(openmode.forread, true);
                    Entity ent = (Entity)trans.GetObject(ids[i], OpenMode.ForRead, false);
                    if(ent.GetType() == typeof(BlockReference))
                    { 
                        DBObjectCollection entitySets = new DBObjectCollection();
                        ent.Explode(entitySets);
                        string text = "";
                        foreach (Entity explodedObj in entitySets)
                        {
                            // if the exploded entity is a blockref or mtext
                            // then explode again
                            if (explodedObj.GetType() == typeof(MText))
                            {
                                text += ((MText)explodedObj).Text;
                            }
                        }
                        ed.WriteMessage("提取块参照中的文字:" + text + "\n");
                    }
                    //BlockTable bt = db.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    //string ent_type = ent.GetType().Name;
                    //switch (ent_type)
                    //{
                    //    case "blockreference":
                    //        BlockReference btr = null;
                    //        if (trans.GetObject(ids[i], OpenMode.ForRead) is BlockReference)
                    //        {
                    //        }

                    //        break;
                    //    default:
                    //        break;
                    //}
                    trans.Commit();
                }
            }
        }

         /**
         * 输入mtext的文本,返回角度数据
         */
        public SizedElement ExtractAngleFromString(string text)
        {
            SizedElement e = new SizedElement();
            e.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.ANGLE;
            e.baseSize = Convert.ToDecimal(text.Replace(TextSpecialSymbol.Degree, ""));
            return e;
        }
        /**
         * 输入mtext的文本,返回直径数据
         */
        public SizedElement ExtractDiameterFromString(string origin)
        {
            SizedElement e = new SizedElement();
            e.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.DIAMETER;
            //格式："%%C{\\Ftxt,@extfont2|c134;15.6}"

            if (origin.Contains('+') && origin.Contains('/'))
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
            return e;
        }
        /**
         * 输入mtext的文本,返回线性尺寸数据
         */
        public SizedElement ExtractLineFromString(string text)
        {
            SizedElement e = new SizedElement();
            text = text.Replace(',', '.').Replace(" ", "");
            if (text.Contains('+') && text.Contains('/'))
            {
                string[] split = text.Split('+');
                e.baseSize = Convert.ToDecimal(split[0]);
                string[] split2 = split[1].Split('/');
                e.upperSize = Convert.ToDecimal(split2[0]);
                e.lowerSize = Convert.ToDecimal(split2[1]);
            }
            else
            {
                e.baseSize = Convert.ToDecimal(text);
                e.upperSize = 0;
                e.lowerSize = 0;
            }
            return e;
        }
        /**
         * 输入转角标注，输出尺寸数据
         */
        public SizedElement ExtractRotatedDimensionFromEntity(RotatedDimension rotatedDimension)
        {
            SizedElement ele = new SizedElement();
            if (rotatedDimension.Prefix == "%%c" || rotatedDimension.Prefix == "%%C")
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
            return ele;
        }
        /**
         * 输入半径标注，输出半径数据
         */
        public SizedElement ExtractRadialDimensionFromEntity(RadialDimension radialDimension)
        {
            SizedElement element = new SizedElement();
            element.baseSize = Convert.ToDecimal(radialDimension.Measurement);
            element.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.ANGLE;
            if (radialDimension.Dimtol)
            {
                element.upperSize = Convert.ToDecimal(radialDimension.Dimtp);
                element.lowerSize = Convert.ToDecimal(-radialDimension.Dimtm);
            }
            else
            {
                element.upperSize = 0;
                element.lowerSize = 0;
            }
            return element;
        }
        /**
         * 输入直径标注，输出直径数据
         */
        public SizedElement ExtractDiametricDimensionFromEntity(DiametricDimension dDimension)
        {
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
            return element2;
        }
        /**
         * 是否为未注倒角内容
         */
        public bool isUndeclaredChamfer(string text)
        {
            if (text.Length == 0) return false;
            List<Regex> regList = new List<Regex>();
            regList.Add(new Regex("[\\d]+\\.*[\\d]*[^0-9][\\d]+\\.*[\\d]*°")); // 1.2 ×25°这种  前后均支持小数点
            regList.Add(new Regex("C\\d+\\.*[\\d]*")); //支持C加数值的组合，支持小数
            regList.Add(new Regex("R\\d+\\.*[\\d]*"));//支持R加数值的组合，支持小数
            try
            {
                bool result = regList.Exists(x => x.IsMatch(text));
                return result;
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }
        /**
         * 是否为行位公差数据
         */
        public bool isGeometricalTolerance(string text)
        {
            if (text.Length == 0) return false;
            Regex reg = new Regex("[a-krtu]\\d+\\.*[\\d]*[A-Z]"); //支持的小写字母a-k r t u 开头接数字 再接大写字母的情况
            try
            {
                bool result = reg.IsMatch(text);
                return result;
            }
            catch(RegexMatchTimeoutException)
            {
                return false;
            }
        }
        /**
         * 是否为表面粗糙度
         */
        public bool isSurfaceRoughness(string text)
        {
            if (text.Length == 0) return false;
            Regex reg = new Regex("Ra\\s*\\d+\\.*[\\d]*"); //Ra开头加数值
            try
            {
                bool result = reg.IsMatch(text);
                return result;
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }
        /**
         * 是否为线性尺寸
         */
        public bool isLineSize(string text)
        {
            //几种情况： 1.整数 2 小数 3 小数点换逗号 4 存在上下 如 3.4 

            //1. 小数点换逗号, 同时去掉空格
            text = text.Replace(',', '.').Replace(" ", "");
            try
            {
                //2.整数 小数可直接转成功，返回true
                Convert.ToDecimal(text);
                return true;
            }
            catch (FormatException)
            {
                if (text.Contains('+') && text.Contains('/'))
                {
                    string[] split = text.Split('+');
                    Convert.ToDecimal(split[0]);
                    string[] split2 = split[1].Split('/');
                    Convert.ToDecimal(split2[0]);
                    Convert.ToDecimal(split2[1]);
                    return true;
                }
                return false;
            }
            catch (System.Exception exception) when (exception is OverflowException || exception is FormatException)
            {
                return false;
            }

        }

        public void generatePingCeTable(SelectionSet sSet)
        {
            List<SizedElement> ses = new List<SizedElement>();
            Dictionary<string, int> srs = new Dictionary<string, int>();
            //List<SurfaceRoughness> srs = new List<SurfaceRoughness>();
            //List<OtherRequirement> ors = new List<OtherRequirement>();
            int UndeclaredChamferCount = 0;
            List<GeometricalTolerance> gts = new List<GeometricalTolerance>();
            List<SafetyRequirement> secs = new List<SafetyRequirement>();

            ObjectId[] ids = sSet.GetObjectIds();

            Database db = HostApplicationServices.WorkingDatabase;

            for (int i = 0; i < ids.Length; i++)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    
                    //Entity ent = (Entity)ids[i].GetObject(OpenMode.ForRead, true);
                    Entity ent = (Entity)trans.GetObject(ids[i],OpenMode.ForWrite, false);
                    string ent_type = ent.GetType().Name;
                    switch (ent_type)
                    {
                        case "MText":
                            MText mtext = (MText)ent;
                            //角度尺寸，带°
                            if (mtext.Contents.Contains(TextSpecialSymbol.Degree))
                            {
                                SizedElement e = ExtractAngleFromString(mtext.Contents);
                                ses.Add(e);
                            }
                            //直径尺寸
                            else if (mtext.Text.Contains("∅")) //包含直径符号，存在直径符号开头和 前置有数值倍数的情况
                            {
                                if (mtext.Text.StartsWith("∅"))
                                {
                                    SizedElement e = ExtractDiameterFromString(mtext.Text.Substring(1));
                                    ses.Add(e);
                                }
                                else
                                {
                                    string[] split1 = mtext.Text.Split('∅');
                                    decimal count = Convert.ToDecimal(split1[0].ExtractNumber());
                                    SizedElement e = ExtractDiameterFromString(split1[1]);
                                    ses.Add(e);
                                }
                            }
                            //表面粗糙度
                            else if (mtext.Text.StartsWith("Ra"))
                            {
                                string RoughnessValue = mtext.Text.Substring(2);
                                if (srs.ContainsKey(RoughnessValue))
                                {
                                    srs[RoughnessValue] = srs[RoughnessValue] + 1;
                                }
                                else
                                {
                                    srs.Add(RoughnessValue, 1);
                                }
                            }
                            //半径尺寸
                            else if (mtext.Text.StartsWith("R"))
                            {
                                SizedElement e2 = new SizedElement();
                                e2.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.RADIAL;
                                e2.baseSize = Convert.ToDecimal(mtext.Text.Substring(1));
                                ses.Add(e2);
                            }
                            //未注倒角
                            else if (isUndeclaredChamfer(mtext.Text))
                            {
                                UndeclaredChamferCount += 1;
                            }
                            //线性尺寸
                            else if (isLineSize(mtext.Text))
                            {
                                SizedElement e3 = ExtractLineFromString(mtext.Text);
                                ses.Add((e3));
                            }
                            break;
                        case "DBText":
                            DBText dtext = (DBText)ent;
                            string text = dtext.TextString;
                            if (text.Contains(TextSpecialSymbol.Degree))
                            {
                                SizedElement e = ExtractAngleFromString(text);
                                ses.Add(e);
                            }
                            //直径尺寸
                            else if (text.Contains("∅")) //包含直径符号，存在直径符号开头和 前置有数值倍数的情况
                            {
                                if (text.StartsWith("∅"))
                                {
                                    SizedElement e = ExtractDiameterFromString(text.Substring(1));
                                    ses.Add(e);
                                }
                                else
                                {
                                    string[] split1 = text.Split('∅');
                                    decimal count = Convert.ToDecimal(split1[0].ExtractNumber());
                                    SizedElement e = ExtractDiameterFromString(split1[1]);
                                    ses.Add(e);
                                }
                            }
                            //表面粗糙度
                            else if (text.StartsWith("Ra"))
                            {
                                string RoughnessValue = text.Substring(2);
                                if (srs.ContainsKey(RoughnessValue))
                                {
                                    srs[RoughnessValue] = srs[RoughnessValue] + 1;
                                }
                                else
                                {
                                    srs.Add(RoughnessValue, 1);
                                }
                            }
                            //半径尺寸
                            else if (text.StartsWith("R"))
                            {
                                SizedElement e2 = new SizedElement();
                                e2.sizeType = ELEMENT_SIZED_ELEMENT_SUB_TYPE.RADIAL;
                                e2.baseSize = Convert.ToDecimal(text.Substring(1));
                                ses.Add(e2);
                            }
                            //未注倒角
                            else if (isUndeclaredChamfer(text))
                            {
                                UndeclaredChamferCount += 1;
                            }
                            //线性尺寸
                            else if (isLineSize(text))
                            {
                                SizedElement e3 = ExtractLineFromString(text);
                                ses.Add((e3));
                            }
                            break;
                        case "RotatedDimension"://转角标注                         
                            RotatedDimension rotatedDimension = (RotatedDimension)ent;
                            SizedElement ele =  ExtractRotatedDimensionFromEntity(rotatedDimension);
                            ses.Add(ele);
                            break;
                        case "RadialDimension"://角度标注
                            RadialDimension rdimension = (RadialDimension)ent;
                            SizedElement element1 = ExtractRadialDimensionFromEntity(rdimension);
                            ses.Add(element1);
                            break;
                        case "DiametricDimension": //直径标注
                            DiametricDimension dDimension = (DiametricDimension)ent;
                            SizedElement element2 = ExtractDiametricDimensionFromEntity(dDimension);
                            ses.Add(element2);
                            break;
                        case "FeatureControlFrame": //形位公差
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
                        case "BlockReference": //块参照
                            BlockReference blockReference = (BlockReference)ent;

                            DBObjectCollection entitySets = new DBObjectCollection();
                            ent.Explode(entitySets);
                            //提取块参照中所有文字内容
                            string blockText = "";
                            foreach (Entity explodedObj in entitySets)
                            {
                                if (explodedObj.GetType() == typeof(MText)) //多行文字
                                {
                                    blockText += ((MText)explodedObj).Text;
                                }
                                if(explodedObj.GetType() == typeof(DBText))//单行文字
                                {
                                    blockText += ((DBText)explodedObj).TextString;
                                }
                            }
                            
                            //根据文字内容符合的标注格式模式进行匹配处理。转角标注未打包成块参照，仍按上面的方式进行解析
                            if (isUndeclaredChamfer(blockText)) //未注倒角
                            {
                                UndeclaredChamferCount += 1;
                            }else if (isGeometricalTolerance(blockText)) //行位公差
                            {
                                GeometricalTolerance gt1 = new GeometricalTolerance();
                                gt1.ToneranceType = blockText.Substring(0, 1);
                                gt1.TonerancePrecision = blockText.Substring(1);
                                gts.Add(gt1);
                            }else if (isSurfaceRoughness(blockText)) //表面粗糙度
                            {
                                string RoughnessValue = blockText.Substring(2).Trim();
                                if (srs.ContainsKey(RoughnessValue))
                                {
                                    srs[RoughnessValue] = srs[RoughnessValue] + 1;
                                }
                                else
                                {
                                    srs.Add(RoughnessValue, 1);
                                }
                            }
                            break;
                        default:
                            break;
                    }
                    ent.Dispose();
                    trans.Commit();
                }
            }
            Element element = new Element();
            element.sizedElements = ses.ToArray();
            //element.surfaceRoughnesses = srs.ToArray();
            element.surfaceRoughnesses = srs;
            //element.otherRequirements = ors.ToArray();
            element.geometricalTolerances = gts.ToArray();
            element.UndeclaredChamferCount = UndeclaredChamferCount;
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
        public int UndeclaredChamferCount;
        //public SurfaceRoughness[] surfaceRoughnesses;
        public Dictionary<string, int> surfaceRoughnesses;
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
