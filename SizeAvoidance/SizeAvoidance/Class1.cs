//所有程序通用
using System;
using System.Collections.Generic;
using System.Linq;//ICollection转List必用到的包头
using System.Text;
using System.Threading.Tasks;

//revit不同的地方
using Autodesk.Revit;

// Revit常用命名空间
using Autodesk.Revit.UI;// U1基础如UIDocument、 ExternalCommandData
using Autodesk.Revit.DB; // 1 数据基础如Element、Reference
using Autodesk.Revit.Attributes;// 模式如TransactionMode、JounalingMode
using Autodesk.Revit.UI.Selection;//选择如ObjectType
using Autodesk.Revit.DB.Architecture; // 1 建筑如ROOM、Stair<DirectedGraph xmlns="http://schemas.microsoft.com/vs/2009/dgml">
using Autodesk.Revit.DB.Plumbing;//管道专业如Pipe、PipeType
using Autodesk.Revit.DB.Electrical;//电气专业如CableTray
using System.Windows;

//主程序
namespace CalculateWindowsEffectiveArea
{

    //“选择时限选过滤器”的类定义，非主程序类       
    [Transaction(TransactionMode.Manual)]
    public class DimensionSelectionFilter : ISelectionFilter
    {

        public String conditionInChinese = null;//存储条件用字符串
        public bool conditionForRestrictSelectionAllowElement(Element elem)
        {
            //通过门和窗和墙的Category 的ID判断，来避免 Revit语言版本的不兼容                           
            Categories categories = elem.Document.Settings.Categories;
            if (conditionInChinese.IndexOf("门") != -1)
            {
                //开始判断，返回true则算集合中一员，返回false则会选不上
                if (elem is FamilyInstance && (elem.Category.Id ==
                    categories.get_Item(BuiltInCategory.OST_Doors).Id))
                {
                    return true;
                }

            }
            if (conditionInChinese.IndexOf("窗") != -1)
            {
                //开始判断，返回true则算集合中一员，返回false则会选不上
                if (elem is FamilyInstance && (elem.Category.Id ==
                    categories.get_Item(BuiltInCategory.OST_Windows).Id))
                {
                    return true;
                }
            }
            if (conditionInChinese.IndexOf("墙") != -1)
            {
                //开始判断，返回true则算集合中一员，返回false则会选不上
                if (elem.Category.Id ==
                    categories.get_Item(BuiltInCategory.OST_Walls).Id)
                {
                    return true;
                }
            }
            if (conditionInChinese.IndexOf("地板") != -1)
            {
                //开始判断，返回true则算集合中一员，返回false则会选不上
                if (elem.Category.Id ==
                    categories.get_Item(BuiltInCategory.OST_Floors).Id)
                {
                    return true;
                }
            }

            if (conditionInChinese.IndexOf("屋顶") != -1)
            {

                //开始判断，返回true则算集合中一员，返回false则会选不上
                if (elem.Category.Id ==
                    categories.get_Item(BuiltInCategory.OST_Roofs).Id)
                {
                    return true;

                }
            }
            if (conditionInChinese.IndexOf("标注") != -1)
            {

                //开始判断，返回true则算集合中一员，返回false则会选不上
                if (elem.Category.Id ==
                    categories.get_Item(BuiltInCategory.OST_Dimensions).Id)
                {
                    return true;

                }
            }

            return false;

        }

        public bool AllowElement(Element elem)
        {

            return conditionForRestrictSelectionAllowElement(elem);

        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
    //主程序类
    //尺寸文字结构体，用以存储获得的文字信息
    struct TextStruct
    {
        public Dimension dim; //文字所属尺寸
        public double length;//文字长度
        public double height;//文字高度
        public XYZ dirX;//文字×方向
        public XYZ dirY;//文字丫方向
        public List<Line> lines;//文字范围框
        public bool isDim;//是否单段尺寸
        public DimensionSegment dSeg;//文字所属尺寸段
    }

    [Transaction(TransactionMode.Manual)]

    public class HelloWorldCommand : IExternalCommand
    {
        private TextStruct GetTextStruct(Document doc, Dimension dim, DimensionSegment dSeg, bool isDim)
        {
            Curve cCurve = dim.Curve;
            //标注文字位置
            XYZ textPoint;
            //引线终点
            XYZ leaderPoint;
            if (isDim)//单段尺寸
            {
                textPoint = dim.TextPosition;
                leaderPoint = dim.LeaderEndPosition;
            }
            else//尺寸段
            {
                textPoint = dSeg.TextPosition;
                leaderPoint = dSeg.LeaderEndPosition;
            }
            //标注文字位置投影到标注线上
            XYZ textPointProject = cCurve.Project(textPoint).XYZPoint;
            //引线终点投影到标注线上
            XYZ leaderPointProject = cCurve.Project(leaderPoint).XYZPoint;
            //获得文字的丫坐标系
            Line lineY = Line.CreateBound(textPointProject, textPoint);
            XYZ dirY = lineY.Direction;
            //获得文字的×坐标系
            Line lineX = Line.CreateBound(leaderPointProject, textPointProject);
            XYZ dirX = lineX.Direction;
            //获得文字高度
            Line projectLine = Line.CreateBound(leaderPointProject, leaderPoint);
            double height = (projectLine.Length - lineY.Length) * 2;
            //获得文字长度
            double lgh = (lineX.Length) * 2;
            //计算标注文字范围框
            XYZ p1 = textPoint + dirX * lgh / 2;
            XYZ p2 = p1 + dirY * height;
            XYZ p3 = p2 + -dirX * lgh;
            XYZ p4 = p3 - dirY * height;
            //存储范围框
            List<Line> lines = new List<Line>();
            lines.Add(Line.CreateBound(p1, p2));
            lines.Add(Line.CreateBound(p2, p3));
            lines.Add(Line.CreateBound(p3, p4));
            lines.Add(Line.CreateBound(p4, p1));
            //存储标注文字信息
            TextStruct tStruct = new TextStruct();
            tStruct.dim = dim;
            tStruct.length = lgh;
            tStruct.height = height;
            tStruct.dirX = dirX;
            tStruct.dirY = dirY;
            tStruct.lines = lines;
            tStruct.isDim = isDim;
            tStruct.dSeg = dSeg;
            return tStruct;
        }

        public IList<Reference> restrictSelectionFilter(UIDocument uiDoc, String condition, String note)
        {

            DimensionSelectionFilter conditionFilter = new DimensionSelectionFilter();
            conditionFilter.conditionInChinese = condition;
            try
            {
                return uiDoc.Selection.PickObjects(ObjectType.Element, conditionFilter, note);
            }
            catch
            {
                // 如果中断选择，结束命令
                MessageBox.Show("用户中断选择，否则会出错，建议退出");
                return null;
            }

        }
        public Result Execute(ExternalCommandData cD, ref string ms, ElementSet set)
        {
            //此处的下一行，缩进一格，开始编写我们自己的代码                    
            UIDocument uiDoc = cD.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Autodesk.Revit.DB.View activeView = uiDoc.ActiveView;
            List<ElementId> listWindowsAndDoors = new List<ElementId>();
            List<ElementId> elemIds = uiDoc.Selection.GetElementIds().ToList();
            IList<Reference> refers = new List<Reference>();
            if (elemIds.Count != 0)
            {

                ICollection<ElementId> elemIDsCollectionWindows = elemIds;// new List<ElementId>();                            
                                                                          //创建一个收集器内含过滤器
                FilteredElementCollector collectorWindows = new FilteredElementCollector(doc, elemIDsCollectionWindows);
                //开始过滤并给出过滤后的结果为list
                collectorWindows.OfCategory(BuiltInCategory.OST_Dimensions);
                List<ElementId> listsWindows = collectorWindows.ToElementIds().ToList();

                listWindowsAndDoors.AddRange(listsWindows);
                //listWindowsAndDoors.AddRange(listsDoor);
                uiDoc.Selection.SetElementIds(listWindowsAndDoors);

            }
            else
            {


                refers = restrictSelectionFilter(uiDoc, "标注", "有限选过滤器，仅能选上标注，其他会被忽略");
                // 将用户选择的对象加入选集
                foreach (Reference refer in refers)
                {
                    //如果前面没有用WallSelectionFilter限制，则这里要做一次判断
                    listWindowsAndDoors.Add(refer.ElementId);
                }

                uiDoc.Selection.SetElementIds(listWindowsAndDoors);

            }


            //构造尺寸文字结构体集合，用以存储所选尺寸各个文字的信息，tStructs就是我们要取的标注包
            List<TextStruct> tStructs = new List<TextStruct>();
            foreach (Reference rf in refers)
            {

                //关键步骤1:获得标注文字的包围框
                Dimension dim = doc.GetElement(rf) as Dimension;
                if (dim.Segments.IsEmpty)
                {
                    tStructs.Add(GetTextStruct(doc, dim, null, true));
                }
                else
                {
                    foreach (DimensionSegment dSeg in dim.Segments)
                    {
                        tStructs.Add(GetTextStruct(doc, dim, dSeg, false));
                    }
                }
            }
            //新建并启动事务
            Transaction transaction = new Transaction(doc, "标注文字避让");
            transaction.Start();
            GetNewTextStruct(doc, tStructs);
            transaction.Commit();



            return Result.Succeeded;
        }

        private void GetNewTextStruct(Document doc, List<TextStruct> tStructs)
        {
            //遍历找出相交的尺寸文字
            for (int i = 0; i < tStructs.Count - 1; i++)
            {
                for (int j = i + 1; j < tStructs.Count; j++)
                {
                    //检测文字的范围框是否相交
                    if (IsIntersect(tStructs[i], tStructs[j]))
                    {
                        //如果相交则移动尺寸文字，然后重新定义该结构体
                        if (tStructs[j].isDim)//单段尺寸
                        {
                            Dimension dim = tStructs[j].dim;
                            //标线垂直于x轴
                            if (IsVerticalX(tStructs[j]))
                            {
                                dim.TextPosition += tStructs[j].dirY * tStructs[j].height * 1.05;
                            }
                            else //标线垂直于y轴
                            {
                                dim.TextPosition += tStructs[j].dirX * tStructs[j].height * 1.05;
                            }
                            tStructs[j] = GetTextStruct(doc, dim, null, true);
                        }
                        else //尺寸段
                        {
                            Dimension dim = tStructs[j].dim;
                            DimensionSegment dimensionSegment = tStructs[j].dSeg;
                            //标线垂直于x轴
                            if (IsVerticalX(tStructs[j]))
                            {
                                dimensionSegment.TextPosition += tStructs[j].dirY * tStructs[j].height * 1.05;
                            }
                            else//标线垂直于y轴
                            {
                                dimensionSegment.TextPosition += tStructs[j].dirX * tStructs[j].height * 1.05;
                            }
                            tStructs[j] = GetTextStruct(doc, dim, dimensionSegment, false);
                        }

                    }
                }
            }
        }

        private bool IsVerticalX(TextStruct textStruct)
        {
            if (textStruct.dirX.X == 0)
            {
                return true;
            }
            //判断标线是否垂直于X轴
            return false;
        }

        private bool IsIntersect(TextStruct text1, TextStruct text2)
        {
            //遍历标注文字范围框，相交时则标注出现重叠
            foreach (Line line1 in text1.lines)
            {
                foreach (Line line2 in text2.lines)
                {
                    //交点数组
                    IntersectionResultArray resultArray = new IntersectionResultArray();
                    //枚举，用于判断相交类型
                    SetComparisonResult setComparisonResult = line1.Intersect(line2,out resultArray);
                    //Disjoint为不相交
                    if (SetComparisonResult.Disjoint != setComparisonResult)
                    {
                        if (resultArray != null)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
    }
}

