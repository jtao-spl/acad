using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;

namespace acad
{
    public static partial class AcdBaseTool
    {
        public static double[,] sizeFieldTolseranceLevelValue = new double[4, 8] {
                {0.05,0.05,0.1,0.15,0.2,0.3,0.5,-1 },
                {0.1,0.1,0.2,0.3,0.5,0.8,1.2,2 },
                {0.2,0.3,0.5,0.8,1.2,2,3,4 },
                {-1,0.5,1,1.5,2.5,4,6,8 }
            };
        public static  double GetRotatedDimensionSize(this RotatedDimension rDimension)
        {
            double DistanceX = rDimension.XLine1Point.X - rDimension.XLine2Point.X;
            double DistanceY = rDimension.XLine1Point.Y - rDimension.XLine2Point.Y;
            return rDimension.XLine1Point.GetDistanceBetweenTowPoint(rDimension.XLine2Point);
        }
        public static double GetDistanceBetweenTowPoint(this Point3d point1, Point3d point2)
        {
            return Math.Sqrt(Math.Pow(point1.X - point2.X, 2) + Math.Pow(point1.Y - point2.Y, 2) + Math.Pow(point1.Z - point2.Z, 2));
        }
        public static ObjectId AddEntityToModelSpace(this Database db,Entity ent)
        {
            ObjectId entId = ObjectId.Null;
            using(Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                entId = btr.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);
                trans.Commit();
            }
            return entId;
        }
        
        public static decimal CalculateDeltaByToleranceLevelAndSizeField(ToleraceLevel tolerace, int sizeField)
        {
            return Convert.ToDecimal(sizeFieldTolseranceLevelValue[(int)tolerace, sizeField]);
        }
        public static int getSizeField(this decimal baseValue)
        {   

            if (baseValue > 0.5m && baseValue <= 3m) return 0;
            else if (baseValue > 3m && baseValue <= 6m) return 1;
            else if (baseValue > 6m && baseValue <= 30m) return 2;
            else if (baseValue > 30m && baseValue <= 120m) return 3;
            else if (baseValue > 120m && baseValue <= 400m) return 4;
            else if (baseValue > 400m && baseValue <= 1000m) return 5;
            else if (baseValue > 1000m && baseValue <= 2000m) return 6;
            else if (baseValue > 2000m && baseValue <= 4000m) return 7;
            return 8;
        }
    }
}
