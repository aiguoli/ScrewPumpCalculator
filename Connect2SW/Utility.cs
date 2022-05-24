using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Connect2SW
{
    public class Utility
    {
        public static ISldWorks SwApp { get; private set; }

        
        public static void Delay(int milliSecond)
        {
            int start = System.Environment.TickCount;
            while (Math.Abs(System.Environment.TickCount - start) < milliSecond)
            {
                Application.DoEvents();
            }
        }

        public static ISldWorks ConnectToSolidWorks()
        {
            if (SwApp != null)
            {
                return SwApp;
            }
            else
            {
                Debug.Print("connect to solidworks on " + DateTime.Now);
                try
                {
                    SwApp = (SldWorks) Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (COMException)
                {
                    try
                    {
                        SwApp = (SldWorks) Marshal.GetActiveObject("SldWorks.Application.23"); //2015
                    }
                    catch (COMException)
                    {
                        try
                        {
                            SwApp = (SldWorks) Marshal.GetActiveObject("SldWorks.Application.26"); //2018
                        }
                        catch (COMException)
                        {
                            MessageBox.Show("Could not connect to SolidWorks.", "SolidWorks", MessageBoxButtons.OK,
                                MessageBoxIcon.Hand);
                            SwApp = null;
                        }
                    }
                }

                return SwApp;
            }
        }

        public static void DrawSplineCurve(string x, string y, ModelDoc2 swModel)
        {
            bool boolstatus = swModel.Extension.SelectByID2(
                "Plane1",
                "PLANE",
                0,
                0,
                0,
                false,
                0,
                null,
                0
            );
            SketchSpline equationDriveCurve = swModel.SketchManager.CreateEquationSpline2(
                x,
                y,
                "",
                "0",
                "pi",
                false,
                0,
                0,
                0,
                true,
                true
            );
        }

        public static double[] GetMassDiameter(Feature swFeat, ModelDoc2 part)
        {
            HelixFeatureData FeatData = swFeat.GetDefinition();
            double circleCount = FeatData.Revolution; // 螺纹圈数

            double x, y, m, massX = 0, massY = 0;
            double[] res = new double[2];
            // 先算Ⅰ面
            for (int i = 1; i < 11; i++)
            {
                // 缩放实体达到切割的效果
                FeatData.Revolution = circleCount * i / 10;
                swFeat.ModifyDefinition(FeatData, part, null);
                Thread.Sleep(2000);

                double[] massCenter = FindMassCenter(part);
                x = massCenter[0] * 1000;
                y = massCenter[1] * 1000;
                m = massCenter[3] * 1000;
                double r = Math.Sqrt(Math.Pow(y, 2) + Math.Pow(x, 2));
                double theta = Math.Atan(y / x) / Math.PI * 180;
                massX += (m * r * i * Math.Sin(theta)) / 10;
                massY += (m * r * i * Math.Cos(theta)) / 10;
            }

            res[0] = Math.Sqrt(Math.Pow(massX, 2) + Math.Pow(massY, 2));
            massX = 0;
            massY = 0;
            // 再算Ⅱ面
            for (int j = 10; j > 0; j--)
            {
                // 缩放实体达到切割的效果
                FeatData.Revolution = circleCount * j / 10;
                swFeat.ModifyDefinition(FeatData, part, null);
                Thread.Sleep(2000);
                
                double[] massCenter = FindMassCenter(part);
                x = massCenter[0] * 1000;
                y = massCenter[1] * 1000;
                m = massCenter[3] * 1000;
                double r = Math.Sqrt(Math.Pow(y, 2) + Math.Pow(x, 2));
                double theta = Math.Abs(Math.Atan(y / x) / Math.PI * 180);
                massX += (m * r * j * Math.Sin(theta)) / 10;
                massY += (m * r * j * Math.Cos(theta)) / 10;
            }
            res[1] = Math.Sqrt(Math.Pow(massX, 2) + Math.Pow(massY, 2));
            // 还原零件圈数
            FeatData.Revolution = circleCount;
            swFeat.ModifyDefinition(FeatData, part, null);
            return res;
        }

        public static double[] FindMassCenter(ModelDoc2 part)
        {
            int massStatus = 0;
            double[] massProperties;
            double[] res = new double[4];
            massProperties = (double[])part.Extension.GetMassProperties(1, ref massStatus);
            if (massProperties != null)
            {
                res[0] = massProperties[0];
                res[1] = massProperties[1];
                res[2] = massProperties[2];
                res[3] = massProperties[5];
            }

            return res;
        }

        public static string GetLatestFeatureName(ModelDoc2 part)
        {
            Feature swFeat = part.FirstFeature();
            string res = null;
            while (swFeat != null)
            {
                res = swFeat.Name;
                swFeat = swFeat.GetNextFeature();
            }
            return res;
        }
    }
}