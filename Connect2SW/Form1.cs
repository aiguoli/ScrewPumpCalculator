using System;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Connect2SW
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            listBox1.Items.Add("初始化日志...");
        }

        public void WriteLog(string text)
        {
            listBox1.Items.Add(text);
            listBox1.TopIndex = listBox1.Items.Count - 1;
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();

            if (swApp != null)
            {
                string msg = "连接成功，当前SolidWorks版本为" + swApp.RevisionNumber();

                WriteLog(msg);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            if (swApp != null)
            {
                //通过GetDocumentTemplate 获取默认模板的路径 ,第一个参数可以指定类型
                string partDefaultTemplate = swApp.GetDocumentTemplate((int) swDocumentTypes_e.swDocPART, 
                    "", 0, 0, 0);
                var newDoc = swApp.NewDocument(partDefaultTemplate, 0, 0, 0);

                if (newDoc != null)
                {
                    WriteLog("创建完成！");

                    //下面获取当前文件
                    ModelDoc2 swModel = (ModelDoc2) swApp.ActiveDoc;

                    //选择对应的草图基准面
                    bool boolstatus = swModel.Extension.SelectByID2("Plane1", "PLANE", 0, 0, 0, 
                        false, 0, null, 0);

                    //创建一个2d草图
                    swModel.SketchManager.InsertSketch(true);

                    //画一条线 长度100mm  (solidworks 中系统单位是米,所以这里写0.1)
                    //参数:x1,y1,z1,x2,y2,z2
                    swModel.SketchManager.CreateLine(0, 0, 0, 0, 0.1, 0);
                    swModel.SketchManager.CreateLine(0, 0, 0, 0, 0, 0.1);

                    string myNewpartPath = @"C:\Users\Administrator\Desktop\myNewpart.SLDPRT";

                    //保存零件.
                    int longstatus = swModel.SaveAs3(myNewpartPath, 0, 1);
                }
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            ModelDoc2 swModel = (ModelDoc2) swApp.ActiveDoc;
            swModel.SketchManager.InsertSketch(true);
            WriteLog("已关闭");
        }

        private PartDoc partDoc = null;

        private void button6_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();

            if (swApp != null)
            {
                ModelDoc2 swModel = (ModelDoc2) swApp.ActiveDoc;
                partDoc = swModel as PartDoc;
                var swSelMgr = (SelectionMgr) swModel.SelectionManager;
                var selectType = swSelMgr.GetSelectedObjectType3(1, -1);
                partDoc.UserSelectionPostNotify += partDoc_UserSelectionPostNotify;
            }
        }


        private int partDoc_UserSelectionPostNotify()
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();

            ModelDoc2 swModel = (ModelDoc2) swApp.ActiveDoc;

            var swSelMgr = (SelectionMgr) swModel.SelectionManager;

            var selectType = swSelMgr.GetSelectedObjectType3(1, -1);

            WriteLog("You Select :" + Enum.GetName(typeof(swSelectType_e), selectType));

            partDoc.UserSelectionPostNotify -= partDoc_UserSelectionPostNotify;

            return 1;
        }


        private void button9_Click(object sender, EventArgs e)
        {
            ModelDoc2 part;
            Sketch swSketch;
            bool boolstatus;

            ISldWorks swApp = Utility.ConnectToSolidWorks();

            if (swApp != null)
            {
                part = (ModelDoc2) swApp.ActiveDoc;
                boolstatus = part.Extension.SelectByID2("草图1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                part.EditSketch();
                part.ClearSelection2(true);
                swSketch = part.SketchManager.ActiveSketch;
                // 找出齿顶圆齿根圆
                double[] arcs = swSketch.GetArcs2();
                double max = 0;
                double min = 2147483647;
                for (int i = 0; i < swSketch.GetArcCount(); i++)
                {
                    double temp = Math.Sqrt(Math.Pow(arcs[i*16+9], 2)+Math.Pow(arcs[i*16+10], 2)+Math.Pow(arcs[i*16+11], 2)) * 1000;
                    if (temp > max)
                    {
                        max = temp;
                    }

                    if (temp < min)
                    {
                        min = temp;
                    }
                }

                textBox2.Text = max.ToString("0");
                textBox1.Text = min.ToString("0");
                part.SketchManager.InsertSketch(true); // 退出编辑草图
                // 找出质心
                double[] massCenter = Utility.FindMassCenter(part);
                if (massCenter != null)
                {
                    massCoordinates.Text = (massCenter[0] * 1000).ToString("#0.00") + ", " +
                                           (massCenter[1] * 1000).ToString("#0.00") + ", " +
                                           (massCenter[2] * 1000).ToString("#0.00");
                    mass.Text = (massCenter[3] * 1000).ToString("#0.00");
                    
                    // 计算面积利用系数
                    double R = max;
                    double r = min;
                    double theta = Math.Acos((min + max) / (2 * max));
                    double s = 2 * Math.PI * Math.Pow(max, 2) -
                               2 * (theta * Math.Pow(max, 2) - 0.5 * Math.Pow(max, 2) * Math.Sin(2 * theta));
                    // 获取端面面积
                    boolstatus = part.Extension.SelectByID2("", "FACE", 0, 0, 0, false, 0, null, 0);
                    SelectionMgr swSelMgr = part.SelectionManager;
                    Face2 swFace = swSelMgr.GetSelectedObject6(1, -1);
                    double[] properties = part.Extension.GetSectionProperties2(swFace);
                    part.ClearSelection2(true);

                    double a = properties[1] * 1000 * 1000;
                    textBox10.Text = ((s - 2 * a) / s).ToString();
                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            ModelDoc2 part;
            SelectionMgr swSelMgr;
            Feature swFeat;
            bool boolstatus;

            ISldWorks swApp = Utility.ConnectToSolidWorks();

            if (swApp != null)
            {
                part = (ModelDoc2) swApp.ActiveDoc;
                boolstatus = part.Extension.SelectByID2("螺旋线/涡状线1", "REFERENCECURVES", 0, 0, 0, 
                    false, 0, null, 0);
                swSelMgr = part.SelectionManager;
                swFeat = swSelMgr.GetSelectedObject6(1, -1);
                HelixFeatureData FeatData = swFeat.GetDefinition();
                textBox6.Text = ((FeatData.Revolution * FeatData.Pitch) * 1000).ToString();

                double[] massDiameter = Utility.GetMassDiameter(swFeat, part);
                textBox4.Text = massDiameter[0].ToString();
                textBox5.Text = massDiameter[1].ToString();
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            ModelDoc2 part;
            bool boolstatus;
            object myRefPlane;
            object skSegment;
            object myFeature;
            SelectionMgr swSelMgr;
            Face2 swFace;
            Sketch swSketch;
            Feature swFeat;

            ISldWorks swApp = Utility.ConnectToSolidWorks();
            checkBox2.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;
            button4.Enabled = false;
            
            if (swApp != null)
            {
                part = (ModelDoc2) swApp.ActiveDoc;
                // 选中前视端面，不是前视面
                boolstatus = part.Extension.SelectByID2("", "FACE", 0, 0, 0, false, 0, null, 0);
                swSelMgr = part.SelectionManager;
                swFace = swSelMgr.GetSelectedObject6(1, -1);
                double[] properties = part.Extension.GetSectionProperties2(swFace);
                if (properties != null)
                {
                    part.ClearSelection2(true);
                    // 载入螺旋线参数
                    boolstatus = part.Extension.SelectByID2("螺旋线/涡状线1", "REFERENCECURVES", 0, 0, 0, 
                        false, 0, null, 0);
                    swSelMgr = part.SelectionManager;
                    swFeat = swSelMgr.GetSelectedObject6(1, -1);
                    HelixFeatureData FeatData = swFeat.GetDefinition();
                    double partLength = FeatData.Revolution * FeatData.Pitch; // 圈数 * 螺距
                    bool clockwise = FeatData.Clockwise;
                    part.ClearSelection2(true);

                    // 选定前视基准面画梯形
                    double A0 = properties[1];
                    double mass_x = properties[2];
                    double mass_y = properties[3];
                    double r0 = Math.Sqrt(Math.Pow(mass_x, 2) + Math.Pow(mass_y, 2));
                    double r = Convert.ToDouble((Convert.ToInt32(textBox1.Text) / 10 + 1) * 10) / 1000; // 大于齿根圆取整
                    // double R = Convert.ToDouble((Convert.ToInt32(textBox2.Text) / 10) * 10) / 1000; // 小于齿顶圆取整
                    double R = (Convert.ToDouble(textBox2.Text) - 5) / 1000; // 至少比齿顶圆小5mm
                    double theta = Math.Asin(A0 * r0 * 3 / 4 / (Math.Pow(R, 3) - Math.Pow(r, 3))) * 180 /
                                   Math.PI; // 扇形圆心角的一半
                                        
                    // 检查用户输入
                    if (textBox7.Text != "")
                    {
                        r = (Convert.ToDouble(textBox7.Text)) / 1000;
                    }
                    if (textBox8.Text != "")
                    {
                        R = (Convert.ToDouble(textBox8.Text)) / 1000;
                    }

                    if (textBox9.Text != "")
                    {
                        theta = Convert.ToDouble(textBox9.Text) / 2;
                    }


                    double fi = Math.Atan(mass_y / mass_x) * 180 / Math.PI; // 端面质心角度
                    double theta_1 = (fi + theta - 360) * Math.PI / 180; // 扇形弧度制角度1
                    double theta_2 = (fi - theta) * Math.PI / 180; // 扇形弧度制角度2
                    boolstatus = part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, true, 0, null, 0);
                    part.SketchManager.InsertSketch(true);
                    part.ClearSelection2(true);
                    swSketch = part.SketchManager.ActiveSketch;
                    part.SketchManager.CreateCircleByRadius(0, 0, 0, r);
                    part.SketchManager.CreateCircleByRadius(0, 0, 0, R);
                    part.SketchManager.CreateLine(r * Math.Cos(theta_1), r * Math.Sin(theta_1), 0,
                        R * Math.Cos(theta_1), R * Math.Sin(theta_1), 0);
                    part.SketchManager.CreateLine(r * Math.Cos(theta_2), r * Math.Sin(theta_2), 0,
                        R * Math.Cos(theta_2), R * Math.Sin(theta_2), 0);
                    part.ClearSelection2(true);
                    part.SketchManager.InsertSketch(true);

                    // 裁剪，目前没有找到很好的裁剪方式
                    part.EditSketch();
                    part.ClearSelection2(true);
                    // 鬼知道为什么X,Y无论怎么输入都切不到，猜测是小数点问题 这个api开发者就该浸猪笼
                    // part.SetPickMode();
                    // boolstatus = part.Extension.SelectByID2("圆弧1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                    // boolstatus = part.SketchManager.SketchTrim(1, -0.029938641780916, 1.17041772497828E-04, 0);
                    // part.SetPickMode();
                    // boolstatus = part.Extension.SelectByID2("圆弧2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                    // boolstatus = part.SketchManager.SketchTrim(1, 0, 0, 0);

                    // 莓0.5s检查是否已经切除
                    if (checkBox2.Checked)
                    {
                        WriteLog("选择裁剪工具切掉多余部分，然后点击已切除");
                    }
                    while (!checkBox1.Checked)
                    {
                        Utility.Delay(500);
                    }

                    checkBox1.Checked = false;
                    part.ClearSelection2(true);
                    part.SketchManager.InsertSketch(true);
                    string sketchName = Utility.GetLatestFeatureName(part);
                    // 螺旋挖槽
                    boolstatus = part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, false, 0, null, 0);
                    part.SketchManager.InsertSketch(true);
                    part.ClearSelection2(true);

                    skSegment = part.SketchManager.CreateCircleByRadius(0, 0, 0, (R + r) / 2);
                    part.ClearSelection2(true);
                    part.SketchManager.InsertSketch(true);
                    part.InsertHelix(false, clockwise, false, false, 0, 0.033,
                        FeatData.Pitch, 0.5, 0, theta_1);
                    part.ClearSelection2(true);
                    string helixName = Utility.GetLatestFeatureName(part);
                    boolstatus = part.Extension.SelectByID2(sketchName, "SKETCH", 0, 0, 0, true, 
                        0, null, 0);
                    boolstatus = part.Extension.SelectByID2(helixName, "REFERENCECURVES", 5.07263217039764E-02,
                        -1.10840138118764E-02, -9.19322465863956E-02, true, 0, null, 0);
                    part.ClearSelection2(true);
                    boolstatus = part.Extension.SelectByID2(sketchName, "SKETCH", 0, 0, 0, false, 
                        1, null, 0);
                    boolstatus = part.Extension.SelectByID2(helixName, "REFERENCECURVES", 5.07263217039764E-02,
                        -1.10840138118764E-02, -9.19322465863956E-02, true, 4, null, 0);
                    myFeature = part.FeatureManager.InsertCutSwept4(false, true, 0,
                        false, false, 0, 0, false, 
                        0, 0, 0,
                        10, true, true, 0, true, 
                        true, true, false);
                    part.ClearSelection2(true);

                    // 建另一端参考面
                    boolstatus = part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, true, 
                        0, null, 0);
                    myRefPlane = part.FeatureManager.InsertRefPlane(8, partLength, 
                        0, 0, 0, 0);
                    string anotherPlaneName = Utility.GetLatestFeatureName(part);
                    part.ClearSelection2(true);
                    // 在端面Ⅱ上画圆
                    boolstatus = part.Extension.SelectByID2(Utility.GetLatestFeatureName(part), "PLANE", 0, 0, 0, true,
                        0, null, 0);
                    part.SketchManager.InsertSketch(true);
                    part.ClearSelection2(true);
                    swSketch = part.SketchManager.ActiveSketch;
                    part.SketchManager.CreateCircleByRadius(0, 0, 0, r);
                    part.SketchManager.CreateCircleByRadius(0, 0, 0, R);
                    // part.SketchManager.CreateLine(-r * Math.Cos(theta_1), -r * Math.Sin(theta_1), 0,
                    //     -R * Math.Cos(theta_1), -R * Math.Sin(theta_1), 0);
                    // part.SketchManager.CreateLine(-r * Math.Cos(theta_2), -r * Math.Sin(theta_2), 0,
                    //     -R * Math.Cos(theta_2), -R * Math.Sin(theta_2), 0);
                    part.SketchManager.CreateLine(r * Math.Cos(theta_1+FeatData.Revolution*Math.PI*2), r * Math.Sin(theta_1+FeatData.Revolution*Math.PI*2), 0,
                        R * Math.Cos(theta_1+FeatData.Revolution*Math.PI*2), R * Math.Sin(theta_1+FeatData.Revolution*Math.PI*2), 0);
                    part.SketchManager.CreateLine(r * Math.Cos(theta_2+FeatData.Revolution*Math.PI*2), r * Math.Sin(theta_2+FeatData.Revolution*Math.PI*2), 0,
                        R * Math.Cos(theta_2+FeatData.Revolution*Math.PI*2), R * Math.Sin(theta_2+FeatData.Revolution*Math.PI*2), 0);

                    // 莓0.5s检查是否已经切除
                    if (checkBox2.Checked)
                    {
                        WriteLog("选择裁剪工具切掉多余部分，然后点击已切除");
                    }
                    while (!checkBox3.Checked)
                    {
                        Utility.Delay(500);
                    }

                    part.InsertSketch2(true);
                    checkBox3.Checked = false;
                    sketchName = Utility.GetLatestFeatureName(part);
                    // 螺旋挖槽
                    boolstatus = part.Extension.SelectByID2(anotherPlaneName, "PLANE", 0, 0, 0, false, 
                        0, null, 0);
                    part.InsertSketch2(true);
                    part.ClearSelection2(true);
                    skSegment = part.SketchManager.CreateCircleByRadius(0, 0, 0, (R + r) / 2);
                    part.ClearSelection2(true);
                    part.InsertSketch2(true);
                    
                    // 插入螺旋线
                    // part.InsertHelix(true, false, false, false, 0, 0.033, 
                    //     FeatData.Pitch, 0.5, 0, 2.6043007468976);
                    part.InsertHelix(true, !clockwise, false, false, 0, 0.033, 
                        FeatData.Pitch, 0.5, 0, theta_2+FeatData.Revolution*Math.PI*2);
                    part.ClearSelection2(true);
                    helixName = Utility.GetLatestFeatureName(part);
                    boolstatus = part.Extension.SelectByID2(sketchName, "SKETCH", 0, 0, 0, false, 
                        1, null, 0);
                    boolstatus = part.Extension.SelectByID2(helixName, "REFERENCECURVES", 5.07263217039764E-02,
                        -1.10840138118764E-02, -9.19322465863956E-02, true, 4, null, 0);
                    myFeature = part.FeatureManager.InsertCutSwept4(false, true, 0, 
                        false, false, 0, 0, false, 0, 0, 0,
                        10, true, true, 0, true, 
                        true, true, false);
                    part.ClearSelection2(true);
                }
            }

            checkBox2.Enabled = true;
            textBox7.Enabled = true;
            textBox8.Enabled = true;
            button4.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ModelDoc2 part;
            Sketch swSketch;
            bool boolstatus;

            ISldWorks swApp = Utility.ConnectToSolidWorks();

            if (swApp != null)
            {
                part = (ModelDoc2) swApp.ActiveDoc;
                boolstatus = part.Extension.SelectByID2("草图1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                part.EditSketch();
                part.ClearSelection2(true);
                swSketch = part.SketchManager.ActiveSketch;
                // 找出齿顶圆齿根圆
                double[] arcs = swSketch.GetArcs2();
                for (int i = 0; i < swSketch.GetArcCount()*16; i++)
                {
                    WriteLog(arcs[i].ToString());
                }

                part.SketchManager.InsertSketch(true);
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            var swModel = (ModelDoc2)swApp.ActiveDoc;

            swModel = (ModelDoc2)swApp.ActiveDoc;

            swModel.ReloadOrReplace(false, swModel.GetPathName(), true);
            WriteLog("重启零件");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            ModelDoc2 part = (ModelDoc2) swApp.ActiveDoc;
            Boolean boolstatus = part.Extension.SelectByID2("", "FACE", 0, 0, 0, false, 0, null, 0);
        }
    }
}