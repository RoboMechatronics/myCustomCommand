using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.IO;
using System.Windows.Forms;
using System.Reflection;

namespace MyCustomCommand
{
    public class MyCustomCommand : SwAddin
    {
        #region Private Members

        private SldWorks swApp = null;
        private int swCookie;
        CommandManager swCommandManager;
        private int swCommandGroupID = 0;

        #endregion Private Members

        struct Point
        {
            public double X;
            public double Y;
            public double Z;
        }

        #region Solidworks Add-In Callbacks
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            #region Local Variables
            string swCommandGroupTitle = "My Command Group";
            string swCommandGroupToolTip = "My Command Group Title";
            string swCommandGroupHint = "My Command Group Hint";
            int Error = 0;

            string TabName = "My Tab";
            
            int[] swCommandID = { 0, 1, 2 };
            int[] docTypes = { (int)swDocumentTypes_e.swDocASSEMBLY,
                                (int)swDocumentTypes_e.swDocPART,
                                (int)swDocumentTypes_e.swDocDRAWING};

            #endregion

            swApp = (SldWorks)ThisSW;
            swCookie = Cookie;

            _ = swApp.SetAddinCallbackInfo2(0, this, swCookie);

            swCommandManager = swApp.GetCommandManager(swCookie);

            #region Add Commands

            CommandGroup swCommandGroup = swCommandManager.CreateCommandGroup2(swCommandGroupID, // UserID
                                                                swCommandGroupTitle, // Title
                                                                swCommandGroupToolTip, // Tool tip
                                                                swCommandGroupHint, // Hint
                                                                -1, // Position
                                                                true, // Ignore previous version
                                                                Error); // out variable

            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string directory = Directory.GetParent(assemblyFolder).Parent.FullName + "\\Resources\\";

            string[] IconLists = {  directory + "ComboIcon20x20.png",
                                    directory + "ComboIcon32x32.png",
                                    directory + "ComboIcon40x40.png",
                                    directory + "ComboIcon64x64.png",
                                    directory + "ComboIcon96x96.png",
                                    directory + "ComboIcon128x128.png"};

            string[] MainIconLists = {directory + "icon20x20.png",
                                      directory + "icon32x32.png",
                                      directory + "icon40x40.png"};

            swCommandGroup.IconList = IconLists;
            swCommandGroup.MainIconList = MainIconLists;
            
            string Name;
            int Position;
            string HintString;
            string ToolTip;
            int ImageListIndex;
            string CallbackFunction;
            string EnableMethod;
            int MenuTBOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

            Name = "CreateRectangleBox";
            Position = 0;
            HintString = "Create a Rectangle Box";
            ToolTip = "Rectangle Box";
            ImageListIndex = 0;
            CallbackFunction = "CreateRectangleBox()";
            EnableMethod = "EnableCreateRectangleBox()";
            
            int swCommandIndex0 = swCommandGroup.AddCommandItem2(Name, Position, HintString, ToolTip, ImageListIndex, CallbackFunction, EnableMethod, swCommandID[0], MenuTBOption);

            Name = "CreateCylinder";
            Position = 1;
            HintString = "Create a cylinder";
            ToolTip = "Cylinder";
            ImageListIndex = 1;
            CallbackFunction = "CreateCylinder()";
            EnableMethod = "EnableCreateCylinder()";

            int swCommandIndex1 = swCommandGroup.AddCommandItem2(Name, Position, HintString, ToolTip, ImageListIndex, CallbackFunction, EnableMethod, swCommandID[1], MenuTBOption);

            Name = "CreateSphere";
            Position = 2;
            HintString = "Create a sphere";
            ToolTip = "Sphere";
            ImageListIndex = 2;
            CallbackFunction = "CreateSphere()";
            EnableMethod = "EnableCreateSphere()";

            int swCommandIndex2 = swCommandGroup.AddCommandItem2(Name, Position, HintString, ToolTip, ImageListIndex, CallbackFunction, EnableMethod, swCommandID[2], MenuTBOption);

            swCommandGroup.HasToolbar = true;
            swCommandGroup.HasMenu = true;
            swCommandGroup.Activate();

            foreach(int docType in docTypes)
            {
                CommandTab swCommandTab;
                swCommandTab  = swCommandManager.GetCommandTab(docType, TabName);

                while (swCommandTab != null)
                {
                    swCommandManager.RemoveCommandTab(swCommandTab);
                    swCommandTab = swCommandManager.GetCommandTab(docType, TabName);
                }

                if (swCommandTab == null)
                {
                    swCommandTab = swCommandManager.AddCommandTab(docType, TabName);
                    CommandTabBox swCommandBox = swCommandTab.AddCommandTabBox();

                    int[] swCommandIDs = new int[3];
                    int[] swTextTypes = new int[3];

                    swCommandIDs[0] = swCommandGroup.get_CommandID(swCommandIndex0);
                    swTextTypes[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    swCommandIDs[1] = swCommandGroup.get_CommandID(swCommandIndex1);
                    swTextTypes[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    swCommandIDs[2] = swCommandGroup.get_CommandID(swCommandIndex2);
                    swTextTypes[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    _ = swCommandBox.AddCommands(swCommandIDs, swTextTypes);
                }
            }
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            swCommandManager.RemoveCommandGroup2(swCommandGroupID, false);
            swCommandManager = null;
            swApp = null;
            return true;
        }
        #endregion Solidworks Add-In callbacks

        #region Create Rectangle Box
        public void CreateRectangleBox()
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            SketchManager swSketchManager = (SketchManager)swModel.SketchManager;
            FeatureManager swFeatureManager = (FeatureManager)swModel.FeatureManager;
            swSketchManager.InsertSketch(true);
            swSketchManager.AddToDB = true;

            double width_mm = 40.0;
            double height_mm = 40.0;

            Point Point1, Point2;
            Point1.X = 0;
            Point1.Y = 0;
            Point1.Z = 0;
            Point2.X = width_mm / 2 / 1000; // milimiter to meter
            Point2.Y = height_mm / 2/ 1000; // milimiter to meter
            Point2.Z = 0;

            _ = swSketchManager.CreateCenterRectangle(Point1.X, Point1.Y, Point1.Z, Point2.X, Point2.Y, Point2.Z);

            swModel.ClearSelection2(true);
            swSketchManager.AddToDB = true;

            bool Sd = true;
            bool Flip = false;
            bool Dir = true;
            int T1 = (int)swEndConditions_e.swEndCondBlind;
            int T2 = (int)swEndConditions_e.swEndCondBlind;
            double D1 = 30.0 / 1000;
            double D2 = 0;
            bool Dchk1 = false;
            bool Dchk2 = false;
            bool Ddir1 = false;
            bool Ddir2 = false;
            double Dang1 = 0;
            double Dang2 = 0;
            bool OffsetReverse1 = false;
            bool OffsetReverse2 = false;
            bool TranslateSurface1 = false;
            bool TranslateSurface2 = false;
            bool Merge = true;
            bool UseFeatScope = true;
            bool UseAutoSelect = true;
            int T0 = 0;
            double StartOffset = 0;
            bool FlipStartOffset = false;
            _ = swFeatureManager.FeatureExtrusion3(Sd, Flip, Dir, T1, T2, D1, D2, Dchk1, Dchk2, Ddir1, Ddir2, Dang1, Dang2, OffsetReverse1, OffsetReverse2, TranslateSurface1, TranslateSurface2, Merge, UseFeatScope, UseAutoSelect, T0, StartOffset, FlipStartOffset);
            
            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(true);
        }
        #endregion

        #region Create Cylinder
        public void CreateCylinder()
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            SketchManager swSketchManager = (SketchManager)swModel.SketchManager;
            FeatureManager swFeatureManager = (FeatureManager)swModel.FeatureManager;
            ModelDocExtension swModelDocExtension = (ModelDocExtension)swModel.Extension;

            swSketchManager.InsertSketch(true);
            swSketchManager.AddToDB = true;

            Point Point1, Point2;
            Point1.X = 0; 
            Point1.Y = 0; 
            Point1.Z = 0; // milimeter to meter

            Point2.X = 10.0/2/1000; 
            Point2.Y = 20.0/2/1000; 
            Point2.Z = 0; // milimeter to meter

            swSketchManager.CreateCornerRectangle(Point1.X, Point1.Y, Point1.Z, Point2.X, Point2.Y, Point2.Z);

            swModel.ClearSelection2(true);
            
            _ = swModelDocExtension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, true, 16, null, (int)swSelectOption_e.swSelectOptionDefault);
            swFeatureManager = swModel.FeatureManager;

            bool SingleDir = true;
            bool IsSolid = true;
            bool IsThin = false;
            bool IsCut = false;
            bool ReverseDir = false;
            bool BothDirectionUpToSameEntity = false;
            int Dir1Type = (int)swEndConditions_e.swEndCondBlind;
            int Dir2Type = (int)swEndConditions_e.swEndCondBlind;
            double Dir1Angle = 360.0*Math.PI/180;  // Degree to Radian
            double Dir2Angle = 0;
            bool OffsetReverse1 = false;
            bool OffsetReverse2 = false;
            double OffsetDistance1 = 0;
            double OffsetDistance2 = 0;
            int ThinType = 0;
            double ThinThickness1 = 0;
            double ThinThickness2 = 0;
            bool Merge = true;
            bool UseFeatScope = true;
            bool UseAutoSelect = true;

            swFeatureManager.FeatureRevolve2(SingleDir, 
                                            IsSolid, 
                                            IsThin, 
                                            IsCut, 
                                            ReverseDir, 
                                            BothDirectionUpToSameEntity, 
                                            Dir1Type, 
                                            Dir2Type,
                                            Dir1Angle, 
                                            Dir2Angle, 
                                            OffsetReverse1, 
                                            OffsetReverse2, 
                                            OffsetDistance1, 
                                            OffsetDistance2, 
                                            ThinType, 
                                            ThinThickness1, 
                                            ThinThickness2, 
                                            Merge, 
                                            UseFeatScope, 
                                            UseAutoSelect);

            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(true);
        }
        #endregion #region Create Cylinder

        #region #region Create Sphere
        public void CreateSphere()
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            SketchManager swSketchManager = (SketchManager)swModel.SketchManager;
            FeatureManager swFeatureManager = (FeatureManager)swModel.FeatureManager;
            ModelDocExtension swModelDocExtension = (ModelDocExtension)swModel.Extension;

            swSketchManager.InsertSketch(true);
            swSketchManager.AddToDB = true;

            Point Point1, Point2;

            Point1.X = 0; Point1.Y = 0; Point1.Z = 0;  // milimeter to meter
            Point2.X = 0; Point2.Y = 30.0 / 1000; Point2.Z = 0;
            _ = swSketchManager.CreateLine(Point1.X, Point1.Y, Point1.Z, Point2.X, Point2.Y, Point2.Z);
            swModel.ClearSelection2(true);

            Point1.X = 0; Point1.Y = Point2.Y; Point1.Z = 0;  // milimeter to meter
            Point2.X = 0; Point2.Y = 0; Point2.Z = 0;
            _ = swSketchManager.CreateTangentArc(Point1.X, Point1.Y, Point1.Z, Point2.X, Point2.Y, Point2.Z, 4);
            swModel.ClearSelection2(true);

            _ = swModelDocExtension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, true, 16, null, (int)swSelectOption_e.swSelectOptionDefault);
            swFeatureManager = swModel.FeatureManager;

            bool SingleDir = true;
            bool IsSolid = true;
            bool IsThin = false;
            bool IsCut = false;
            bool ReverseDir = false;
            bool BothDirectionUpToSameEntity = false;
            int Dir1Type = (int)swEndConditions_e.swEndCondBlind;
            int Dir2Type = (int)swEndConditions_e.swEndCondBlind;
            double Dir1Angle = 360.0 * Math.PI / 180;  // Degree to Radian
            double Dir2Angle = 0;
            bool OffsetReverse1 = false;
            bool OffsetReverse2 = false;
            double OffsetDistance1 = 0;
            double OffsetDistance2 = 0;
            int ThinType = 0;
            double ThinThickness1 = 0;
            double ThinThickness2 = 0;
            bool Merge = true;
            bool UseFeatScope = true;
            bool UseAutoSelect = true;

            swFeatureManager.FeatureRevolve2(SingleDir,
                                            IsSolid,
                                            IsThin,
                                            IsCut,
                                            ReverseDir,
                                            BothDirectionUpToSameEntity,
                                            Dir1Type,
                                            Dir2Type,
                                            Dir1Angle,
                                            Dir2Angle,
                                            OffsetReverse1,
                                            OffsetReverse2,
                                            OffsetDistance1,
                                            OffsetDistance2,
                                            ThinType,
                                            ThinThickness1,
                                            ThinThickness2,
                                            Merge,
                                            UseFeatScope,
                                            UseAutoSelect);

            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(true);
        }

        #endregion #region Create Sphere

        #region Enable Method
        /// <summary>
        /// Optional function that controls the state of the item; if specified, then SOLIDWORKS calls this function before displaying the item
        /// if return 1, the Solidworks deselects and enables the item; this is the default state if no update function is specified.
        /// See detail about AddCommandItem2() method here: https://help.solidworks.com/2023/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icommandgroup~addcommanditem2.html
        /// </summary>
        /// <returns></returns>

        public int EnableCreateRectangleBox()
        {
            return 1;
        }
        public int EnableCreateCylinder()
        {
            return 1;
        }
        public int EnableCreateSphere()
        {
            return 1;
        }
        #endregion Enable Method

        #region COM Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            RegistryKey mLocalMachine = Registry.LocalMachine;

            string subKey = "SOFTWARE\\SOLIDWORKS\\ADDINS\\{" + t.GUID.ToString() + "}";
            RegistryKey mRegistryKey = mLocalMachine.CreateSubKey(subKey);

            mRegistryKey.SetValue(null, 0);
            mRegistryKey.SetValue("Description", "My Solidworks Add-In");
            mRegistryKey.SetValue("Title", "My Custom Command");
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            RegistryKey mLocalMachine = Registry.LocalMachine;

            string subkey = "SOFTWARE\\SOLIDWORKS\\ADDINS\\{" + t.GUID.ToString() + "}";
            mLocalMachine.DeleteSubKey(subkey);
            mLocalMachine = null;
        }
        #endregion COM Registration
    }
}
