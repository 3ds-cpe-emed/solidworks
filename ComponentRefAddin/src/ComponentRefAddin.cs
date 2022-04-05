using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Win32;

namespace dsemedcpe
{
    /// <summary>
    /// Summary description for Addin.
    /// </summary>
    [Guid("41ab69a2-3093-4461-bf9c-5c61fa6c8d95"), ComVisible(true)]
    [SwAddin(Description = "ComponentReference tools", Title = "ComponentReferenceAddin", LoadAtStartup = true)]
    public class ComponentReferenceAddin : ISwAddin
    {
        #region Local Variables

        ISldWorks       m_app = null;
        ICommandManager m_cmd = null;

        int addinID = 0;

        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        //public const int flyoutGroupID = 91;


        // Public Properties
        public ISldWorks App
        {
            get { return m_app; }
        }
        public ICommandManager Cmd
        {
            get { return m_cmd; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute

            SwAddinAttribute SWattr = null;

            Type type = typeof(ComponentReferenceAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }
            #endregion

            try
            {
                RegistryKey hklm = Registry.LocalMachine;
                RegistryKey hkcu = Registry.CurrentUser;
                
                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";

                RegistryKey addinkey = hklm.CreateSubKey(keyname);

                addinkey.SetValue(null, 0);
                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";

                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                RegistryKey hklm = Registry.LocalMachine;
                RegistryKey hkcu = Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ComponentReferenceAddin Implementation
        public ComponentReferenceAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            m_app = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            App.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            m_cmd= App.GetCommandManager(cookie);

            AddCommandMgr();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_cmd);
            m_cmd = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_app);
            m_app = null;

            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;

            if (iBmp == null)
                iBmp = new BitmapHandler();

            Assembly thisAssembly;
          //  int cmdIndex0, cmdIndex1;
            int cmdIndex0;

            string Title = "ComponentReference", ToolTip = "ComponentReference Tools";

            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());

            int cmdGroupErr = 0;

            bool ignorePrevious = false;

            object registryIDs;

            //get the ID information stored in the registry
            bool getDataResult = Cmd.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[2] { mainItemID1, mainItemID2 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = Cmd.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);

            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("dsemedcpe.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("dsemedcpe.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("dsemedcpe.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("dsemedcpe.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            //int menuToolbarOption = (int)swCommandItemType_e.swToolbarItem;

            cmdIndex0 = cmdGroup.AddCommandItem2("SetComponentRefOrder", -1, "Set Component Reference", "Set Component Reference", 0, "SetComponentReference", "EnableSetComponentReference", mainItemID1, menuToolbarOption);
            
            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;

           foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = m_cmd.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = m_cmd.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = m_cmd.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs   = new int[1];
                    int[] TextType = new int[1];

                    cmdIDs[0]   = cmdGroup.get_CommandID(cmdIndex0);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);

                }
            }

            thisAssembly = null;
        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            Cmd.RemoveCommandGroup(mainCmdGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }


        #endregion

        #region UI Callbacks
        // ....
        // Sample inspired on public example
        // http://help.solidworks.com/2020/english/api/sldworksapi/Expand_and_Collapse_FeatureManager_Design_Tree_Nodes_Example_CSharp.htm
        // ....
        public void SetComponentReference()
        {
            #region Local variable declaration
            ModelDoc2 modelDoc            = default(ModelDoc2);
            FeatureManager featureManager = default(FeatureManager);
            TreeControlItem rootNode      = default(TreeControlItem);
            Component2 componentNode      = default(Component2);
            TreeControlItem node          = default(TreeControlItem);
            #endregion

            modelDoc       = (ModelDoc2)App.ActiveDoc;
            featureManager = modelDoc.FeatureManager;
            rootNode       = featureManager.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneBottom);

            node = rootNode.GetFirstChild();

            int i = 0;

            while (node != null)
            {
                if (node.ObjectType == (int)swTreeControlItemType_e.swFeatureManagerItem_Component)
                {
                    componentNode = (Component2) node.Object;

                    //TODO: Check if null
                    //TODO: Consider the following:
                    //-----------------------------
                    // a) Logic to follow in case of Supression?
                    //            suppr = componentNode.GetSuppression();
                    //            switch ((suppr))
                    //            {
                    //                case (int)swComponentSuppressionState_e.swComponentFullyResolved:
                    //                    break;
                    //                case (int)swComponentSuppressionState_e.swComponentLightweight:
                    //                    break;
                    //                case (int)swComponentSuppressionState_e.swComponentSuppressed:
                    //                    break;
                    //            }
                    // b) Logic to follow in case of hidden
                    //            vis = componentNode.Visible;
                    //            switch ((vis))
                    //            {
                    //                case (int)swComponentVisibilityState_e.swComponentHidden:
                    //                    break;
                    //                case (int)swComponentVisibilityState_e.swComponentVisible:
                    //                    break;
                    //            }
                    // c) Solidworks configuration
                    //            refConfigName = componentNode.ReferencedConfiguration;

                    componentNode.ComponentReference = (++i).ToString();
                }

                node = node.GetNext();
            }
        }


        //Enable sequence only if the active document is an assembly document
        public int EnableSetComponentReference()
        {
            IModelDoc2 modelDoc = (IModelDoc2)(m_app.ActiveDoc);

            if (modelDoc.GetType() == ((int)swDocumentTypes_e.swDocASSEMBLY))
                return 1;

            //?
            //if (modelDoc.GetType() == ((int)swDocumentTypes_e.swDocIMPORTED_ASSEMBLY))
            //    return 1;

            return 0;
        }
        #endregion  
    }
}
