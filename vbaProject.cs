using System;
using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.vbaProject
{
    public class Project
    {
        ProjectProperties PROJECT { get; }

        public string ProjectID { get { return PROJECT.projectID.Value; } set { throw new NotImplementedException(); } }
        public string helpFile { get { return PROJECT.helpFile.Value; } set { throw new NotImplementedException(); } }
        public string exeName32 { get { return PROJECT.exeName32.Value; } set { throw new NotImplementedException(); } }
        public string Name { get { return PROJECT.name.Value; } set { throw new NotImplementedException(); } }
        public string helpID { get { return PROJECT.helpID.Value; } set { throw new NotImplementedException(); } }
        public string Description { get { return PROJECT.description.Value; } set { throw new NotImplementedException(); } }
        public string cMG { get { return PROJECT.cMG.Value; } set { throw new NotImplementedException(); } }
        public string Password { get { throw new NotImplementedException(); } set { throw new NotImplementedException(); } }
        public string projectVisibilityState { get { return PROJECT.projectVisibilityState.Value; } set { throw new NotImplementedException(); } }
        public List<string> hostExtenderRefs { get { return PROJECT.hostExtenderRefs.Select(x => x.Value).ToList(); } }
        public List<ProjectWindow> ProjectWindows { get { return PROJECT.ProjectWindows; } }

        PROJECTwm PROJECTwm { get; }
        PROJECTlk PROJECTlk { get; }

        public List<LicenseInfoRecord> Licenses { get { return PROJECTlk.LicenseInfoRecords; } }

        dir dir { get; }

        public ProjectInformation ProjectInformation { get { return dir.InformationRecord; } }
        public ProjectReferences ProjectReferences { get { return dir.ReferencesRecord; } }
        internal ProjectModules ProjectModulesMeta { get { return dir.ModulesRecord; } }

        _VBA_PROJECT _VBA_PROJECT { get; }
        List<DesignerInformation> vbaForms { get; } = new List<DesignerInformation>();
        public List<Module> Modules { get; } = new List<Module>();
        public Project(System.IO.Stream stream)
        {
            OpenMcdf.CompoundFile cf = new OpenMcdf.CompoundFile(stream);

            OpenMcdf.CFStorage VBA = cf.RootStorage.GetStorage("VBA");
            dir = new dir(VBA.GetStream("dir"));
            var mbcsEncoding = System.Text.Encoding.GetEncoding(dir.InformationRecord.CodePage);

            List<OpenMcdf.CFStream> VBAEntries = new List<OpenMcdf.CFStream>();
            VBA.VisitEntries(delegate (OpenMcdf.CFItem targetNode) { VBAEntries.Add((OpenMcdf.CFStream)targetNode); }, false);
            foreach (OpenMcdf.CFStream item in VBAEntries)
            {
                switch (item.Name)
                {
                    case "dir": break; //Have already read 'dir'
                    case "_VBA_PROJECT": _VBA_PROJECT = new _VBA_PROJECT(item); break;
                    case string s when s.StartsWith("__SRP_"): break;   //Documentation recommends discarding (2.2.6)
                    default:
                        Modules.Add(new Module(item, dir.ModulesRecord.Find(q => q.Name == item.Name), mbcsEncoding));
                        break;
                }
            }

            List<OpenMcdf.CFItem> rootEntries = new List<OpenMcdf.CFItem>();
            cf.RootStorage.VisitEntries(delegate (OpenMcdf.CFItem targetNode) { rootEntries.Add(targetNode); }, false);
            List<OpenMcdf.CFStorage> vbaFormStorage = new List<OpenMcdf.CFStorage>();
            foreach (OpenMcdf.CFItem item in rootEntries)
            {
                if (item.IsStream)
                    switch (item.Name)
                    {
                        case "PROJECT": PROJECT = new ProjectProperties(item, mbcsEncoding); break;
                        case "PROJECTwm": PROJECTwm = new PROJECTwm(item, mbcsEncoding); break;
                        case "PROJECTlk": PROJECTlk = new PROJECTlk(item); break;
                        default: System.Diagnostics.Debug.WriteLine("Unknown Stream in CF Root: " + item.Name); break;
                    }
                if (item.IsStorage && item.Name != "VBA")
                    vbaFormStorage.Add((OpenMcdf.CFStorage)item); //Contains VBAFrame stream encoded in (2.3.5)
                                                                  //  -- VBFrame Name = \u0003VBFrame
                                                                  //Contains ActiveX storage(s)/stream(s) encoded in ([MS-OFORMS] section 2)
            }
            foreach (OpenMcdf.CFStorage container in vbaFormStorage)
            {

                //List<OpenMcdf.CFItem> VBAFormEntries = new List<OpenMcdf.CFItem>();
                //container.VisitEntries(delegate (OpenMcdf.CFItem targetNode) { VBAFormEntries.Add(targetNode); }, false);


                try
                {
                    OpenMcdf.CFStream frameStream = container.GetStream("\u0003VBFrame");
                    DesignerInformation frame = new DesignerInformation(frameStream, mbcsEncoding);

                }
                catch
                {
                    //List<OpenMcdf.CFItem> formItemList = new List<OpenMcdf.CFItem>();
                    //container.VisitEntries(delegate (OpenMcdf.CFItem targetNode) { formItemList.Add(targetNode); }, false);
                    //formItemList.GetType();
                }
            }
        }

    }
}
