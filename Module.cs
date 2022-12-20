using System;
using System.Collections.Generic;

namespace Edmosoft.Office.vbaProject
{
    public class Module
    {
        public ModuleHeader Header;
        ProjectModule ModuleMeta;

        public string Name { get { return ModuleMeta.Name; } }
        public string Body;
        internal Module(OpenMcdf.CFItem item, ProjectModule moduleMeta, System.Text.Encoding mbcsEncoding) : this((OpenMcdf.CFStream)item, moduleMeta, mbcsEncoding) { }
        internal Module(OpenMcdf.CFStream cFStream, ProjectModule moduleMeta, System.Text.Encoding mbcsEncoding) : this(cFStream.GetData(), moduleMeta, mbcsEncoding) { }
        internal Module(byte[] bytes, ProjectModule moduleMeta, System.Text.Encoding mbcsEncoding) : this(new System.IO.MemoryStream(bytes), moduleMeta, mbcsEncoding) { }
        internal Module(System.IO.Stream stream, ProjectModule moduleMeta, System.Text.Encoding mbcsEncoding) : this(new Edmosoft.IO.StreamReader(stream), moduleMeta, mbcsEncoding) { }
        internal Module(Edmosoft.IO.StreamReader streamReader, ProjectModule moduleMeta, System.Text.Encoding mbcsEncoding)
        {
            ModuleMeta = moduleMeta;
            Header = new ModuleHeader() { ModuleType = moduleMeta.Type };
            byte[] PerformanceCache = streamReader.ReadBlock((int)moduleMeta.Offset);
            System.IO.Stream SourceCodeStream = Edmosoft.IO.vbaCompression.vbaStreamReader.Decode(streamReader);

            System.IO.StreamReader sourceReader = new System.IO.StreamReader(SourceCodeStream, mbcsEncoding);
            switch (moduleMeta.Type)
            {
                case ProjectModule.ModuleType.procedural:
                    ProjectProperty abnf1 = new ProjectProperty(sourceReader.ReadLine());
                    if (abnf1.Name.Split(new char[] { ' ' }, 2)[0] == "Attribute")
                        if (abnf1.Name.Split(new char[] { ' ' }, 2)[1].Trim() == "VB_Name")
                            Header.Name = abnf1.Value;
                    break;
                case ProjectModule.ModuleType.@class:
                    string line = sourceReader.ReadLine();
                    do
                    {
                        ProjectProperty abnf2 = new ProjectProperty(line);
                        if (abnf2.Name.Split(new char[] { ' ' }, 2)[0] == "Attribute")
                            switch (abnf2.Name.Split(new char[] { ' ' }, 2)[1])
                            {
                                case "VB_Name":
                                    Header.Name = abnf2.Value.Trim();
                                    break;
                                case "VB_Base":
                                    Header.Base = abnf2.Value.Trim();
                                    break;
                                case "VB_GlobalNameSpace":
                                    Header.GlobalNameSpace = abnf2.Value.Trim() == "True";
                                    break;
                                case "VB_Creatable":
                                    Header.Creatable = abnf2.Value.Trim() == "True";
                                    break;
                                case "VB_PredeclaredId":
                                    Header.PredeclareId = abnf2.Value.Trim() == "True";
                                    break;
                                case "VB_Exposed":
                                    Header.Exposed = abnf2.Value.Trim() == "True";
                                    break;
                                case "VB_TemplateDerived":
                                    Header.TemplateDerived = abnf2.Value.Trim() == "True";
                                    break;
                                case "VB_Customizable":
                                    Header.Customizable = abnf2.Value.Trim() == "True";
                                    break;
                                case "VB_Control":
                                    Header.Control = abnf2.Value.Trim();
                                    break;
                                default:
                                    Console.WriteLine(string.Format("Unknown VBA Attribute {0}", abnf2.Name.Split(new char[] { ' ' }, 2)[1]));
                                    break;
                            }
                        line = sourceReader.ReadLine();
                    } while (line != null && line.StartsWith("Attribute"));
                    Body = line + "\r\n";
                    break;
                default:
                    throw new FormatException("Invalid Module Type Record");
            }
            Body += sourceReader.ReadToEnd();
        }

        public class ModuleHeader
        {
            internal ProjectModule.ModuleType ModuleType;
            public string Name;
            public string Base;
            public bool GlobalNameSpace;
            public bool Creatable;
            public bool PredeclareId;
            public bool Exposed;
            public bool TemplateDerived;
            public bool Customizable;
            public string Control;
        }
    }
}
