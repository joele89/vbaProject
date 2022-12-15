using System;
using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.vbaProject
{
    public class DesignerInformation
    {
        string CLSID;
        string DesignerName;
        List<ProjectProperty> properties;
        public DesignerInformation(OpenMcdf.CFItem item, System.Text.Encoding mbcsEncoding) : this((OpenMcdf.CFStream)item, mbcsEncoding) { }
        public DesignerInformation(OpenMcdf.CFStream cFStream, System.Text.Encoding mbcsEncoding) : this(cFStream.GetData(), mbcsEncoding) { }
        public DesignerInformation(byte[] bytes, System.Text.Encoding mbcsEncoding) : this(new System.IO.MemoryStream(bytes), mbcsEncoding) { }
        public DesignerInformation(System.IO.Stream stream, System.Text.Encoding mbcsEncoding) : this(new Edmosoft.IO.StreamReader(stream), mbcsEncoding) { }
        public DesignerInformation(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            streamReader.encoding = mbcsEncoding;
            if (streamReader.ReadLine() != "VERSION 5.00") throw new FormatException("Unexpected Designer Properties VBFrame");
            string[] DesignerHeaders = streamReader.ReadLine().Split(new char[] { '\x09', '\x20' }, StringSplitOptions.RemoveEmptyEntries);
            if (DesignerHeaders[0] != "Begin") throw new FormatException("Unexpected Designer Properties");
            CLSID = DesignerHeaders[1];
            DesignerName = DesignerHeaders[2];
            properties = new List<ProjectProperty>();

            string propertyLine = streamReader.ReadLine();
            do
            {
                string[] propertyData = propertyLine.Split(new char[] { '\x09', '\x20', '\'' }, 4, StringSplitOptions.RemoveEmptyEntries);
                if (propertyData.Length == 4)
                {
                    properties.Add(new ProjectProperty(propertyData[0], propertyData[2], propertyData[3]));
                }
                else
                {
                    properties.Add(new ProjectProperty(propertyData[0], propertyData[2], null));
                }
                propertyLine = streamReader.ReadLine();
            } while (propertyLine.Trim() != "End");

            /*
            VBFrameText = "VERSION 5.00" NWLN
                          "Begin" 1*WSP DesignerCLSID 1*WSP DesignerName *WSP NWLN
                          DesignerProperties "End" NWLN
            DesignerCLSID = GUID
            DesignerName = ModuleIdentifier
            */
            /*
            DesignerProperties = [ *WSP DesignerCaption *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerHeight *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerLeft *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerTop *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerWidth *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerEnabled *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerHelpContextId *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerRTL *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerShowModal *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerStartupPosition *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerTag *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerTypeInfoVer *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerVisible *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerWhatsThisButton *WSP [ Comment ] NWLN ]
                                 [ *WSP DesignerWhatsThisHelp *WSP [ Comment ] NWLN ]
            Comment = "'" *ANYCHAR
            */
        }
    }
}
