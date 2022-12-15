using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Office.vbaProject
{
    internal class _VBA_PROJECT
    {
        UInt16 Version { get; set; }
        public _VBA_PROJECT(OpenMcdf.CFItem item) : this((OpenMcdf.CFStream)item) { }
        public _VBA_PROJECT(OpenMcdf.CFStream cFStream) : this(cFStream.GetData()) { }
        public _VBA_PROJECT(byte[] bytes) : this(new System.IO.MemoryStream(bytes)) { }
        public _VBA_PROJECT(System.IO.Stream stream) : this(new Edmosoft.IO.StreamReader(stream)) { }
        public _VBA_PROJECT(Edmosoft.IO.StreamReader streamReader)
        {
            byte[] Reserved1 = streamReader.ReadBlock(2);
            Version = streamReader.ReadUInt16();
            byte Reserved2 = streamReader.ReadByte();
            byte[] Reserved3 = streamReader.ReadBlock(2);
            string PerformanceCache = streamReader.ReadToEnd();
        }
    }
}
