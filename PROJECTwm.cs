using System;
using System.Collections.Generic;
using System.Linq;

namespace Edmosoft.Office.vbaProject
{
    internal class PROJECTwm
    {
        public List<NameMapRecord> nameMapRecords;

        public PROJECTwm(OpenMcdf.CFItem item, System.Text.Encoding mbcsEncoding) : this((OpenMcdf.CFStream)item, mbcsEncoding) { }
        public PROJECTwm(OpenMcdf.CFStream cFStream, System.Text.Encoding mbcsEncoding) : this(cFStream.GetData(), mbcsEncoding) { }
        public PROJECTwm(byte[] bytes, System.Text.Encoding mbcsEncoding) : this(new System.IO.MemoryStream(bytes), mbcsEncoding) { }
        public PROJECTwm(System.IO.Stream stream, System.Text.Encoding mbcsEncoding) : this(new Edmosoft.IO.StreamReader(stream), mbcsEncoding) { }
        public PROJECTwm(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            nameMapRecords = new List<NameMapRecord>();
            do
            {
                nameMapRecords.Add(new NameMapRecord(streamReader, mbcsEncoding));
            } while (streamReader.DataAvailable && (streamReader.PeekChar() != "\x0000"));
        }
    }
    internal class NameMapRecord
    {
        public string ModuleName { get; set; }
        public string ModuleNameUnicode { get; set; }
        public NameMapRecord(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            streamReader.encoding = mbcsEncoding;
            ModuleName = streamReader.ReadLine(true);
            streamReader.encoding = System.Text.Encoding.Unicode;
            ModuleNameUnicode = streamReader.ReadLine(true);
        }
    }




}
