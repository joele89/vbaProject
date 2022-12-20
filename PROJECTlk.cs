using System;
using System.Collections.Generic;
using System.Linq;

namespace Edmosoft.Office.vbaProject
{
    internal class PROJECTlk
    {
        UInt16 Version { get; }
        internal List<LicenseInfoRecord> LicenseInfoRecords { get; }
        public PROJECTlk(OpenMcdf.CFItem item) : this((OpenMcdf.CFStream)item) { }
        public PROJECTlk(OpenMcdf.CFStream cFStream) : this(cFStream.GetData()) { }
        public PROJECTlk(byte[] bytes) : this(new System.IO.MemoryStream(bytes)) { }
        public PROJECTlk(System.IO.Stream stream) : this(new Edmosoft.IO.StreamReader(stream)) { }
        public PROJECTlk(Edmosoft.IO.StreamReader streamReader)
        {
            Version = streamReader.ReadUInt16();
            UInt32 Count = streamReader.ReadUInt32();
            LicenseInfoRecords = new List<LicenseInfoRecord>();
            for (UInt32 i = 0; i <= Count; i++)
            {
                LicenseInfoRecords.Add(new LicenseInfoRecord(streamReader));
            }
        }
    }
    public class LicenseInfoRecord
    {
        public Guid ClassID { get; set; }
        public UInt32 SizeOfLicenseKey { get; set; }
        public byte[] LicenseKey { get; set; }
        public UInt32 LicenseRequired { get; set; }

        public LicenseInfoRecord(Edmosoft.IO.StreamReader streamReader)
        {
            ClassID = new Guid(streamReader.ReadBlock(16));
            SizeOfLicenseKey = streamReader.ReadUInt32();
            LicenseKey = streamReader.ReadBlock((int)SizeOfLicenseKey);
            LicenseRequired = streamReader.ReadUInt32();
        }
    }
}
