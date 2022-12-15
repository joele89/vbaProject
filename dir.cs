using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Microsoft.Office.vbaProject.ProjectReferences;

namespace Microsoft.Office.vbaProject
{
    internal class dir
    {
        public ProjectInformation InformationRecord { get; set; }
        public ProjectReferences ReferencesRecord { get; set; }
        public ProjectModules ModulesRecord { get; set; }

        public dir(OpenMcdf.CFItem item) : this((OpenMcdf.CFStream)item) { }
        public dir(OpenMcdf.CFStream cFStream) : this(cFStream.GetData()) { }
        public dir(byte[] bytes) : this(new System.IO.MemoryStream(bytes)) { }
        public dir(System.IO.Stream stream) : this(new Edmosoft.IO.StreamReader(stream)) { }
        public dir(Edmosoft.IO.StreamReader streamReader)
        {
            System.IO.Stream stream = Edmosoft.IO.vbaCompression.vbaStreamReader.Decode(streamReader);
            streamReader = new Edmosoft.IO.StreamReader(stream);
            InformationRecord = new ProjectInformation(streamReader); //2.3.4.2.1
            System.Text.Encoding mbcsEncoding = System.Text.Encoding.GetEncoding(InformationRecord.CodePage);
            ReferencesRecord = new ProjectReferences(streamReader, mbcsEncoding); //2.3.4.2.2
            ModulesRecord = new ProjectModules(streamReader, mbcsEncoding); //2.3.4.2.3
            UInt16 Terminator = streamReader.ReadUInt16();
            byte[] Reserved = streamReader.ReadBlock(4);
        }
    }
    public class ProjectInformation
    {
        ProjectInformationRecord<SysKind> SysKindRecord;
        ProjectInformationRecord<UInt32> CompactVersionRecord;
        private ProjectInformationRecord<UInt16> CodePageRecord;
        ProjectInformationRecord<UInt32> LcidRecord;
        ProjectInformationRecord<UInt32> LcidInvokeRecord;
        ProjectInformationRecord<string> NameRecord;
        ProjectInformationRecord<string> DocStringRecord;
        ProjectInformationRecord<string> DocStringUnicodeRecord;
        ProjectInformationRecord<string> HelpFile1PathRecord;
        ProjectInformationRecord<string> HelpFile2PathRecord;
        ProjectInformationRecord<UInt32> HelpContextRecord;
        ProjectInformationRecord<UInt32> LibFlagsRecord;
        PROJECTVERSION VersionRecord;
        PROJECTCONSTANTS ConstantsRecord;
        public UInt16 CodePage { get { return CodePageRecord.Value; } }
        private System.Text.Encoding mbcsEncoding;
        public ProjectInformation(Edmosoft.IO.StreamReader streamReader)
        {
            SysKindRecord = new ProjectInformationRecord<SysKind>(streamReader, 0x01); //2.3.4.2.1.1
            try
            {
                CompactVersionRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x4A); //2.3.4.2.1.2
            }
            catch { }
            LcidRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x02); //2.3.4.2.1.3
            LcidInvokeRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x14); //2.3.4.2.1.4
            CodePageRecord = new ProjectInformationRecord<UInt16>(streamReader, 0x03); //2.3.4.2.1.5
            mbcsEncoding = System.Text.Encoding.GetEncoding(CodePageRecord.Value);
            NameRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x04); //2.3.4.2.1.6
            DocStringRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x05); //2.3.4.2.1.7
            DocStringUnicodeRecord = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x40); //2.3.4.2.1.7
            HelpFile1PathRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x06); //2.3.4.2.1.8
            HelpFile2PathRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x3D); //2.3.4.2.1.8
            HelpContextRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x07); //2.3.4.2.1.9
            LibFlagsRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x08); //2.3.4.2.1.10
            VersionRecord = new PROJECTVERSION(streamReader); //2.3.4.2.1.11
            ConstantsRecord = new PROJECTCONSTANTS(streamReader, mbcsEncoding); //2.3.4.2.1.12
        }

        public enum SysKind : UInt32
        {
            x16Bit,
            x32Bit,
            Macos,
            x64Bit
        }

        public class PROJECTVERSION
        {
            public UInt16 Id;
            public UInt32 VersionMajor;
            public UInt16 VersionMinor;
            public PROJECTVERSION(Edmosoft.IO.StreamReader streamReader)
            {
                long start = streamReader.BaseStream.Position;
                Id = streamReader.ReadUInt16();
                if (Id != 0x09)
                {
                    streamReader.BaseStream.Position = start;
                    throw new FormatException(string.Format("Requested ID does not match stream. Expected {0} Got {1}", 0x09, Id));
                }
                byte[] Reserved = streamReader.ReadBlock(4);
                VersionMajor = streamReader.ReadUInt32();
                VersionMinor = streamReader.ReadUInt16();
            }
        }

        public class PROJECTCONSTANTS
        {
            ProjectInformationRecord<string> ConstantsMBCS;
            ProjectInformationRecord<string> ConstantsUnicode;
            public PROJECTCONSTANTS(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
            {
                ConstantsMBCS = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x0C);
                ConstantsUnicode = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x3C);
            }
        }
    }

    internal class ProjectInformationRecord<T>
    {
        public UInt16 Id;
        public T Value;
        public ProjectInformationRecord(Edmosoft.IO.StreamReader streamReader)
        {
            Id = streamReader.ReadUInt16();
            UInt32 Size = streamReader.ReadUInt32();
            if (typeof(T) == typeof(UInt32))
            {
                if (Size != 4) throw new FormatException("Encoded size of UInt32 must be 0x00000004");
                Value = (T)Convert.ChangeType(streamReader.ReadUInt32(), typeof(T));
            }
            else if (typeof(T) == typeof(UInt16))
            {
                if (Size != 2) throw new FormatException("Encoded size of UInt16 must be 0x00000002");
                Value = (T)Convert.ChangeType(streamReader.ReadUInt16(), typeof(T));
            }
            else
                throw new NotImplementedException();
        }
        public ProjectInformationRecord(Edmosoft.IO.StreamReader streamReader, UInt16 ExpectID)
        {
            long start = streamReader.BaseStream.Position;
            Id = streamReader.ReadUInt16();
            if (Id != ExpectID)
            {
                streamReader.BaseStream.Position = start;
                throw new FormatException(string.Format("Requested ID does not match stream. Expected {0} Got {1}", ExpectID, Id));
            }
            UInt32 Size = streamReader.ReadUInt32();
            if (typeof(T) == typeof(UInt32) || typeof(T).BaseType == typeof(UInt32)) 
            {
                if (Size != 4) throw new FormatException("Encoded size of UInt32 must be 0x00000004");
                Value = (T)Convert.ChangeType(streamReader.ReadUInt32(), typeof(T));
            }
            else if (typeof(T) == typeof(UInt16))
            {
                if (Size != 2) throw new FormatException("Encoded size of UInt16 must be 0x00000002");
                Value = (T)Convert.ChangeType(streamReader.ReadUInt16(), typeof(T));
            }
            else
                throw new NotImplementedException();
        }
        public ProjectInformationRecord(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding encoding)
        {
            Id = streamReader.ReadUInt16();
            UInt32 Size = streamReader.ReadUInt32();
            if (typeof(T) == typeof(string))
                Value = (T)Convert.ChangeType(encoding.GetString(streamReader.ReadBlock((int)Size)), typeof(T));
            else
                throw new NotImplementedException();
        }
        public ProjectInformationRecord(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding encoding, UInt16 ExpectID)
        {
            long start = streamReader.BaseStream.Position;
            Id = streamReader.ReadUInt16();
            if (Id != ExpectID)
            {
                streamReader.BaseStream.Position = start;
                throw new FormatException(string.Format("Requested ID does not match stream. Expected {0} Got {1}", ExpectID, Id));
            }
            UInt32 Size = streamReader.ReadUInt32();
            if (typeof(T) == typeof(string))
                Value = (T)Convert.ChangeType(encoding.GetString(streamReader.ReadBlock((int)Size)), typeof(T));
            else
                throw new NotImplementedException();
        }
    }

    public class ProjectReferences : List<ProjectReference>
    {
        public ProjectReferences(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            UInt16 peek = 0;
            do
            {
                this.Add(new ProjectReference(streamReader, mbcsEncoding));
                peek = streamReader.ReadUInt16();
                streamReader.BaseStream.Seek(-2, System.IO.SeekOrigin.Current);
            } while (peek != 0x0F);
        }
    }

    public class ProjectReference
    {
        ProjectInformationRecord<string> NameRecord;
        ProjectInformationRecord<string> NameRecordUnicode;
        public UInt16 ReferenceRecordType;
        public object ReferenceRecord;
        public string Name { get { return NameRecordUnicode.Value; } }
        public ProjectReference(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            NameRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x16);
            NameRecordUnicode = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x3E);
            ReferenceRecordType = streamReader.ReadUInt16();
            switch (ReferenceRecordType)
            {
                case 0x2F:
                    ReferenceRecord = new ReferenceControl(streamReader, mbcsEncoding);
                    break;
                case 0x33:
                    ReferenceOrigional origional = new ReferenceOrigional(streamReader, mbcsEncoding);
                    UInt16 Peek = streamReader.ReadUInt16();
                    if (Peek == 0x2F)
                    {
                        ReferenceRecord = new ReferenceControl(streamReader, origional);
                    }
                    else
                    {
                        streamReader.BaseStream.Seek(-2, System.IO.SeekOrigin.Current);
                        ReferenceRecord = origional;
                    }
                    break;
                case 0x0D:
                    ReferenceRecord = new ReferenceRegistered(streamReader, mbcsEncoding);
                    break;
                case 0x0E:
                    ReferenceRecord = new ReferenceProject(streamReader, mbcsEncoding);
                    break;
                case 0x0F:
                    streamReader.BaseStream.Seek(-2, System.IO.SeekOrigin.Current);
                    return;
                default:
                    throw new FormatException("Unexpected byte in stream");
            }
        }

        public class ReferenceControl
        {
            public string LibidTwiddled;
            ProjectInformationRecord<string> NameRecordExtended;
            ProjectInformationRecord<string> NameRecordExtendedUnicode;
            public string LibidExtended;
            public Guid OrigionalTypeLib;
            public UInt32 Cookie;
            public ReferenceControl(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
            {
                ReferenceOrigional OrigionalRecord = new ReferenceOrigional(streamReader, mbcsEncoding);
                UInt16 Id = streamReader.ReadUInt16();
                UInt32 SizeTwiddled = streamReader.ReadUInt32();
                UInt32 SizeOfLibidTwiddled = streamReader.ReadUInt32();
                LibidTwiddled = mbcsEncoding.GetString(streamReader.ReadBlock((int)SizeOfLibidTwiddled));
                byte[] Reserved1 = streamReader.ReadBlock(4);
                byte[] Reserved2 = streamReader.ReadBlock(2);
                long start = streamReader.BaseStream.Position;
                try
                {
                    NameRecordExtended = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.ASCII, 0x16);  //Optional
                    NameRecordExtendedUnicode = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x3E);  //Optional
                }
                catch { streamReader.BaseStream.Position = start; }
                Byte[] Reserved3 = streamReader.ReadBlock(2);
                UInt32 SizeExtended = streamReader.ReadUInt32();
                UInt32 SizeOfLibidExtended = streamReader.ReadUInt32();
                LibidExtended = mbcsEncoding.GetString(streamReader.ReadBlock((int)SizeOfLibidExtended));
                byte[] Reserved4 = streamReader.ReadBlock(4);
                byte[] Reserved5 = streamReader.ReadBlock(2);
                OrigionalTypeLib = new Guid(streamReader.ReadBlock(16));
                Cookie = streamReader.ReadUInt32();
            }
            public ReferenceControl(Edmosoft.IO.StreamReader streamReader, ReferenceOrigional OrigionalRecord)
            {
                //OrigionalRecord Provided
                //UInt16 Id = streamReader.ReadUInt16();
                UInt32 SizeTwiddled = streamReader.ReadUInt32();
                UInt32 SizeOfLibidTwiddled = streamReader.ReadUInt32();
                LibidTwiddled = System.Text.Encoding.ASCII.GetString(streamReader.ReadBlock((int)SizeOfLibidTwiddled));
                byte[] Reserved1 = streamReader.ReadBlock(4);
                byte[] Reserved2 = streamReader.ReadBlock(2);
                long start = streamReader.BaseStream.Position;
                try
                {
                    NameRecordExtended = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.ASCII, 0x16);  //Optional
                    NameRecordExtendedUnicode = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x3E);  //Optional
                }
                catch
                {
                    streamReader.BaseStream.Position = start;
                }
                Byte[] Reserved3 = streamReader.ReadBlock(2);
                UInt32 SizeExtended = streamReader.ReadUInt32();
                UInt32 SizeOfLibidExtended = streamReader.ReadUInt32();
                LibidExtended = System.Text.Encoding.ASCII.GetString(streamReader.ReadBlock((int)SizeOfLibidExtended));
                byte[] Reserved4 = streamReader.ReadBlock(4);
                byte[] Reserved5 = streamReader.ReadBlock(2);
                OrigionalTypeLib = new Guid(streamReader.ReadBlock(16));
                Cookie = streamReader.ReadUInt32();
            }
        }
        public class ReferenceOrigional
        {
            private string LibidOrigional;
            public ReferenceOrigional(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
            {
                UInt32 SizeOfLibidOrigional = streamReader.ReadUInt32();
                LibidOrigional = mbcsEncoding.GetString(streamReader.ReadBlock((int)SizeOfLibidOrigional));
            }
        }
        public class ReferenceRegistered
        {
            public string Libid;
            public ReferenceRegistered(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
            {
                UInt32 Size = streamReader.ReadUInt32();
                UInt32 SizeOfLibid = streamReader.ReadUInt32();
                Libid = mbcsEncoding.GetString(streamReader.ReadBlock((int)SizeOfLibid));
                byte[] Reserved1 = streamReader.ReadBlock(4);
                byte[] Reserved2 = streamReader.ReadBlock(2);
            }
        }
        public class ReferenceProject
        {
            public string LibidAbsolute;
            public string LibidRelative;
            public UInt32 MajorVersion;
            public UInt16 MinorVersion;
            public ReferenceProject(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
            {
                UInt32 Size = streamReader.ReadUInt32();
                UInt32 SizeOfLibidAbsolute = streamReader.ReadUInt32();
                LibidAbsolute = mbcsEncoding.GetString(streamReader.ReadBlock((int)SizeOfLibidAbsolute));
                UInt32 SizeOfLibidRelative = streamReader.ReadUInt32();
                LibidRelative = mbcsEncoding.GetString(streamReader.ReadBlock((int)SizeOfLibidRelative));
                MajorVersion = streamReader.ReadUInt32();
                MinorVersion = streamReader.ReadUInt16();
            }
        }
    }

    internal class ProjectModule
    {
        public enum ModuleType : UInt16
        {
            procedural = 0x21,
            @class = 0x22
        }
        private ProjectInformationRecord<string> NameRecord;
        private ProjectInformationRecord<string> NameUnicodeRecord;
        private ProjectInformationRecord<string> StreamNameRecord;
        private ProjectInformationRecord<string> StreamNameUnicodeRecord;
        private ProjectInformationRecord<string> DocStringRecord;
        private ProjectInformationRecord<string> DocStringUnicodeRecord;
        private ProjectInformationRecord<UInt32> OffsetRecord;
        private ProjectInformationRecord<UInt32> HelpContextRecord;
        private ProjectInformationRecord<UInt16> CookieRecord;
        private ProjectInformationRecord<string> TypeRecord;
        internal UInt32 Offset { get { return OffsetRecord.Value; } }
        public string Name { get { return NameUnicodeRecord.Value; } }
        public ModuleType Type { get { return (ModuleType)TypeRecord.Id; } }
        public ProjectModule(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            NameRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x19);
            NameUnicodeRecord = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x47);
            StreamNameRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x1A);
            StreamNameUnicodeRecord = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x32);
            DocStringRecord = new ProjectInformationRecord<string>(streamReader, mbcsEncoding, 0x1C);
            DocStringUnicodeRecord = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.Unicode, 0x48);
            OffsetRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x31);
            HelpContextRecord = new ProjectInformationRecord<UInt32>(streamReader, 0x1E);
            CookieRecord = new ProjectInformationRecord<UInt16>(streamReader, 0x2C);
            TypeRecord = new ProjectInformationRecord<string>(streamReader, System.Text.Encoding.ASCII); //0x22 or 0x21
            UInt16 NextBlock = 0x2B;
            do
            {
                NextBlock = streamReader.ReadUInt16();
                switch (NextBlock)
                {
                    case 0x25:
                        byte[] ReadOnlyRecordReserved = streamReader.ReadBlock(4); //Optional
                        break;
                    case 0x28:
                        byte[] PrivateRecordReserved = streamReader.ReadBlock(4); //Optional
                        break;
                    case 0x2B:
                        break;
                    default:
                        throw new FormatException("Unexpected byte in stream");
                }
            } while (NextBlock != 0x2B);

            byte[] Reserved = streamReader.ReadBlock(4);
        }
    }

    internal class ProjectModules : List<ProjectModule>
    {
        ProjectInformationRecord<UInt16> ProjectCookieRecord;
        public ProjectModules(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            UInt16 Id = streamReader.ReadUInt16();
            UInt32 Size = streamReader.ReadUInt32();
            UInt16 Count = streamReader.ReadUInt16();
            ProjectCookieRecord = new ProjectInformationRecord<UInt16>(streamReader);
            for (int i = 0; i < Count; i++)
            {
                this.Add(new ProjectModule(streamReader, mbcsEncoding));
            }
        }

        
    }
}

