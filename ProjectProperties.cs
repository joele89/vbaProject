using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Office.vbaProject
{

    internal class ProjectProperties
    {
        internal ProjectProperty projectID { get; set; }
        internal List<ProjectProperty> projectItems { get; set; }
        internal ProjectProperty helpFile { get; set; }
        internal ProjectProperty exeName32 { get; set; }
        internal ProjectProperty name { get; set; }
        internal ProjectProperty helpID { get; set; }
        internal ProjectProperty description { get; set; }
        internal ProjectProperty cMG { get; set; }
        internal ProjectProperty password { get; set; }
        internal ProjectProperty projectVisibilityState { get; set; }

        internal List<ProjectProperty> hostExtenderRefs { get; set; }

        internal List<ProjectWindow> ProjectWindows { get; set; }

        public ProjectProperties(OpenMcdf.CFItem item, System.Text.Encoding mbcsEncoding) : this((OpenMcdf.CFStream)item, mbcsEncoding) { }
        public ProjectProperties(OpenMcdf.CFStream cFStream, System.Text.Encoding mbcsEncoding) : this(cFStream.GetData(), mbcsEncoding) { }
        public ProjectProperties(byte[] bytes, System.Text.Encoding mbcsEncoding) : this(new System.IO.MemoryStream(bytes), mbcsEncoding) { }
        public ProjectProperties(System.IO.Stream stream, System.Text.Encoding mbcsEncoding) : this(new Edmosoft.IO.StreamReader(stream), mbcsEncoding) { }
        public ProjectProperties(Edmosoft.IO.StreamReader streamReader, System.Text.Encoding mbcsEncoding)
        {
            streamReader.encoding = mbcsEncoding;

            /*
             VBAPROJECTText = ProjectProperties NWLN
                              HostExtenders
                              [NWLN ProjectWorkspace]
            */
            decodeProjectProperties(streamReader);
            if (streamReader.ReadLine() != "") throw new FormatException("Expected Empty Line");
            decodeHostExtenders(streamReader);
            try
            {
                if (streamReader.Peek() < 0) return;
            }
            catch { return; }
            decodeProjectWorkspace(streamReader);
        }

        void decodeProjectProperties(Edmosoft.IO.StreamReader streamReader)
        {
            /*
             ProjectProperties = ProjectId
                                 *ProjectItem
                                 [ProjectHelpFile]
                                 [ProjectExeName32]
                                 ProjectName
                                 ProjectHelpId
                                 [ProjectDescription]
                                 [ProjectVersionCompat32]
                                 ProjectProtectionState
                                 ProjectPassword
                                 ProjectVisibilityState
            */
            projectID = new ProjectProperty(streamReader.ReadLine(), "ID"); //"ID=" DQUOTE ProjectCLSID DQUOTE NWLN
                                                                            //ProjectCLSID = GUID
            decodeProjectItems(streamReader); // ProjectItem = ( ProjectModule / ProjectPackage ) NWLN
            ProjectProperty abnf = new ProjectProperty(streamReader.ReadLine());
            if (abnf.Name == "HelpFile")
            {
                helpFile = abnf; // ProjectHelpFile = "HelpFile=" PATH NWLN
                abnf = new ProjectProperty(streamReader.ReadLine());
            }
            if (abnf.Name == "ExeName32")
            {
                exeName32 = abnf; // ProjectExeName32 = "ExeName32=" PATH NWLN
                abnf = new ProjectProperty(streamReader.ReadLine());
            }
            if (abnf.Name == "Name")    // ProjectName = "Name=" DQUOTE ProjectIdentifier DQUOTE NWLN
                name = abnf;            // ProjectIdentifier = 1*128QUOTEDCHAR
            else
                throw new FormatException(abnf.Name + " was not expected at this time.");

            helpID = new ProjectProperty(streamReader.ReadLine(), "HelpContextID"); // ProjectHelpId = "HelpContextID=" DQUOTE TopicId DQUOTE NWLN
                                                                                    // TopicId = INT32
            abnf = new ProjectProperty(streamReader.ReadLine());
            if (abnf.Name == "Description")
            {
                description = abnf; // ProjectDescription = "Description=" DQUOTE DescriptionText DQUOTE NWLN
                                    // DescriptionText = *2000QUOTEDCHAR
                abnf = new ProjectProperty(streamReader.ReadLine());
            }
            if (abnf.Name == "VersionCompatible32")
            {
                // ProjectVersionCompat32 = "VersionCompatible32=" DQUOTE "393222000" DQUOTE NWLN
                abnf = new ProjectProperty(streamReader.ReadLine());
            }
            if (abnf.Name == "CMG")
                cMG = abnf;  // ProjectProtectionState = "CMG=" DQUOTE EncryptedState DQUOTE NWLN
            else             // EncryptedState = 22*28HEXDIG
                throw new FormatException(abnf.Name + " was not expected at this time.");
            password = new ProjectProperty(streamReader.ReadLine()); // ProjectPassword = "DPB=" DQUOTE EncryptedPassword DQUOTE NWLN
                                                                     // EncryptedPassword = 16*HEXDIG
            projectVisibilityState = new ProjectProperty(streamReader.ReadLine());  // ProjectVisibilityState = "GC=" DQUOTE EncryptedProjectVisibility DQUOTE NWLN
                                                                                    // EncryptedProjectVisibility = 16*22HEXDIG
        }

        void decodeProjectItems(Edmosoft.IO.StreamReader streamReader)
        {
            /*
               ProjectModule = (ProjectDocModule /
                                ProjectStdModule /
                                ProjectClassModule /
                                ProjectDesignerModule)

               ProjectPackage = "Package=" GUID
            */

            projectItems = new List<ProjectProperty>();

            do
            {
                long start = streamReader.BaseStream.Position;
                ProjectProperty aBNF = new ProjectProperty(streamReader.ReadLine());
                switch (aBNF.Name)
                {
                    case "Package":             // ProjectPackage = "Package=" GUID
                    case "Document":            // ProjectDocModule = "Document=" ModuleIdentifier %x2f DocTlibVer
                                                // DocTlibVer = HEXINT32
                    case "Module":              // ProjectStdModule = "Module=" ModuleIdentifier
                    case "Class":               // ProjectClassModule = "Class=" ModuleIdentifier
                    case "BaseClass":           // ProjectDesignerModule = "BaseClass=" ModuleIdentifier
                        projectItems.Add(aBNF); // ModuleIdentifier -- SHOULD be an identifier as specified by [MS-VBAL] section 3.3.5
                        break;
                    default:
                        streamReader.BaseStream.Position = start;
                        return;

                }
            } while (true);
        }
        void decodeHostExtenders(Edmosoft.IO.StreamReader streamReader)
        {
            /*
             HostExtenders = "[Host Extender Info]" NWLN
                               *HostExtenderRef

             HostExtenderRef = ExtenderIndex "=" ExtenderGuid ";" LibName ";" CreationFlags NWLN

             ExtenderIndex = HEXINT32

             ExtenderGuid = GUID

             LibName = "VBE" / *(%x21-3A / %x3C-FF)

             CreationFlags = HEXINT32
            */
            if (streamReader.ReadLine() != "[Host Extender Info]") throw new FormatException("Expected [Host Extender Info]");
            hostExtenderRefs = new List<ProjectProperty>();
            do
            {
                string line = streamReader.ReadLine();
                if (line == "") break;
                hostExtenderRefs.Add(new ProjectProperty(line));
            } while (streamReader.DataAvailable);
        }
        void decodeProjectWorkspace(Edmosoft.IO.StreamReader streamReader)
        {
            /*
            ProjectWorkspace = "[Workspace]" NWLN
                                *ProjectWindowRecord
            
            ProjectWindowRecord = ModuleIdentifier "=" ProjectWindowState NWLN

            ProjectWindowState = CodeWindow [ ", " DesignerWindow ]

            CodeWindow = ProjectWindow

            DesignerWindow = ProjectWindow
            
            ProjectWindow = WindowLeft ", " WindowTop ", " WindowRight ", " WindowBottom ", " WindowState
            WindowLeft = INT32
            WindowTop = INT32
            WindowRight = INT32
            WindowBottom = INT32
            WindowState = [("C" / "Z" / "I")]
            */
            if (streamReader.ReadLine() != "[Workspace]") throw new FormatException("Expected [Workspace]");
            ProjectWindows = new List<ProjectWindow>();
            do
            {
                string line = streamReader.ReadLine();
                if (line == "") break;
                ProjectWindows.Add(new ProjectWindow(new ProjectProperty(line)));
            } while (streamReader.DataAvailable);
        }
    }
    public class ProjectWindow
    {
        internal ProjectWindow(ProjectProperty projectProperty)
        {
            this.Name = projectProperty.Name;
            string[] boundaries = projectProperty.Value.Split(',');
            if (boundaries.Length >= 4)
            {
                CodeWindowLeft = int.Parse(boundaries[0].Trim());
                CodeWindowTop = int.Parse(boundaries[1].Trim());
                CodeWindowRight = int.Parse(boundaries[2].Trim());
                CodeWindowBottom = int.Parse(boundaries[3].Trim());
                CodeWindowState = WorkspaceWindowState.Normal;
            }
            if (boundaries.Length >= 5)
            {
                switch (boundaries[4].Trim())
                {
                    case "C":
                        CodeWindowState = WorkspaceWindowState.Closed;
                        break;
                    case "Z":
                        CodeWindowState = WorkspaceWindowState.Maximized;
                        break;
                    case "I":
                        CodeWindowState = WorkspaceWindowState.Minimized;
                        break;
                    default:
                        CodeWindowState = WorkspaceWindowState.Normal;
                        break;
                }
            }
            if (boundaries.Length >= 9)
            {
                DesignerWindowLeft = int.Parse(boundaries[5].Trim());
                DesignerWindowTop = int.Parse(boundaries[6].Trim());
                DesignerWindowRight = int.Parse(boundaries[7].Trim());
                DesignerWindowBottom = int.Parse(boundaries[8].Trim());
                DesignerWindowState = WorkspaceWindowState.Normal;
            }
            if (boundaries.Length >= 10)
            {
                switch (boundaries[9].Trim())
                {
                    case "C":
                        DesignerWindowState = WorkspaceWindowState.Closed;
                        break;
                    case "Z":
                        DesignerWindowState = WorkspaceWindowState.Maximized;
                        break;
                    case "I":
                        DesignerWindowState = WorkspaceWindowState.Minimized;
                        break;
                    default:
                        DesignerWindowState = WorkspaceWindowState.Normal;
                        break;
                }
            }
        }

        public string Name { get; set; }

        public Int32 CodeWindowLeft { get; set; }
        public Int32 CodeWindowTop { get; set; }
        public Int32 CodeWindowRight { get; set; }
        public Int32 CodeWindowBottom { get; set; }

        public WorkspaceWindowState CodeWindowState { get; set; }

        public Int32? DesignerWindowLeft { get; set; }
        public Int32? DesignerWindowTop { get; set; }
        public Int32? DesignerWindowRight { get; set; }
        public Int32? DesignerWindowBottom { get; set; }

        public WorkspaceWindowState? DesignerWindowState { get; set; }

        public enum WorkspaceWindowState
        {
            Closed,
            Maximized,
            Minimized,
            Normal
        }
    }

    internal class ProjectProperty
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string[] Comments { get; set; }

        public ProjectProperty(string line, string expectName) : this(line)
        {
            if (Name != expectName) throw new FormatException(Name + " was not expected at this time.");
        }
        public ProjectProperty(string line)
        {
            string[] spline = line.Split(new char[] { '=' }, 2);
            Name = spline[0].Trim();
            string[] spval = spline[1].Trim().Split(new char[] { ';' });
            Value = spval[0].Trim();
            if (spval.Length > 1)
            {
                Comments = new string[spval.Length - 1];
                Array.Copy(spval, 1, Comments, 0, (long)(spval.Length - 1));
            }
        }
        public ProjectProperty(string Name, string Value, string Comments)
        {
            this.Name = Name;
            this.Value = Value;
            this.Comments = new string[] { Comments };
        }
    }
}


/*
* Example
ID="{00000000-0000-0000-0000-000000000000}"
Document=ThisWorkbook/&H00000000
Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}
Document=Sheet2/&H00000000
Document=Sheet4/&H00000000
Document=Sheet5/&H00000000
Document=Sheet6/&H00000000
Document=Sheet8/&H00000000
Document=Sheet7/&H00000000
Document=Sheet10/&H00000000
Document=Sheet11/&H00000000
Document=Sheet12/&H00000000
Document=Sheet13/&H00000000
Document=Sheet14/&H00000000
Document=Sheet15/&H00000000
Document=TrackerID/&H00000000
Document=Sheet17/&H00000000
Document=Sheet18/&H00000000
Document=Sheet19/&H00000000
Document=Sheet20/&H00000000
Document=Sheet21/&H00000000
Document=Sheet22/&H00000000
Document=Sheet1/&H00000000
Document=Sheet23/&H00000000
Document=Sheet24/&H00000000
Document=Sheet3/&H00000000
Document=Sheet9/&H00000000
Module=aaa_Dev_Notes
Module=aImport_1_Main
Module=aImport_2_FileFormat
Module=aImport_3_Process_EXL_Invoice
Module=aImport_3_Process_EXL_TimeSht
Module=aImport_4_Process_PDF
Module=aImport_4_Process_PDF_ADC
Module=aImport_4_Process_PDF_CMG
Module=aImport_5_Check
Module=aImport_6_MasterFile
Module=aImport_8_ExceptionReport_out
Module=aImport_9_ExceptionCheck
Module=aImport_9_ExceptionReport_In
Module=bDatastore_1_Control
Module=bDataStore_2_Write
Module=bDataStore_3_Read
Module=bDatastore_4_DuplicateData
Module=bDatastore_4_TimesheetAppend
Module=bDatastore_4_TimesheetReplace
Module=bDatastore_5_IntegrityCheck
Module=bDatastore_6_Attachments
Module=bDatastore_7_Backup
Class=ClsEvent_AppController
Class=ClsEvent_UserFormController
Class=Cls_Attachment
Class=Cls_Backup
Class=Cls_Break
Class=Cls_Charged
Class=Cls_DuplicateResponse
Class=Cls_Email
Class=Cls_EmailInvoice
Class=Cls_ImportStat
Class=Cls_IndexedArray
Class=Cls_InvoiceEntry
Class=Cls_InvoiceSummary
Class=Cls_StoredFile
Class=Cls_TimesheetEntry
Class=Cls_TimesheetSummary
Class=Cls_UserFormButtonStyles
Module=cReports_1_Main
Module=cReports_2_Spawn
Module=cReports_3_Control
Module=cReports_4_Validation
Module=dEmail_1_Main
Module=dForms_1_Control
Module=dForms_2_Attachments
Module=fConfig_Settings
Module=fFormattingPass
BaseClass=Frm_About
BaseClass=Frm_Amend
BaseClass=Frm_ConfirmNo
BaseClass=Frm_DuplicateInvoice
BaseClass=Frm_Emails
BaseClass=Frm_Status
BaseClass=Frm_Validation
BaseClass=Frm_ValidationAlt
Module=iChecks
Module=iFunctions_Bespoke
Module=xSystem_ProgressBar
Module=yAddins
Module=yConstants
Module=yDev
Module=yPublic
Module=yRibbonControls
Module=zFunctions_Arrays
Module=zFunctions_Collection
Module=zFunctions_Dictionaries
Module=zFunctions_Excel
Module=zFunctions_FileSystem
Module=zFunctions_Hashing
Module=zFunctions_HTML
Module=zFunctions_PDF
Module=zFunctions_Standard
Module=zFunctions_Tables
Module=zFunctions_Userform2
Document=Sheet16/&H00000000
HelpFile="130684965"
Name="VBAProject"
HelpContextID="0"
VersionCompatible32="393222000"
CMG="5557F98E0B1F0F1F0F1B131B13"
DPB="E2E04E01C61EC61E39E2C71E62AFC2F814D81047C781302ADD6EBD48DCAD18687A844C4AE8"
GC="6F6DC394DDAC6BAD6BAD6B"

[Host Extender Info]
&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000

[Workspace]
ThisWorkbook=96, 96, 1227, 516, 
Sheet2=182, 182, 1272, 577, 
Sheet4=0, 0, 1152, 427, 
Sheet5=208, 208, 1407, 680, 
Sheet6=26, 26, 1178, 453, 
Sheet8=0, 0, 0, 0, C
Sheet7=104, 104, 1472, 587, 
Sheet10=0, 0, 0, 0, C
Sheet11=0, 0, 0, 0, C
Sheet12=0, 0, 1209, 557, 
Sheet13=156, 156, 1465, 531, 
Sheet14=78, 78, 1342, 541, C
Sheet15=0, 0, 0, 0, C
TrackerID=26, 26, 1235, 583, 
Sheet17=0, 0, 0, 0, C
Sheet18=0, 0, 0, 0, C
Sheet19=52, 52, 1361, 427, 
Sheet20=0, 0, 1309, 533, 
Sheet21=130, 130, 1491, 572, C
Sheet22=0, 0, 0, 0, C
Sheet1=0, 0, 0, 0, C
Sheet23=0, 0, 0, 0, C
Sheet24=26, 26, 925, 481, 
Sheet3=0, 0, 0, 0, C
Sheet9=0, 0, 0, 0, C
aaa_Dev_Notes=0, 0, 1268, 530, 
aImport_1_Main=156, 156, 1424, 686, 
aImport_2_FileFormat=156, 156, 1532, 623, 
aImport_3_Process_EXL_Invoice=104, 104, 1413, 549, 
aImport_3_Process_EXL_TimeSht=26, 26, 890, 489, 
aImport_4_Process_PDF=52, 52, 1346, 528, 
aImport_4_Process_PDF_ADC=52, 52, 1428, 519, 
aImport_4_Process_PDF_CMG=78, 78, 1454, 545, 
aImport_5_Check=104, 104, 1480, 571, 
aImport_6_MasterFile=26, 26, 1294, 556, 
aImport_8_ExceptionReport_out=52, 52, 1346, 528, 
aImport_9_ExceptionCheck=0, 0, 1431, 467, 
aImport_9_ExceptionReport_In=0, 0, 1294, 476, 
bDatastore_1_Control=52, 52, 1361, 497, 
bDataStore_2_Write=78, 78, 1457, 641, 
bDataStore_3_Read=78, 78, 1398, 497, 
bDatastore_4_DuplicateData=52, 52, 1428, 519, 
bDatastore_4_TimesheetAppend=52, 52, 1361, 497, 
bDatastore_4_TimesheetReplace=104, 104, 1535, 571, Z
bDatastore_5_IntegrityCheck=104, 104, 1478, 571, 
bDatastore_6_Attachments=0, 0, 1268, 530, 
bDatastore_7_Backup=130, 130, 1439, 575, 
ClsEvent_AppController=208, 208, 1639, 675, 
ClsEvent_UserFormController=234, 234, 1502, 764, 
Cls_Attachment=0, 0, 0, 0, C
Cls_Backup=0, 0, 0, 0, C
Cls_Break=156, 156, 1530, 623, 
Cls_Charged=182, 182, 1556, 649, 
Cls_DuplicateResponse=0, 0, 0, 0, C
Cls_Email=156, 156, 1424, 686, 
Cls_EmailInvoice=0, 0, 0, 0, C
Cls_ImportStat=0, 0, 0, 0, C
Cls_IndexedArray=130, 130, 1505, 597, 
Cls_InvoiceEntry=0, 0, 0, 0, C
Cls_InvoiceSummary=130, 130, 1504, 597, 
Cls_StoredFile=0, 0, 0, 0, C
Cls_TimesheetEntry=208, 208, 1582, 675, 
Cls_TimesheetSummary=26, 26, 890, 489, 
Cls_UserFormButtonStyles=0, 0, 864, 463, 
cReports_1_Main=52, 52, 1320, 582, 
cReports_2_Spawn=26, 26, 1535, 348, 
cReports_3_Control=52, 52, 1561, 374, 
cReports_4_Validation=104, 104, 1478, 571, Z
dEmail_1_Main=130, 130, 1398, 660, 
dForms_1_Control=208, 208, 1476, 738, 
dForms_2_Attachments=260, 260, 1528, 790, 
fConfig_Settings=130, 130, 1398, 660, 
fFormattingPass=26, 26, 1400, 493, 
Frm_About=208, 208, 1583, 675, , 78, 78, 547, 536, 
Frm_Amend=104, 104, 1372, 634, , 104, 104, 573, 562, C
Frm_ConfirmNo=0, 0, 0, 0, C, 130, 130, 599, 588, C
Frm_DuplicateInvoice=0, 0, 0, 0, C, 156, 156, 625, 614, C
Frm_Emails=104, 104, 1372, 634, , 182, 182, 651, 640, C
Frm_Status=0, 0, 1379, 563, , 208, 208, 677, 666, C
Frm_Validation=104, 104, 1535, 571, , 0, 0, 469, 458, C
Frm_ValidationAlt=78, 78, 1346, 608, , 26, 26, 495, 484, C
iChecks=0, 0, 0, 0, C
iFunctions_Bespoke=0, 0, 1294, 476, 
xSystem_ProgressBar=26, 26, 1405, 589, 
yAddins=52, 52, 521, 510, Z
yConstants=182, 182, 1613, 649, 
yDev=208, 208, 1582, 675, 
yPublic=0, 0, 0, 0, C
yRibbonControls=182, 182, 1450, 712, 
zFunctions_Arrays=26, 26, 1400, 493, 
zFunctions_Collection=78, 78, 1454, 545, 
zFunctions_Dictionaries=182, 182, 1046, 645, 
zFunctions_Excel=156, 156, 1020, 619, 
zFunctions_FileSystem=130, 130, 1398, 660, 
zFunctions_Hashing=0, 0, 0, 0, C
zFunctions_HTML=208, 208, 1072, 671, 
zFunctions_PDF=0, 0, 864, 463, Z
zFunctions_Standard=156, 156, 1424, 686, 
zFunctions_Tables=156, 156, 1531, 623, 
zFunctions_Userform2=182, 182, 1450, 712, 
Sheet16=0, 0, 0, 0, C
*/