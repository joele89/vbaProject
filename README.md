Library for reading Macros from Macro-Enabled Office Documents

Releases can be found on NUGET

## Usage
```

System.IO.StreamReader docReader = new System.IO.StreamReader("/path/to/Office.docm");
System.IO.Compression.ZipArchive docCompression = new System.IO.Compression.ZipArchive(docReader.BaseStream);
System.IO.Compression.ZipArchiveEntry vbaProjectArchive = docCompression.GetEntry("word/vbaProject.bin");
Edmosoft.Office.vbaProject.Project vbaProject = new Edmosoft.Office.vbaProject.Project(vbaProjectArchive.Open());

vbaProject.Modules.Where(x => x.Name == "Module1").First().Body;
```
