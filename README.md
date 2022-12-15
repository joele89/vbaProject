Library for reading Macros from Macro-Enabled Office Documents


## Usage
```

System.IO.StreamReader docReader = new System.IO.StreamReader("/path/to/Office.docm");
System.IO.Compression.ZipArchive docCompression = new System.IO.Compression.ZipArchive(docReader.BaseStream);
System.IO.Compression.ZipArchiveEntry vbaProjectArchive = docCompression.GetEntry("word/vbaProject.bin");
Microsoft.Office.vbaProject.Project vbaProject = new Microsoft.Office.vbaProject.Project(vbaProjectArchive.Open());

vbaProject.Modules.Where(x => x.Name == "Module1").First().Body;
```
