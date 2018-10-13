/*
* simpleExcelWriter
* https://github.com/ewwink/simpleExcelWriter
* C# Class to create simple Excel .XLSX file 
*/

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace myAppNamespace
{
    public static class simpleExcelWriter
    {
        public static void create(string fileName, string rowContent)
        {
            Dictionary<string, string> xlsx = new Dictionary<string, string>
            {
               {"[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/worksheet.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" /></Types>"},
               {"styles.xml", "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font /></fonts><fills count=\"1\"><fill /></fills><borders count=\"1\"><border /></borders><cellStyleXfs count=\"1\"><xf /></cellStyleXfs><cellXfs count=\"2\"><xf /><xf applyAlignment=\"1\"><alignment wrapText=\"1\"/></xf></cellXfs></styleSheet>"},
               {"workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Instagram Profiles\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"},
               {"worksheet.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData>{0}</sheetData></worksheet>"},
               {"_rels\\.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"workbook.xml\"/></Relationships>"},
               {"_rels\\workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheet.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\" /></Relationships>"}
            };

            using (var memoryStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    foreach (var item in xlsx)
                    {
                        var demoFile = archive.CreateEntry(item.Key);

                        using (var entryStream = demoFile.Open())
                        using (var streamWriter = new StreamWriter(entryStream))
                        {
                            string value = item.Value;
                            if (item.Key == "worksheet.xml")
                                value = string.Format(value, rowContent);
                            streamWriter.Write(value);
                        }
                    }
                }

                using (var fileStream = new FileStream(fileName, FileMode.Create))
                {
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    memoryStream.CopyTo(fileStream);
                }
            }
        }
    }
}
