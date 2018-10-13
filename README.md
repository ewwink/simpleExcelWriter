# simpleExcelWriter
Very simple C# Class to create Excel `.xlsx` file

## Requirement
- Microsoft .NET Framework 4.5 or later (`ZipArchive`)

## Example
For creating 2 rows and column we need

**XML code**

    <row>
      <c t="inlineStr"><is><t>A1</t></is></c>
      <c t="inlineStr"><is><t>B1</t></is></c>
    </row>
    <row>
      <c t="inlineStr"><is><t>A2</t></is></c>
      <c t="inlineStr"><is><t>B2</t></is>
      </c>
    </row>

**C# code**

    string rowContent = "<row><c t=\"inlineStr\"><is><t>A1</t></is></c><c t=\"inlineStr\"><is><t>B1</t></is></c></row><row><c t=\"inlineStr\"><is><t>A2</t></is></c><c t=\"inlineStr\"><is><t>B2</t></is></c></row>";
    simpleExcelWriter.create(@"D:\tes.xlsx", rowContent);

## Style

this class has 1 style example to make cell accept  new line using "Wrap text", to do that create new line character with `&#13;&#10;` and add `s="1"` inside `<c>` 

**XML code**

    <row>
      <c t="inlineStr" s="1"><is><t>New&#13;&#10;Line</t></is></c>
    </row>

**C# code**

    string rowContent = "<row><c t=\"inlineStr\" s=\"1\"><is><t>New&#13;&#10;Line</t></is></c></row>";
    simpleExcelWriter.create(@"D:\tes.xlsx", rowContent);
