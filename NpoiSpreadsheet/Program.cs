using NPOI.SS.UserModel;
using System.Reflection;

using var stream = GetTemplateStream();
using var workbook = WorkbookFactory.Create(stream);

var sheet = workbook.GetSheetAt(0);

for (var i = 1; i <= 34; i++)
{
    sheet.AutoSizeColumn(i);
}

return;

static Stream? GetTemplateStream()
{
    var assembly = Assembly.GetExecutingAssembly();
    var resourcePath = assembly.GetManifestResourceNames()
        .Single(str => str.EndsWith("Demo.xlsx"));

    var stream = assembly.GetManifestResourceStream(resourcePath);
    return stream;
}