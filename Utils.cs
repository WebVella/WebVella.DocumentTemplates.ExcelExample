using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebVella.DocumentTemplates.ExcelExample;
public class Utils
{
	public MemoryStream? LoadWorkbookAsMemoryStream(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"{fileName}");
		var fi = new FileInfo(path);
		var templateWB = new XLWorkbook(path);
		MemoryStream? ms = new();
		templateWB.SaveAs(ms);
		return ms;
	}
	public void SaveWorkbookFromMemoryStream(MemoryStream content, string fileName)
	{
		DirectoryInfo? debugFolder = Directory.GetParent(Environment.CurrentDirectory);
		if (debugFolder is null) throw new Exception("debugFolder not found");
		var projectFolder = debugFolder.Parent?.Parent;
		if (projectFolder is null) throw new Exception("projectFolder not found");

		var path = Path.Combine(projectFolder.FullName, $"result-{fileName}");
		using (FileStream fs = new FileStream(path, FileMode.Create))
		{
			content.WriteTo(fs);
		}
	}
}
