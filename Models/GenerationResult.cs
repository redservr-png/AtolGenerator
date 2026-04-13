using System.IO;
using AtolGenerator.Services;

namespace AtolGenerator.Models;

public class GenerationResult
{
    public string     OrderNum    { get; set; } = string.Empty;
    public double     Amount      { get; set; }
    public string     XmlPath     { get; set; } = string.Empty;
    public string?    DocxPath    { get; set; }
    public string     BaseName    { get; set; } = string.Empty;
    public CheckData? CheckData   { get; set; }  // для отложенной записи XML
    public string  XmlFilename  => Path.GetFileName(XmlPath);
    public string  DocxFilename => DocxPath is not null ? Path.GetFileName(DocxPath) : string.Empty;
    public bool    HasDocx      => DocxPath is not null;
}
