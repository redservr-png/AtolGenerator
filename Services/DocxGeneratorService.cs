using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using AtolGenerator.Constants;

namespace AtolGenerator.Services;

public static class DocxGeneratorService
{
    // 1 cm = 567 twips (1440 twips/inch ÷ 2.54 cm/inch)
    private const uint Cm2   = 1134;   // 2 cm
    private const uint Cm3   = 1701;   // 3 cm
    private const uint Cm15  = 851;    // 1.5 cm

    public static void Generate(MemoData memo, string filepath)
    {
        using var doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();

        var body = new Body();

        var sectPr = new SectionProperties(
            new PageMargin
            {
                Top    = (Int32Value)(int)Cm2,
                Bottom = (Int32Value)(int)Cm2,
                Left   = Cm3,
                Right  = Cm15,
            });

        // ── БЛОК 1: Шапка (правый угол) ──
        body.Append(MakePara(AppConstants.RecipientName,      align: JustificationValues.Right, spaceAfter: 20));
        body.Append(MakePara($"от {AppConstants.FromPosition}", align: JustificationValues.Right, spaceAfter: 20));
        body.Append(MakePara(AppConstants.FromName,             align: JustificationValues.Right, spaceAfter: 20));
        body.Append(MakePara(AppConstants.City,                 align: JustificationValues.Right, spaceAfter: 20));

        // ── Заголовок ──
        body.Append(MakePara("Служебная записка.", bold: true, italic: false, fontSize: 28,
                              align: JustificationValues.Center, spaceBefore: 120, spaceAfter: 100));

        // ── БЛОК 2: Дата события + описание операции (курсив) ──
        body.Append(MakePara($"{memo.EventDate} при {memo.OperationDesc}",
                              italic: true, spaceBefore: 40, spaceAfter: 60));

        // ── БЛОК 3: ФИО, Сумма, Документ (курсив) ──
        if (!string.IsNullOrEmpty(memo.CustomerName))
            body.Append(MakePara($"ФИО Покупателя: {memo.CustomerName}", italic: true, spaceAfter: 20));
        body.Append(MakePara($"Сумма: {FormatMoney(memo.Amount)} руб.", italic: true, spaceAfter: 20));
        if (!string.IsNullOrEmpty(memo.OrderInfo))
            body.Append(MakePara($"Документ: {memo.OrderInfo}", italic: true, spaceAfter: 60));

        // ── БЛОК 4: Проблема — не был пробит чек (курсив) ──
        body.Append(MakePara(
            $"Не был пробит кассовый чек на контрольно-кассовом аппарате " +
            $"{AppConstants.KktModel}, заводской номер {AppConstants.KktSerial}, " +
            $"регистрационный номер {AppConstants.KktReg}, режим передачи фискальных данных " +
            $"(формат {AppConstants.FfdVersion}).",
            italic: true, spaceBefore: 0, spaceAfter: 60));

        // ── БЛОК 5: Дата пробития коррекции + описание (курсив) ──
        body.Append(MakePara(
            $"{memo.TodayDate} на контрольно-кассовом аппарате {AppConstants.KktModel}, " +
            $"заводской номер {AppConstants.KktSerial}, регистрационный номер {AppConstants.KktReg}, " +
            $"режим передачи фискальных данных (формат {AppConstants.FfdVersion}) был сформирован " +
            $"{memo.CorrectionDesc} на сумму {FormatMoney(memo.Amount)} руб.",
            italic: true, spaceBefore: 0, spaceAfter: 100));

        // ── БЛОК 6: Прилагаю + подпись ──
        body.Append(MakePara("Копию чека прилагаю к настоящей служебной записке.",
                              italic: true, spaceBefore: 0, spaceAfter: 180));

        body.Append(MakeSignaturePara(memo.TodayDate));

        body.Append(sectPr);

        mainPart.Document = new Document(body);
        mainPart.Document.Save();
    }

    /// <summary>Строка подписи: "Дата: XX.XX.XXXX    ПОДПИСЬ    ФИО кассира  Полюшков К.Н."</summary>
    private static Paragraph MakeSignaturePara(string date)
    {
        var ppr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Left },
            new SpacingBetweenLines { Before = "0", After = "60" });

        var rprBase = new RunProperties();
        rprBase.Append(new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        rprBase.Append(new FontSize { Val = "24" });
        rprBase.Append(new Italic());

        // "Дата: XX.XX.XXXX      ПОДПИСЬ            "
        var run1 = new Run((RunProperties)rprBase.CloneNode(true),
            new Text($"Дата: {date}      ПОДПИСЬ            ")
            { Space = SpaceProcessingModeValues.Preserve });

        // "ФИО кассира" — подчёркнутый
        var rprUnderline = (RunProperties)rprBase.CloneNode(true);
        rprUnderline.Append(new Underline { Val = UnderlineValues.Single });
        var run2 = new Run(rprUnderline,
            new Text("ФИО кассира") { Space = SpaceProcessingModeValues.Preserve });

        // "  Полюшков К.Н."
        var run3 = new Run((RunProperties)rprBase.CloneNode(true),
            new Text($"  {AppConstants.CashierShort}") { Space = SpaceProcessingModeValues.Preserve });

        return new Paragraph(ppr, run1, run2, run3);
    }

    private static Paragraph MakePara(
        string text,
        bool   bold        = false,
        bool   italic      = false,
        int    fontSize    = 24,           // half-points: 24 = 12pt
        JustificationValues? align = null,
        int    spaceBefore = 0,
        int    spaceAfter  = 60)
    {
        var actualAlign = align ?? JustificationValues.Both;
        var rpr = new RunProperties();
        rpr.Append(new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        rpr.Append(new FontSize { Val = fontSize.ToString() });
        if (bold)   rpr.Append(new Bold());
        if (italic) rpr.Append(new Italic());

        var ppr = new ParagraphProperties(
            new Justification { Val = actualAlign },
            new SpacingBetweenLines
            {
                Before = spaceBefore.ToString(),
                After  = spaceAfter.ToString(),
            });

        return new Paragraph(ppr,
            new Run(rpr,
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static string FormatMoney(double amount)
    {
        // "13 016,00" — space thousands, comma decimal (ru-RU style)
        return amount.ToString("N2", new System.Globalization.CultureInfo("ru-RU"));
    }
}

public class MemoData
{
    public string EventDate      { get; set; } = string.Empty;  // дата события (ввод пользователя)
    public string TodayDate      { get; set; } = string.Empty;  // дата пробития коррекции (сегодня)
    public string OperationDesc  { get; set; } = string.Empty;
    public string CustomerName   { get; set; } = string.Empty;
    public double Amount         { get; set; }
    public string OrderInfo      { get; set; } = string.Empty;
    public string CorrectionDesc { get; set; } = string.Empty;
}
