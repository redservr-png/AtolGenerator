using System.Text;
using System.Windows;
using AtolGenerator.Models;
using AtolGenerator.Services;
using AtolGenerator.Views;

namespace AtolGenerator.ViewModels;

public sealed class CorrectionPunchOutcome
{
    public string StatusText { get; init; } = string.Empty;
    public bool OpenedXmlWindow { get; init; }
    public bool AllowCorrectionXml { get; init; }
    public IReadOnlyList<string> PendingUuids { get; init; } = Array.Empty<string>();
}

public static class CorrectionPunchCoordinator
{
    public static async Task<CorrectionPunchOutcome> AfterGenerateAsync(
        IReadOnlyList<GenerationResult> results,
        AtolCredentials credentials,
        Window? owner,
        Action<string>? status = null)
    {
        var plan = CorrectionPunchPlanner.FromResults(results);
        if (!plan.HasApiReceipts && !plan.HasCorrectionXml)
        {
            return new CorrectionPunchOutcome { StatusText = "Нет чеков для пробития." };
        }

        if (plan.HasApiReceipts)
        {
            if (string.IsNullOrWhiteSpace(credentials.Login) ||
                string.IsNullOrWhiteSpace(credentials.GroupCode))
            {
                Alert(owner,
                    "Заполните логин и код группы АТОЛ в настройках.\n" +
                    "Обычные чеки (возвраты/приходы) нужно пробить через API до загрузки XML коррекций.",
                    "Нет настроек АТОЛ",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return new CorrectionPunchOutcome
                {
                    StatusText = plan.HasCorrectionXml
                        ? "API не настроен: XML коррекций не открыт, чтобы не оставить коррекцию без отмены."
                        : "API не настроен: обычные чеки не отправлены.",
                    AllowCorrectionXml = false,
                };
            }

            if (!ConfirmApiPunch(plan, owner))
            {
                return new CorrectionPunchOutcome
                {
                    StatusText = plan.HasCorrectionXml
                        ? "Отправка API отменена. Окно XML не открыто: сначала нужен успешный возврат."
                        : "Отправка API отменена.",
                    AllowCorrectionXml = false,
                };
            }

            var punch = await PunchReceiptsAsync(plan.ApiReceipts, credentials, status);
            if (punch.Fail > 0)
            {
                var errors = punch.Errors.ToString().TrimEnd();
                Alert(owner,
                    "Обычные чеки не все прошли через API. XML коррекций загружать нельзя: " +
                    "иначе в журнале останется коррекция без отмены исходного чека.\n\n" +
                    errors,
                    "Ошибка пробития",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                ShowPendingUuidWarning(owner, punch.Pending);
                return new CorrectionPunchOutcome
                {
                    StatusText = $"API: успешно {punch.Ok}, ошибок {punch.Fail}. XML коррекций не открыт.",
                    PendingUuids = punch.Pending,
                    AllowCorrectionXml = false,
                };
            }

            if (!plan.HasCorrectionXml)
            {
                return new CorrectionPunchOutcome
                {
                    StatusText = punch.AlreadyProcessed > 0
                        ? $"API: новых {punch.Ok - punch.AlreadyProcessed}, уже в журнале {punch.AlreadyProcessed}."
                        : $"API: пробито {punch.Ok} обычных чеков. Чеков коррекции в пачке нет.",
                    PendingUuids = punch.Pending,
                };
            }
        }

        if (!plan.HasCorrectionXml)
            return new CorrectionPunchOutcome { StatusText = "Нет XML коррекций." };

        OpenXmlUploadWindow(plan, owner);
        var apiNote = plan.HasApiReceipts
            ? "Возвраты ушли в API. "
            : string.Empty;
        return new CorrectionPunchOutcome
        {
            StatusText = $"{apiNote}Открыто окно загрузки XML коррекций ({plan.CorrectionXmlPaths.Count}).",
            OpenedXmlWindow = true,
            AllowCorrectionXml = true,
        };
    }

    public static Window? FindOwner() =>
        Application.Current?.Windows.OfType<Window>().FirstOrDefault(window => window.IsActive)
        ?? Application.Current?.MainWindow;

    public static void OpenXmlUploadWindow(CorrectionPunchPlan plan, Window? owner)
    {
        var window = new XmlUploadWindow(plan);
        if (owner is not null && !ReferenceEquals(owner, window))
            window.Owner = owner;
        window.Show();
        window.Activate();
    }

    private static bool ConfirmApiPunch(CorrectionPunchPlan plan, Window? owner)
    {
        var preview = string.Join(Environment.NewLine, plan.ApiReceipts.Take(8)
            .Select(result =>
                $"• {result.OrderNum}: {result.Amount:N2} ₽ ({OperationLabel(result.CheckData?.OperationType)})"));
        if (plan.ApiReceipts.Count > 8)
            preview += Environment.NewLine + $"• и ещё {plan.ApiReceipts.Count - 8}";

        var xmlNote = plan.HasCorrectionXml
            ? "\nПосле успеха откроется одно окно кабинета АТОЛ для загрузки XML коррекций. " +
              "Файлы возвратов в это окно загружать не нужно."
            : string.Empty;

        var message =
            "Обычные чеки (возвраты, приходы, расходы) будут отправлены через API АТОЛ Online.\n" +
            "Чеки коррекции через API не проходят — для них только XML.\n\n" +
            $"Количество: {plan.ApiReceipts.Count}\n\n" +
            $"{preview}\n" +
            xmlNote +
            "\n\nЕсли возврат не пройдёт, XML коррекций загружать нельзя.";

        return Alert(owner, message, "Подтверждение пробития",
            MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) == MessageBoxResult.Yes;
    }

    private static async Task<(int Ok, int Fail, int AlreadyProcessed, StringBuilder Errors, List<string> Pending)>
        PunchReceiptsAsync(
            IReadOnlyList<GenerationResult> receipts,
            AtolCredentials credentials,
            Action<string>? status)
    {
        var errors = new StringBuilder();
        var pending = new List<string>();
        var ok = 0;
        var fail = 0;
        var already = 0;

        for (var i = 0; i < receipts.Count; i++)
        {
            var result = receipts[i];
            var check = result.CheckData!;
            status?.Invoke($"Пробиваем через API {i + 1} из {receipts.Count}: {result.OrderNum}...");
            var punch = await AtolApiService.PunchCheckDataAsync(credentials, check, result.OrderNum);
            if (punch.Success)
            {
                ok++;
                if (punch.AlreadyProcessed) already++;
                continue;
            }

            fail++;
            errors.AppendLine($"❌ {result.OrderNum}: {punch.Error}");
            if (string.Equals(punch.Status, "wait", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(punch.Uuid))
                pending.Add($"{result.OrderNum}: {punch.Uuid}");
        }

        return (ok, fail, already, errors, pending);
    }

    private static void ShowPendingUuidWarning(Window? owner, IReadOnlyList<string> pending)
    {
        if (pending.Count == 0) return;

        var list = string.Join(Environment.NewLine, pending.Take(10));
        if (pending.Count > 10)
            list += Environment.NewLine + $"• и ещё {pending.Count - 10}";

        Alert(owner,
            "АТОЛ принял документ, но статус за 20 секунд не пришёл.\n" +
            "Чек мог уже фискализироваться. Не отправляйте его повторно и не грузите XML коррекции, пока не проверите журнал.\n\n" +
            "Проверьте журнал АТОЛ Online по UUID:\n" +
            list,
            "Статус чека не получен",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private static MessageBoxResult Alert(
        Window? owner,
        string message,
        string caption,
        MessageBoxButton buttons,
        MessageBoxImage icon,
        MessageBoxResult defaultResult = MessageBoxResult.OK)
    {
        return owner is null
            ? MessageBox.Show(message, caption, buttons, icon, defaultResult)
            : MessageBox.Show(owner, message, caption, buttons, icon, defaultResult);
    }

    private static string OperationLabel(string? operation) => operation switch
    {
        "sell" => "приход",
        "sell_refund" => "возврат прихода",
        "buy" => "расход",
        "buy_refund" => "возврат расхода",
        _ => operation ?? "чек",
    };
}
