# Разработка, сборка и выпуск

## Требования

- Windows 10/11;
- .NET 8 SDK;
- Visual Studio с workload для .NET Desktop или обычный `dotnet` CLI;
- WebView2 Runtime для встроенных кабинетов;
- клиент 1С нужной разрядности для COM-подключения.

## Локальная сборка

```powershell
dotnet restore
dotnet build AtolGenerator.csproj -c Debug
```

Результат: `bin/Debug/net8.0-windows/AtolGenerator.exe`.

## Публикация одного exe

```powershell
dotnet publish AtolGenerator.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish/x64
dotnet publish AtolGenerator.csproj -c Release -r win-x86 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish/x86
```

`x64` используется по умолчанию. `x86` нужен, если установлен только 32-разрядный
COM-коннектор 1С.

## Версия

Перед выпуском одинаковое значение задаётся в `AtolGenerator.csproj`:

- `Version`;
- `AssemblyVersion`;
- `FileVersion`;
- `InformationalVersion`.

Тег Git имеет форму `v2.2`. Переходы версий выполняются последовательно;
переход на новый major выполняется только по отдельному решению владельца проекта.

## GitHub Actions

Workflow `.github/workflows/release.yml`:

1. собирает self-contained x64 и x86;
2. сохраняет оба exe как artifacts;
3. при теге `v*` создаёт GitHub Release;
4. прикладывает `AtolGenerator-x64.exe` и `AtolGenerator-x86.exe`.

## Проверки перед выпуском

```powershell
dotnet build AtolGenerator.csproj -c Debug
dotnet publish AtolGenerator.csproj -c Release -r win-x64 --self-contained true
dotnet publish AtolGenerator.csproj -c Release -r win-x86 --self-contained true
git diff --check
```

Дополнительно вручную проверяются: запуск exe, открытие отдельного окна
исправлений, блокировка API для коррекций, XML по всем сценариям и импорт отчётов.
