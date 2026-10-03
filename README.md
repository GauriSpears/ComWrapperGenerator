# ComWrapperGenerator

**RU** | [EN](#comwrappergenerator-en)

Автоматический генератор COM-обёрток для managed .NET-сборок (C# DLL) **без исходного кода**.  
На выходе — **готовый Class Library-проект**, который сразу собирается (`dotnet build`) и регистрируется через `regasm`.

Позволяет вызывать классы, методы, свойства, индексаторы и события из VBA / Excel / Word / Access и других COM-клиентов.

Совместим с **.NET 8/9/10 SDK** и **Visual Studio 2022/2026 Build Tools**.

---

## Быстрый старт

```bash
# 1. Собрать генератор (.NET 8/9/10)
dotnet new console -n ComWrapperGenerator -f net10.0 -o Generator
# вставить код из ComWrapperGenerator.cs в Program.cs
dotnet build -c Release -o Generator/bin

# 2. Сгенерировать COM-проект из вашей DLL
Generator/bin/ComWrapperGenerator.exe "C:\Libs\MyLib.dll" "C:\Output\MyWrapper" "MyCompany.Wrapper" net48

# 3. Собрать обёртку
cd C:\Output\MyWrapper
dotnet build -c Release

# 4. Зарегистрировать (администратор, разрядность = Office)
regasm bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

В VBA: **Tools → References** → type library → использовать классы.

---

## Что создаёт генератор

```
MyWrapper/
├── GeneratedComWrapper.csproj   # Class Library, готовый к сборке
├── ComWrapper.cs                # сгенерированный код обёрток
├── MyLib.dll                    # копия исходной DLL (для Reference)
├── build.bat                    # сборка + подсказки по regasm
├── build.sh
└── README_GENERATED.md          # краткая инструкция
```

### Параметры командной строки

```text
ComWrapperGenerator <path-to-dll> [output-folder] [namespace] [target-framework]
```

| Аргумент | По умолчанию | Описание |
|----------|--------------|----------|
| `path-to-dll` | — | Исходная managed DLL (обязательный) |
| `output-folder` | `<dll-dir>/ComWrapper` | Куда писать проект |
| `namespace` | `GeneratedComWrapper` | Namespace и часть ProgId |
| `target-framework` | `net48` | `net48` (рекомендуется для Office), `net8.0-windows`, `net10.0-windows` и т.д. |

---

## Совместимость со сборкой

| Среда | Генератор | Обёртка `net48` | Обёртка `net10.0-windows` |
|--------|-----------|-----------------|---------------------------|
| **.NET 8/9/10 SDK** | ✅ | ✅ (нужен targeting pack 4.8) | ✅ |
| **VS 2022 Build Tools** | ✅ | ✅ | ✅ |
| **VS 2026 Build Tools** | ✅ | ✅ (workload .NET Framework 4.8) | ✅ (.NET 10 SDK) |

Для `net48` в Build Tools должен быть установлен **.NET Framework 4.8 targeting pack**.  
Для Office обычно надёжнее **`net48` + regasm**, чем современный TFM + comhost.

```bash
# Обёртка сразу под .NET 10
ComWrapperGenerator.exe "C:\Libs\MyLib.dll" "C:\Out" "MyCompany.Wrapper" net10.0-windows
```

---

## Что работает

| Возможность | Как реализовано | Пример в VBA |
|-------------|-----------------|--------------|
| **Классы с конструктором без параметров** | Обычный `New` | `Set obj = New MyWrapper.MyClass` |
| **Классы с параметрами в конструкторе** | Фабричный метод `Create(...)` | `Set obj = factory.Create("conn", 30)` |
| **Защита до Create** | `EnsureInitialized()` → понятная ошибка | — |
| **Instance-методы и свойства** | Проксирование через `_inner` | `obj.DoWork()`, `obj.Name = "..."` |
| **Статические методы и свойства** | Instance-члены обёртки (Create не нужен) | `utils.Calculate(1, 2)` |
| **Перегрузки методов** | Переименование: `DoWork`, `DoWork_Int32`, `DoWork_String` | `obj.DoWork_Int32(42)` |
| **Индексаторы** | `GetItem` / `SetItem` (+ суффиксы при перегрузках) | `obj.GetItem(0)`, `obj.SetItem 0, x` |
| **`List<T>`, `IEnumerable<T>`, `IList<T>`, `ICollection<T>`, массивы** | Конвертация в/из `object()` | `arr = obj.GetItems()` |
| **`out` / `ref` параметры** | Поддерживаются | `obj.TryParse(s, result)` |
| **События с аргументами** | `ComSourceInterfaces` + делегаты | `Private WithEvents o As ...` |
| **Сложные объекты из той же DLL** | Рекурсивные обёртки + конвертация | `Set child = obj.GetChild()` |

### Индексаторы

| В C# | В обёртке |
|------|-----------|
| `this[int index]` | `GetItem(int index)` / `SetItem(int index, value)` |
| `this[string key]` (2-я перегрузка) | `GetItem_String` / `SetItem_String` |
| `this[int i, int j]` | `GetItem_Int32_Int32` / `SetItem_Int32_Int32` |

```vb
v = obj.GetItem(0)
obj.SetItem 0, "hello"
v = obj.GetItem_String("name")
```

### Как это устроено

1. Загрузка DLL через **Reflection**.
2. Поиск публичных не-generic классов и рекурсивный сбор зависимостей.
3. Генерация для каждого класса:
   - интерфейс `IMyClass` (`InterfaceIsDual`);
   - при событиях — `IMyClassEvents` (`InterfaceIsIDispatch`);
   - прокси с `[ProgId]`, `Create`, методами/свойствами/`GetItem`/`SetItem`.
4. Запись полного **`.csproj`** (`ComVisible`, `EnableComHosting`, для `net48` — `RegisterForComInterop`, Reference на DLL).
5. Конвертация на границе COM ↔ .NET: примитивы напрямую, коллекции → `object[]`, сложные типы → обёртки (`Inner`).

### Пример VBA

```vb
Dim factory As MyCompany.Wrapper.Calculator
Dim calc As MyCompany.Wrapper.ICalculator

Set factory = New MyCompany.Wrapper.Calculator
Set calc = factory.Create("Server=.;Database=Test", 30)

Debug.Print calc.Add(10, 20)
Debug.Print calc.GetItem(0)

Dim obj As Object
Set obj = CreateObject("MyCompany.Wrapper.Calculator")
Set obj = obj.Create("Server=.;Database=Test", 30)
```

---

## Что не работает (ограничения COM)

| Возможность | Почему нельзя |
|-------------|---------------|
| **Открытые generic-классы** (`MyClass<T>`) | COM не описывает параметр типа `T`. |
| **Закрытые generic-классы** (`MyClass<string>`) | Возможны в теории; генератор не создаёт. |
| **Generic-методы** (`void Foo<T>(T x)`) | Та же причина. |
| **Нативная диспетчеризация перегрузок в VBA** | VBA/`IDispatch` слабо различает перегрузки → **переименование**. |
| **«Настоящий» indexer `obj(i)` везде** | Надёжнее `GetItem`/`SetItem`; parameterized properties поддерживаются не везде. |
| **Произвольные .NET-объекты** | Только типы из той же сборки + примитивы/строки/коллекции. |
| **Async** (`Task`, `async`) | COM синхронный. |

Это ограничения модели COM, не генератора.

---

## Сборка генератора

```bash
dotnet new console -n ComWrapperGenerator -f net10.0 -o .
# Program.cs ← ComWrapperGenerator.cs
dotnet build -c Release
```

Требуется .NET 8/9/10 SDK (или VS Build Tools с соответствующим workload).

---

## Сборка и регистрация обёртки

```bash
cd C:\Output\MyWrapper
dotnet build -c Release
# или build.bat
```

```bat
REM 32-bit Office
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb

REM 64-bit Office
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

Для `net8.0-windows` / `net10.0-windows`: `regsvr32` на `.comhost.dll`, TLB при необходимости через **dscom**.

### VBA

1. **Tools → References**
2. Библиотека или Browse → `.tlb`
3. Раннее или позднее связывание

---

## Требования

- .NET Framework 4.x и/или .NET 8/9/10 там, где крутятся DLL и обёртка
- Совпадение x86/x64 у Office, обёртки и зависимостей
- Права на COM-регистрацию (или `/codebase` + DLL рядом с клиентом)
- Для `net48`: targeting pack **.NET Framework 4.8**

---

## Рекомендации

- Перегенерируйте проект при смене публичного API исходной DLL
- GUID’ы случайные — для релиза зафиксируйте вручную
- Продакшен: strong name, стабильный путь или GAC
- События в VBA: `WithEvents`, контроль lifetime
- Для Excel/Office предпочтителен **`net48`**

---

# ComWrapperGenerator (EN)

Automatic COM wrapper generator for managed .NET assemblies (C# DLLs) **without source code**.  
Produces a **complete Class Library project** ready for `dotnet build` and `regasm`.

Call classes, methods, properties, **indexers**, and events from VBA / Excel / Word / Access and other COM clients.

Compatible with **.NET 8/9/10 SDK** and **Visual Studio 2022/2026 Build Tools**.

---

## Quick start

```bash
# 1. Build the generator (.NET 8/9/10)
dotnet new console -n ComWrapperGenerator -f net10.0 -o Generator
# paste ComWrapperGenerator.cs into Program.cs
dotnet build -c Release -o Generator/bin

# 2. Generate a COM project from your DLL
Generator/bin/ComWrapperGenerator.exe "C:\Libs\MyLib.dll" "C:\Output\MyWrapper" "MyCompany.Wrapper" net48

# 3. Build the wrapper
cd C:\Output\MyWrapper
dotnet build -c Release

# 4. Register (Administrator; bitness must match Office)
regasm bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

In VBA: **Tools → References** → type library → use the types.

---

## What the generator produces

```
MyWrapper/
├── GeneratedComWrapper.csproj   # Class Library, ready to build
├── ComWrapper.cs                # generated wrapper code
├── MyLib.dll                    # copy of the source DLL (for Reference)
├── build.bat                    # build + regasm hints
├── build.sh
└── README_GENERATED.md          # short instructions
```

### Command-line arguments

```text
ComWrapperGenerator <path-to-dll> [output-folder] [namespace] [target-framework]
```

| Argument | Default | Description |
|----------|---------|-------------|
| `path-to-dll` | — | Source managed DLL (required) |
| `output-folder` | `<dll-dir>/ComWrapper` | Project output directory |
| `namespace` | `GeneratedComWrapper` | Namespace and ProgId prefix |
| `target-framework` | `net48` | `net48` (best for Office), `net8.0-windows`, `net10.0-windows`, etc. |

---

## Build compatibility

| Environment | Generator | Wrapper `net48` | Wrapper `net10.0-windows` |
|-------------|-----------|-----------------|---------------------------|
| **.NET 8/9/10 SDK** | ✅ | ✅ (needs 4.8 targeting pack) | ✅ |
| **VS 2022 Build Tools** | ✅ | ✅ | ✅ |
| **VS 2026 Build Tools** | ✅ | ✅ (.NET Framework 4.8 workload) | ✅ (.NET 10 SDK) |

For `net48`, install the **.NET Framework 4.8 targeting pack**.  
For Office, **`net48` + regasm** is usually more reliable than modern TFM + comhost.

```bash
# Wrapper targeting .NET 10
ComWrapperGenerator.exe "C:\Libs\MyLib.dll" "C:\Out" "MyCompany.Wrapper" net10.0-windows
```

---

## What works

| Feature | Implementation | VBA example |
|---------|----------------|-------------|
| **Parameterless constructor** | Plain `New` | `Set obj = New MyWrapper.MyClass` |
| **Parameterized constructors** | Factory `Create(...)` | `Set obj = factory.Create("conn", 30)` |
| **Guard before Create** | `EnsureInitialized()` | — |
| **Instance methods & properties** | Proxied via `_inner` | `obj.DoWork()`, `obj.Name = "..."` |
| **Static methods & properties** | Instance members (no Create) | `utils.Calculate(1, 2)` |
| **Method overloads** | Renamed: `DoWork`, `DoWork_Int32`, … | `obj.DoWork_Int32(42)` |
| **Indexers** | `GetItem` / `SetItem` (+ suffixes) | `obj.GetItem(0)`, `obj.SetItem 0, x` |
| **`List<T>`, `IEnumerable<T>`, `IList<T>`, …, arrays** | ↔ `object()` | `arr = obj.GetItems()` |
| **`out` / `ref`** | Supported | `obj.TryParse(s, result)` |
| **Events with arguments** | `ComSourceInterfaces` | `Private WithEvents o As ...` |
| **Complex objects (same DLL)** | Recursive wrappers | `Set child = obj.GetChild()` |

### Indexers

| C# | Wrapper |
|----|---------|
| `this[int index]` | `GetItem(int)` / `SetItem(int, value)` |
| `this[string key]` (2nd overload) | `GetItem_String` / `SetItem_String` |
| `this[int i, int j]` | `GetItem_Int32_Int32` / `SetItem_Int32_Int32` |

```vb
v = obj.GetItem(0)
obj.SetItem 0, "hello"
v = obj.GetItem_String("name")
```

### How it works

1. Load the DLL via **Reflection**.
2. Discover public non-generic classes; collect dependencies recursively.
3. Emit interfaces, optional event interfaces, and proxy classes with `Create` / members / `GetItem`/`SetItem`.
4. Write a full **`.csproj`** (`ComVisible`, `EnableComHosting`, `RegisterForComInterop` on net48, DLL reference).
5. Convert at the COM boundary: primitives direct, collections → `object[]`, complex types → wrappers (`Inner`).

---

## What does not work (COM limitations)

| Feature | Why |
|---------|-----|
| **Open generics** (`MyClass<T>`) | COM cannot describe `T`. |
| **Closed generics** (`MyClass<string>`) | Not emitted by the generator. |
| **Generic methods** | Same reason. |
| **Native VBA overload resolution** | Methods are **renamed**. |
| **Universal `obj(i)` indexer syntax** | Prefer `GetItem`/`SetItem`. |
| **Arbitrary .NET objects** | Only same-assembly wrapped types + primitives/collections. |
| **Async** | COM is synchronous. |

These are COM model limits, not generator bugs.

---

## Building the generator

```bash
dotnet new console -n ComWrapperGenerator -f net10.0 -o .
# Program.cs ← ComWrapperGenerator.cs
dotnet build -c Release
```

Requires .NET 8/9/10 SDK (or VS Build Tools with the matching workload).

---

## Building and registering the wrapper

```bash
cd C:\Output\MyWrapper
dotnet build -c Release
```

```bat
REM 32-bit Office
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb

REM 64-bit Office
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

For `net8.0-windows` / `net10.0-windows`: `regsvr32` on `.comhost.dll`; TLB via **dscom** if needed.

---

## Requirements

- .NET Framework 4.x and/or .NET 8/9/10 where the DLL and wrapper run
- Matching x86/x64 for Office, wrapper, and dependencies
- COM registration rights (or `/codebase` + DLL beside the client)
- For `net48`: **.NET Framework 4.8 targeting pack**

---

## Recommendations

- Re-generate when the source DLL public API changes
- Pin GUIDs for stable releases
- Production: strong-name; stable path or GAC
- VBA events: `WithEvents` and careful lifetime
- Prefer **`net48`** for Excel/Office scenarios
