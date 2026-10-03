# ComWrapperGenerator

**RU** | [EN](#comwrappergenerator-en)

Автоматический генератор COM-обёрток для managed .NET-сборок (C# DLL) **без исходного кода**.  
На выходе — **готовый Class Library-проект**, который сразу собирается (`dotnet build`) и регистрируется через `regasm`.

Позволяет вызывать классы, методы, свойства и события из VBA / Excel / Word / Access и других COM-клиентов.

---

## Быстрый старт

```bash
# 1. Собрать генератор
dotnet new console -n ComWrapperGenerator -o Generator
# вставить код из ComWrapperGenerator.cs в Program.cs
dotnet build -c Release -o Generator/bin

# 2. Сгенерировать COM-проект из вашей DLL
Generator/bin/ComWrapperGenerator.exe "C:\Libs\MyLib.dll" "C:\Output\MyWrapper" "MyCompany.Wrapper" net48

# 3. Собрать обёртку
cd C:\Output\MyWrapper
dotnet build -c Release

# 4. Зарегистрировать (от имени администратора, разрядность = Office)
regasm bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

В VBA: **Tools → References** → выбрать type library → использовать классы.

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
| `target-framework` | `net48` | `net48` (рекомендуется для Office) или `net8.0-windows` |

---

## Что работает

| Возможность | Как реализовано | Пример в VBA |
|-------------|-----------------|--------------|
| **Классы с конструктором без параметров** | Обычный `New` | `Set obj = New MyWrapper.MyClass` |
| **Классы с параметрами в конструкторе** | Фабричный метод `Create(...)` | `Set obj = factory.Create("conn", 30)` |
| **Защита до Create** | `EnsureInitialized()` → понятная ошибка | — |
| **Instance-методы и свойства** | Проксирование через `_inner` | `obj.DoWork()`, `obj.Name = "..."` |
| **Статические методы и свойства** | Становятся instance-членами (Create не нужен) | `utils.Calculate(1, 2)` |
| **Перегрузки методов** | Переименовываются: `DoWork`, `DoWork_Int32`, `DoWork_String` | `obj.DoWork_Int32(42)` |
| **`List<T>`, `IEnumerable<T>`, `IList<T>`, `ICollection<T>`, массивы** | Конвертация в/из `object()` | `arr = obj.GetItems()` |
| **`out` / `ref` параметры** | Поддерживаются | `obj.TryParse(s, result)` |
| **События с аргументами** | `ComSourceInterfaces` + делегаты | `Private WithEvents o As ...` |
| **Сложные объекты из той же DLL** | Рекурсивная генерация обёрток + конвертация | `Set child = obj.GetChild()` |

### Как это устроено

1. Генератор загружает DLL через **Reflection**.
2. Находит публичные не-generic классы и рекурсивно собирает зависимости.
3. Генерирует для каждого класса:
   - COM-visible интерфейс (`IMyClass`, `InterfaceIsDual`);
   - при событиях — `IMyClassEvents` (`InterfaceIsIDispatch`);
   - класс-прокси с `[ProgId]`, фабрикой `Create` и прокси-методами.
4. Пишет **полный `.csproj`**:
   - `ComVisible`, `EnableComHosting`, для `net48` — `RegisterForComInterop`;
   - Reference на исходную DLL;
   - готовность к `dotnet build`.
5. На границе COM ↔ .NET:
   - простые типы — напрямую;
   - коллекции → `object[]`;
   - сложные объекты → обёртки (через `Inner`).

### Пример использования из VBA

```vb
' Раннее связывание (после Tools → References)
Dim factory As MyCompany.Wrapper.Calculator
Dim calc As MyCompany.Wrapper.ICalculator

Set factory = New MyCompany.Wrapper.Calculator
Set calc = factory.Create("Server=.;Database=Test", 30)

Debug.Print calc.Add(10, 20)

Dim items As Variant
items = calc.GetHistory()   ' object()

' Позднее связывание
Dim obj As Object
Set obj = CreateObject("MyCompany.Wrapper.Calculator")
Set obj = obj.Create("Server=.;Database=Test", 30)
```

---

## Что не работает (ограничения COM)

| Возможность | Почему нельзя |
|-------------|---------------|
| **Открытые generic-классы** (`MyClass<T>`) | COM не поддерживает параметры типа. Type library не может описать `T`. |
| **Закрытые generic-классы** (`MyClass<string>`) | Технически возможны, генератор их не создаёт (редко нужны). |
| **Generic-методы** (`void Foo<T>(T x)`) | Та же причина. |
| **«Умная» диспетчеризация перегрузок в VBA** | VBA/`IDispatch` плохо различает перегрузки по типам → генератор **переименовывает** методы. |
| **Произвольные .NET-объекты** | Только типы из той же сборки (с обёртками) + примитивы/строки/коллекции. |
| **Индексаторы** | Пропускаются. |
| **Async** (`Task`, `async`) | COM-модель синхронная. |

Это ограничения модели COM, а не генератора.

---

## Сборка генератора

```bash
dotnet new console -n ComWrapperGenerator -o .
# Замените Program.cs содержимым ComWrapperGenerator.cs
dotnet build -c Release
```

Требуется .NET 6/8 SDK.

---

## Сборка и регистрация обёртки

```bash
cd C:\Output\MyWrapper
dotnet build -c Release
# или: build.bat
```

Регистрация (Windows, **разрядность regasm = разрядность Office**):

```bat
REM 32-bit Office
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb

REM 64-bit Office
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

Для `net8.0-windows` используйте `EnableComHosting` + `regsvr32` для `.comhost.dll` и при необходимости **dscom** для TLB.

### Подключение в VBA

1. **Tools → References**
2. Найдите библиотеку или Browse → `.tlb`
3. Пишите код с ранним или поздним связыванием (см. выше)

---

## Требования

- .NET Framework 4.x и/или .NET 6/8 на машине с DLL и обёрткой
- Совпадение x86/x64 у Office, обёртки и зависимостей
- Права на регистрацию COM (или `/codebase` + DLL рядом с клиентом)

---

## Рекомендации

- Перегенерируйте проект при изменении публичного API исходной DLL
- GUID’ы случайные при каждом запуске — для релизов зафиксируйте вручную
- Для продакшена: strong name, предсказуемый путь или GAC
- События в VBA: `WithEvents` и контроль времени жизни объектов
- Для Excel/Office чаще надёжнее **net48**, чем современный .NET

---

# ComWrapperGenerator (EN)

Automatic COM wrapper generator for managed .NET assemblies (C# DLLs) **without source code**.  
Output is a **complete Class Library project** ready to `dotnet build` and register with `regasm`.

Call classes, methods, properties and events from VBA / Excel / Word / Access and other COM clients.

---

## Quick start

```bash
# 1. Build the generator
dotnet new console -n ComWrapperGenerator -o Generator
# paste ComWrapperGenerator.cs into Program.cs
dotnet build -c Release -o Generator/bin

# 2. Generate a COM project from your DLL
Generator/bin/ComWrapperGenerator.exe "C:\Libs\MyLib.dll" "C:\Output\MyWrapper" "MyCompany.Wrapper" net48

# 3. Build the wrapper
cd C:\Output\MyWrapper
dotnet build -c Release

# 4. Register (Administrator, bitness must match Office)
regasm bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

In VBA: **Tools → References** → select the type library → use the types.

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
| `target-framework` | `net48` | `net48` (best for Office) or `net8.0-windows` |

---

## What works

| Feature | How it is implemented | VBA example |
|---------|----------------------|-------------|
| **Parameterless constructor** | Plain `New` | `Set obj = New MyWrapper.MyClass` |
| **Parameterized constructors** | Factory method `Create(...)` | `Set obj = factory.Create("conn", 30)` |
| **Guard before Create** | `EnsureInitialized()` throws a clear error | — |
| **Instance methods & properties** | Proxied via `_inner` | `obj.DoWork()`, `obj.Name = "..."` |
| **Static methods & properties** | Exposed as instance members (no Create) | `utils.Calculate(1, 2)` |
| **Method overloads** | Renamed: `DoWork`, `DoWork_Int32`, `DoWork_String` | `obj.DoWork_Int32(42)` |
| **`List<T>`, `IEnumerable<T>`, `IList<T>`, `ICollection<T>`, arrays** | Converted to/from `object()` | `arr = obj.GetItems()` |
| **`out` / `ref` parameters** | Supported | `obj.TryParse(s, result)` |
| **Events with arguments** | `ComSourceInterfaces` + delegates | `Private WithEvents o As ...` |
| **Complex objects from the same DLL** | Recursive wrappers + conversion | `Set child = obj.GetChild()` |

### How it works

1. Loads the DLL via **Reflection**.
2. Discovers public non-generic classes and recursively collects dependencies.
3. Emits for each class:
   - a COM-visible interface (`IMyClass`, `InterfaceIsDual`);
   - an events interface when needed (`IMyClassEvents`, `InterfaceIsIDispatch`);
   - a proxy class with `[ProgId]`, factory `Create`, and proxy members.
4. Writes a full **`.csproj`**:
   - `ComVisible`, `EnableComHosting`, and for `net48` `RegisterForComInterop`;
   - a reference to the source DLL;
   - ready for `dotnet build`.
5. At the COM ↔ .NET boundary:
   - primitives — direct;
   - collections → `object[]`;
   - complex objects → wrappers (via `Inner`).

### VBA usage example

```vb
' Early binding (after Tools → References)
Dim factory As MyCompany.Wrapper.Calculator
Dim calc As MyCompany.Wrapper.ICalculator

Set factory = New MyCompany.Wrapper.Calculator
Set calc = factory.Create("Server=.;Database=Test", 30)

Debug.Print calc.Add(10, 20)

Dim items As Variant
items = calc.GetHistory()   ' object()

' Late binding
Dim obj As Object
Set obj = CreateObject("MyCompany.Wrapper.Calculator")
Set obj = obj.Create("Server=.;Database=Test", 30)
```

---

## What does not work (COM limitations)

| Feature | Why it is not supported |
|---------|-------------------------|
| **Open generic classes** (`MyClass<T>`) | COM has no type parameters; a type library cannot describe `T`. |
| **Closed generic classes** (`MyClass<string>`) | Possible in theory; the generator does not emit them (rarely needed). |
| **Generic methods** (`void Foo<T>(T x)`) | Same reason. |
| **Natural overload resolution in VBA** | VBA/`IDispatch` does not distinguish overloads by type well → methods are **renamed**. |
| **Arbitrary .NET objects** | Only types from the same assembly (with wrappers) plus primitives/strings/collections. |
| **Indexers** | Skipped. |
| **Async** (`Task`, `async`) | COM is synchronous. |

These are fundamental COM limitations, not generator bugs.

---

## Building the generator

```bash
dotnet new console -n ComWrapperGenerator -o .
# Replace Program.cs with ComWrapperGenerator.cs
dotnet build -c Release
```

Requires .NET 6/8 SDK.

---

## Building and registering the wrapper

```bash
cd C:\Output\MyWrapper
dotnet build -c Release
# or: build.bat
```

Registration (Windows, **regasm bitness = Office bitness**):

```bat
REM 32-bit Office
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb

REM 64-bit Office
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe bin\Release\net48\GeneratedComWrapper.dll /codebase /tlb
```

For `net8.0-windows`, use `EnableComHosting` + `regsvr32` on the `.comhost.dll`, and **dscom** for a TLB if needed.

### Using from VBA

1. **Tools → References**
2. Select the library or Browse to the `.tlb`
3. Use early or late binding (see examples above)

---

## Requirements

- .NET Framework 4.x and/or .NET 6/8 where the DLL and wrapper run
- Matching x86/x64 for Office, wrapper, and dependencies
- Permission to register COM (or `/codebase` with the DLL next to the client)

---

## Recommendations

- Re-run the generator when the source DLL public API changes
- GUIDs are random each run — hard-code them for stable releases
- For production: strong-name the assembly; use a predictable path or GAC
- Events in VBA: use `WithEvents` and manage object lifetime carefully
- For Excel/Office, **net48** is usually more reliable than modern .NET
