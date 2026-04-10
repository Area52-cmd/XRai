# Manual Greenfield Setup (when `xrai init` is unavailable)

Use this when `XRai.Tool.exe init` is not available in your installed skill build.

## 1. Create the project

```bash
dotnet new classlib -n MyAddin -f net8.0-windows
cd MyAddin
rm Class1.cs
```

## 2. Replace the .csproj

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn" Version="1.9.0" />
    <PackageReference Include="XRai.Hooks" Version="1.0.*" />
  </ItemGroup>
</Project>
```

## 3. Create AddInEntry.cs

```csharp
using ExcelDna.Integration;
using XRai.Hooks;

public class AddInEntry : IExcelAddIn
{
    public void AutoOpen()
    {
        Pilot.Start();
        var vm = new MainViewModel();
        Pilot.ExposeModel(vm);
        var pane = new MainTaskPane { DataContext = vm };
        Pilot.Expose(pane);
    }

    public void AutoClose()
    {
        Pilot.Stop();
    }
}
```

## 4. Create MainTaskPane.xaml

```xml
<UserControl x:Class="MyAddin.MainTaskPane"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             Width="420" Background="#1a1a2e">
    <StackPanel Margin="16">
        <TextBox x:Name="InputBox" />
        <Button x:Name="GoButton" Content="Go" Margin="0,8,0,0" Command="{Binding GoCommand}" />
        <TextBlock x:Name="ResultLabel" Text="{Binding Result}" Margin="0,8,0,0" Foreground="White" />
    </StackPanel>
</UserControl>
```

## 5. Create MainViewModel.cs

```csharp
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

public class MainViewModel : INotifyPropertyChanged
{
    private string _result = "";
    public string Result
    {
        get => _result;
        set { _result = value; OnPropertyChanged(); }
    }

    public ICommand GoCommand { get; }

    public MainViewModel()
    {
        GoCommand = new RelayCommand(() => Result = "Hello from XRai!");
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
```

## 6. Build and verify

```bash
dotnet build
```

Output: `bin/Debug/net8.0-windows/publish/MyAddin-AddIn-packed.xll`

Load in Excel, then verify:
```json
{"cmd":"connect"}
{"cmd":"pane"}
{"cmd":"model"}
```
