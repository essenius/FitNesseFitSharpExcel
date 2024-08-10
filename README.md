# FitNesseFitSharpExcel
This repo contains a fixture to enable testing of Excel spreadsheets along with a number of demo FitNesse pages.
It's a bit different from the other fixtures since it is based on .NET Framework 4.5, while all others use .NET 5.
This is because the Excel fixture depends on functionality that is not available in .NET Core (Office Interop Assemblies).

# Installation
The steps to install are very similar to that of installing the [FibonacciDemo](../../../FitNesseFitSharpFibonacciDemo).

Differences are:
* Download the repo code as a zip file and extract the contents of the folder `FitNesseFitSharpExcel-master`. 
* Since the fixture uses .NET Framework, you cannot build with the `dotnet` command. Use Visual Studio or msbuild instead, and take the contents from `ExcelFixtureTest\bin\Release`. Or grab the binaries from the [releases](../../releases) and put them in `%LOCALAPPDATA%\FitNesse\ExcelFixture\ExcelFixture\bin\Release`.
* Go to folder: `cd /D %LOCALAPPDATA%\FitNesse\ExcelFixture\ExcelFixture\bin\release`.
* Run the suite: Open a browser and enter the URL http://localhost:8080/FitSharpDemos.ExcelSuite?suite.

# Tutorial and Reference
See the [Wiki](../../wiki)

# Contribute
Enter an [issue](../../issues) or provide a [pull request](../../pulls). 
