# FitNesseFitSharpExcel
This repo contains a fixture to enable testing of Excel spreadsheets along with a number of demo FitNesse pages

# Getting Started
1. Download FitNesse (http://fitnesse.org) and install it to C:\Apps\FitNesse
2. Download FitSharp (https://github.com/jediwhale/fitsharp) and install it to C:\Apps\FitNesse\FitSharp.
3. Clone the repo to a local folder (C:\Data\FitNesseDemo)
4. Update plugins.properties to point to the FitSharp folder (if you took other folders than suggested)
5. Build all projects in the solution ExcelFixture (Release)
6. Ensure you have Java installed (1.7 or higher)
7. Start FitNesse with the root repo folder as the data folder as well as the current directory:
	cd /D C:\Data\FitNesseDemo
	java -jar C:\Apps\FitNesse\fitnesse-standalone.jar -d . -e 0
8. Open a browser and enter the URL http://localhost:8080/FitSharpDemos.ExcelSuite?suite

# Notes

I decided not to port this fixture to .NET 5 since it is Windows specific anyway and the Office Interop Assemblies that this fixture depends on don't seem to work reliably in .NET 5.

# Contribute

Enter an issue or provide a pull request.
