using ExcelToExcel.Models;
using ExcelToExcel.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Xunit;

namespace ExcelToExcel.Tests
{
    public class EspeceXLTests
    {
        MainViewModel vm;
        string excelFilesPath;

        public EspeceXLTests()
        {
            vm = new MainViewModel();

            Uri codeBaseUrl = new Uri(Assembly.GetExecutingAssembly().Location);
            var codeBasePath = Uri.UnescapeDataString(codeBaseUrl.AbsolutePath);
            var dirPath = Path.GetDirectoryName(codeBasePath);

            /// Va chercher le dossier Data à partir du dossier de compilation
            /// Adapter selon votre réalité
            excelFilesPath = Path.Combine(dirPath, @"..\..\..\..\..\data");
        }

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void GetCSV_WrongFileContent_Should_Fail(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);
            especeXL.LoadFile();

            /// Act
            /// 
            Action act = () => especeXL.GetCSV();

            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        [Theory]
        [MemberData(nameof(GoodExcelFileTestData))]
        public void GetCSV_GoodFileContent_Should_Pass(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);
            especeXL.LoadFile();
            var notExpected = "";

            /// Act
            /// 
            var actual = especeXL.GetCSV();

            /// Assert
            /// 
            Assert.NotEqual(notExpected, actual);
        }

        [Fact]        
        public void LoadFile_ShouldFail_When_NoFile()
        {
            /// Arrange
            /// 
            var especeXL = new EspeceXL("");
            
            /// Act
            /// 
            Action act = () => especeXL.LoadFile();

            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        [Theory]
        [InlineData("invalide_fichier_type.txt")]
        public void LoadFile_ShouldFail_When_BadFile(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);

            /// Act
            /// 
            Action act = () => especeXL.LoadFile();

            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        // TODO : Q05 : Créez le test « SaveCSV_BadFileName_Should_Fail »

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void SaveCSV_BadFileName_Should_Fail(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);

            /// Act
             Action act = () => especeXL.SaveCSV(filename);


            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        // TODO : Q06 : Créez le test « SaveJson_BadFileName_Should_Fail »

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void SaveJson_BadFileName_Should_Fail(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);

            /// Act
            Action act = () => especeXL.SaveJson(filename);


            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        // TODO : Q07 : Créez le test « SaveXls_BadFileName_Should_Fail »

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void SaveXls_BadFileName_Should_fail(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);

            /// Act
            Action act = () => especeXL.SaveXls(filename);


            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void SaveToFile_BadFileName_Should_Fail(string fn)
        {
            /// Arrange
            /// 
            var filename = Path.Combine(excelFilesPath, fn);
            var especeXL = new EspeceXL(filename);

            /// Act
            Action act = () => especeXL.SaveToFile(filename);


            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        [Theory]
        [InlineData(".txt")]
        [InlineData(".exe")]
        public void SaveToFile_BadExtension_Should_Fai(string ext)
        {
            /// Arrange
            /// 
            var extention = Path.Combine(excelFilesPath, ext);
            var especeXL = new EspeceXL(extention);

            /// Act
            Action act = () => especeXL.SaveToFile(extention);


            /// Assert
            /// 
            Assert.Throws<ArgumentException>(act);
        }

        public static IEnumerable<object[]> BadExcelFilesTestData = new List<object[]>
        {
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
        };

        public static IEnumerable<object[]> GoodExcelFileTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"liste_especes_multifeuilles.xlsx"},
        };
    }
}
