using ExcelToExcel.Models;
using ExcelToExcel.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Xunit;

namespace ExcelToExcel.Tests
{
    public class MainViewModelTests
    {
        
        string excelFilesPath;

        MainViewModel vm;

        public MainViewModelTests()
        {
            vm = new MainViewModel();

            Uri codeBaseUrl = new Uri(Assembly.GetExecutingAssembly().Location);
            var codeBasePath = Uri.UnescapeDataString(codeBaseUrl.AbsolutePath);
            var dirPath = Path.GetDirectoryName(codeBasePath);

            /// Va chercher le dossier Data à partir du dossier de compilation
            /// Adapter selon votre réalité
            excelFilesPath = Path.Combine(dirPath, @"..\..\..\..\..\data");
        }

        private void resetData()
        {
            vm.InputFilename = "";
        }

        #region Tests à compléter

        public void InputFile_IsEmpty_Message_ShouldBe_Empty()
        {
            /// TODO : Q1a. Compléter le test
            /// TODO : Q1b. Ne pas briser la batterie de tests après ce tests
            /// 
            Assert.True(false);
        }

        // TODO : Q2 : Compléter le test CanExecuteSaveCommand_FileNotLoaded_ShouldReturn_False

        // TODO : Q3 : Compléter le test CanExecuteSaveCommand_OutputFileInvalid_ShouldReturn_False

        // TODO : Q4 : Compléter le test CanExecuteSaveCommand_OutputFileValid_ShouldReturn_True(string filename)


        #endregion

        #region Tests corrigés

        [Theory]
        [MemberData(nameof(ExistingFilesTestData))]
        public void LoadContentCommand_CanExecute_FileExists_ShouldReturn_true(string fn)
        {            
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            var actual = vm.LoadContentCommand.CanExecute("");

            /// Assert
            Assert.True(actual);
        }

        [Theory]
        [MemberData(nameof(NonExistentFilesTestData))]
        public void LoadContentCommand_CanExecute_FileNotExists_ShouldReturn_false(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            var actual = vm.LoadContentCommand.CanExecute("");

            /// Assert
            Assert.False(actual);
        }

        [Theory]
        [MemberData(nameof(NonExistentFilesTestData))]
        public void LoadContentCommand_CanExecute_FileNotExists_Message_ShouldBe_FileNonExistent(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;
            var expected = "Fichier inexistant!";

            /// Act
            vm.LoadContentCommand.CanExecute("");
            var actual = vm.Message;

            /// Assert
            Assert.Equal(expected, actual);
        }

        [Theory]
        [MemberData(nameof(BadFileTypesTestData))]
        public void LoadContentCommand_FileWrongFormat_Message_ShouldBe_BadFormat(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;
            var expected = "Mauvais format de fichier!";

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.Message;

            /// Assert
            Assert.Equal(expected, actual);
        }

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void LoadContentCommand_ExcelWrongFormat_Message_ShouldBe_BadFormat(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;
            var expected = "Mauvais format de fichier!";

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.Message;

            /// Assert
            Assert.Equal(expected, actual);
        }

        [Theory]
        [MemberData(nameof(GoodExcelFileTestData))]
        public void LoadContentCommand_GoodExcel_Should_HaveContent(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.FileContent;

            /// Assert
            Assert.NotEqual("", actual);
        }


        [Theory]
        [MemberData(nameof(GoodExcelFileTestData))]
        public void LoadContentCommand_ReadOnlyExcel_ShouldDisplay_Message(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.FileContent;

            /// Assert
            Assert.NotEqual("", actual);
        }

        #endregion


        #region TEST DATA
        public static IEnumerable<object[]> ExistingFilesTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
            new object[] {"invalide_fichier_type.txt"},
        };

        public static IEnumerable<object[]> ExcelFilesTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"liste_especes_multifeuilles.xlsx"},
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
        };

        public static IEnumerable<object[]> BadExcelFilesTestData = new List<object[]>
        {
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
        };

        public static IEnumerable<object[]> NonExistentFilesTestData = new List<object[]>
        {
            new object[] {"a.xlsx"},
            new object[] {"test.txt"},
            new object[] {"fasfdsfa.xlsx"},
            new object[] {"loooolollol.xlsx"},
        };

        public static IEnumerable<object[]> BadFileTypesTestData = new List<object[]>
        {
            new object[] {"invalide_fichier_type.txt"},
        };

        public static IEnumerable<object[]> GoodExcelFileTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"liste_especes_multifeuilles.xlsx"},           
        };


        #endregion
    }
}
