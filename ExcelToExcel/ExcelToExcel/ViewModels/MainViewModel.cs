
using ExcelToExcel.Commands;
using ExcelToExcel.Models;
using System;
using System.IO;

namespace ExcelToExcel.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private string inputFilename;
        private string outputFilename;
        private string message;
        private EspeceXL especes;

        public string InputFilename
        {
            get { return inputFilename; }
            set { 
                inputFilename = value;
                OnPropertyChanged();

                /// Commentaire pédagogique
                /// Sert à envoyer un signal au UI pour valider si
                /// la commande peut être exécuté
                ValidateExcelCommand.RaiseCanExecuteChanged();
                LoadContentCommand.RaiseCanExecuteChanged();
            }
        }

        public string OutputFilename
        {
            get { return outputFilename; }
            set { 
                outputFilename = value;
                OnPropertyChanged();
            }
        }

        

        /// <summary>
        /// Utiliser cette propriété pour passer un message à l'utilisateur
        /// </summary>
        public string Message
        {
            get { return message; }
            set {
                message = value;
                OnPropertyChanged();
            }
        }


        private string fileContent;

        public string FileContent
        {
            get { return fileContent; }
            set {
                fileContent = value;
                OnPropertyChanged();
            }
        }


        public DelegateCommand<string> ValidateExcelCommand { get; set; }
        public DelegateCommand<string> LoadContentCommand { get; set; }
        public DelegateCommand<string> TestCommand { get; set; }

        public MainViewModel()
        {
            initCommands();
        }

        private void initCommands()
        {
            ValidateExcelCommand = new DelegateCommand<string>(ValidateExcel, CanExecuteValidateExcelCommand);
            LoadContentCommand = new DelegateCommand<string>(LoadContent, CanExecuteLoadContentCommand);
            TestCommand = new DelegateCommand<string>(TestAction);
        }

        private bool CanExecuteLoadContentCommand(string obj)
        {
            bool result;

            if (!string.IsNullOrEmpty(obj))
                InputFilename = obj;

            var fileExists = File.Exists(InputFilename);

            if (!fileExists)
            {
                Message = "Fichier inexistant!";
                result = false;
            }
            else
            {
                Message = "";
                result = true;
            }

            return result;
        }

        private void LoadContent(string obj = null)
        {
            try
            {
                especes = new EspeceXL(InputFilename);
                especes.LoadFile();
                FileContent = especes.GetCSV();
            } catch (ArgumentException ex)
            {
                Message = "Mauvais format de fichier!";
            } catch (IOException ex)
            {
                especes.LoadFileReadOnly();
                FileContent = especes.GetCSV();
                Message = "Fichier en lecture seule";
            }
        }

        private void TestAction(string obj)
        {
            
        }

        /// <summary>
        /// Commentaire pédagogique
        /// Cette fonction permet d'indiquer si l'on peut exécuter ou non la commande
        /// On l'utilise principalement pour activer ou désactiver des fonctionnalités
        /// dans le UI
        /// Cette fonction n'est appelé que lorsque la méthode RaiseExecuteChanged() est
        /// appelée
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private bool CanExecuteValidateExcelCommand(string obj)
        {
            return !string.IsNullOrEmpty(InputFilename);
        }

        private void ValidateExcel(string obj)
        {
            throw new NotImplementedException();
        }
    }
}