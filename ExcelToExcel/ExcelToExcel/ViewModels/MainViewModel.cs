
using ExcelToExcel.Commands;
using System;

namespace ExcelToExcel.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private string inputFilename;
        private string outputFilename;

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

        public DelegateCommand<string> ValidateExcelCommand { get; set; }

        public MainViewModel()
        {
            initCommands();
        }

        private void initCommands()
        {
            ValidateExcelCommand = new DelegateCommand<string>(ValidateExcel, CanExecuteValidateExcelCommand);
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