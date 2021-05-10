using System;
using System.Windows.Input;

namespace ExcelToExcel.Commands
{
    /// <summary>
    /// Classe qui encapsule l'interface ICommand pour en faciliter son
    /// utilisation
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class DelegateCommand<T> : ICommand
    {
        /// <summary>
        /// Commentaire pédagogique
        /// Un prédicat est une fonction qui retourne un booléen
        /// qui sert de prémisse à l'exécution d'une commande
        /// Si le prédicat est faux, on ne pourra exécuter la commande
        /// </summary>
        private readonly Predicate<T> _canExecute;

        /// <summary>
        /// Commentaire pédagogique
        /// L'action est la méthode qui prend un argument de type T
        /// en paramètre
        /// </summary>
        private readonly Action<T> _execute;

        public DelegateCommand(Action<T> execute)
            : this(execute, null)
        {
        }

        public DelegateCommand(Action<T> execute, Predicate<T> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            if (_canExecute == null)
                return true;

            return _canExecute((T)parameter);
        }

        public void Execute(object parameter)
        {
            _execute((T)parameter);
        }

        //public event EventHandler CanExecuteChanged
        //{
        //    add { CommandManager.RequerySuggested += value; }
        //    remove { CommandManager.RequerySuggested -= value; }
        //}

        public event EventHandler CanExecuteChanged;
        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
