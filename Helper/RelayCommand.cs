using System.Windows.Input;
using DocumentFormat.OpenXml.Bibliography;

namespace FirstProject.Helper
{
    internal class RelayCommand : ICommand
    {
        //it is a method execute
        private readonly Action _execute;
        //it is also a function with return type boolean
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            this._execute = execute;
            this._canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute();
        //This method determines whether the command can execute on the current command target.

        public void Execute(object parameter) => _execute();
       // This method performs the actions associated with the command.








    }
}