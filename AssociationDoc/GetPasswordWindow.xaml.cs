using System.Windows;

namespace AssociationDoc
{
    /// <summary>
    /// Логика взаимодействия для GetPasswordWindow.xaml
    /// </summary>
    public partial class GetPasswordWindow : Window
    {
        public GetPasswordWindow(string fileName)
        {
            InitializeComponent();
            TextBlockPassword.Text = "Для файла \"" + fileName + "\"";
        }

        private void Password_Click(object sender, RoutedEventArgs e)
        {
            Password.password = TextBoxPassword.Text;
            Close();
        }
    }
}
