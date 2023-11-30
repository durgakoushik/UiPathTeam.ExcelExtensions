using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace UiPathTeam.Excel.Extensions.Activities.Design.Designers
{
    // Interaction logic for ChangeFormatDesigner.xaml
    public partial class ChangeFormatDesigner
    {
        public ChangeFormatDesigner()
        {
            InitializeComponent();
            cbFormat.ItemsSource = EnumVal<ChangeFormat.format>.GetDirectionValues();
            cbFormat.DisplayMemberPath = nameof(EnumVal.Name);
            cbFormat.SelectedValuePath = nameof(EnumVal.Value);
        }
    }
}
