using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
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
    // Interaction logic for GoToDesigner.xaml
    public partial class GoToDesigner
    {
        public GoToDesigner()
        {
            InitializeComponent();
            cbDirection.ItemsSource = EnumVal<GoTo.direction>.GetDirectionValues();
            cbDirection.DisplayMemberPath = nameof(EnumVal.Name);
            cbDirection.SelectedValuePath = nameof(EnumVal.Value);
        }
    


    }
}
