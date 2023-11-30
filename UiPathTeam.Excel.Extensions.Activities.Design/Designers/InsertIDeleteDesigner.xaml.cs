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
using UiPathTeam.Excel.Extensions.Activities;

namespace UiPathTeam.Excel.Extensions.Activities.Design.Designers
{
    // Interaction logic for InsertIDeleteDesigner.xaml
    public partial class InsertIDeleteDesigner
    {
        public InsertIDeleteDesigner()
        {
            InitializeComponent();
            cbInsertOrDelete.ItemsSource = EnumVal<InsertIDelete.iorD>.GetDirectionValues();
            cbInsertOrDelete.DisplayMemberPath = nameof(EnumVal.Name);
            cbInsertOrDelete.SelectedValuePath = nameof(EnumVal.Value);

            cbRowOrColumn.ItemsSource = EnumVal<InsertIDelete.rorC>.GetDirectionValues();
            cbRowOrColumn.DisplayMemberPath = nameof(EnumVal.Name);
            cbRowOrColumn.SelectedValuePath = nameof(EnumVal.Value);
        }
    }
}
