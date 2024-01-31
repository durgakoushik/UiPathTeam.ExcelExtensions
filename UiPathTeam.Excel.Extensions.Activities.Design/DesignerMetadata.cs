using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using UiPathTeam.Excel.Extensions.Activities.Design.Designers;
using UiPathTeam.Excel.Extensions.Activities.Design.Properties;

namespace UiPathTeam.Excel.Extensions.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ExcelExtensionScope), categoryAttribute);
            builder.AddCustomAttributes(typeof(ExcelExtensionScope), new DesignerAttribute(typeof(ExcelExtensionScopeDesigner)));
            builder.AddCustomAttributes(typeof(ExcelExtensionScope), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(CopyAsPitcure), categoryAttribute);
            builder.AddCustomAttributes(typeof(CopyAsPitcure), new DesignerAttribute(typeof(CopyAsPictureDesigner)));
            builder.AddCustomAttributes(typeof(CopyAsPitcure), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ActivateSheet), categoryAttribute);
            builder.AddCustomAttributes(typeof(ActivateSheet), new DesignerAttribute(typeof(ActivateSheetDesigner)));
            builder.AddCustomAttributes(typeof(ActivateSheet), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(AutoFill), categoryAttribute);
            builder.AddCustomAttributes(typeof(AutoFill), new DesignerAttribute(typeof(AutoFillDesigner)));
            builder.AddCustomAttributes(typeof(AutoFill), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(AutoFit), categoryAttribute);
            builder.AddCustomAttributes(typeof(AutoFit), new DesignerAttribute(typeof(AutoFitDesigner)));
            builder.AddCustomAttributes(typeof(AutoFit), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ChangeFormat), categoryAttribute);
            builder.AddCustomAttributes(typeof(ChangeFormat), new DesignerAttribute(typeof(ChangeFormatDesigner)));
            builder.AddCustomAttributes(typeof(ChangeFormat), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ClearContent), categoryAttribute);
            builder.AddCustomAttributes(typeof(ClearContent), new DesignerAttribute(typeof(ClearDesigner)));
            builder.AddCustomAttributes(typeof(ClearContent), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DeleteSheet), categoryAttribute);
            builder.AddCustomAttributes(typeof(DeleteSheet), new DesignerAttribute(typeof(DeleteSheetDesigner)));
            builder.AddCustomAttributes(typeof(DeleteSheet), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(FillColor), categoryAttribute);
            builder.AddCustomAttributes(typeof(FillColor), new DesignerAttribute(typeof(FillColorDesigner)));
            builder.AddCustomAttributes(typeof(FillColor), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(FindAll), categoryAttribute);
            builder.AddCustomAttributes(typeof(FindAll), new DesignerAttribute(typeof(FindAllDesigner)));
            builder.AddCustomAttributes(typeof(FindAll), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(FontStyle), categoryAttribute);
            builder.AddCustomAttributes(typeof(FontStyle), new DesignerAttribute(typeof(FontStyleDesigner)));
            builder.AddCustomAttributes(typeof(FontStyle), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetLastColumn), categoryAttribute);
            builder.AddCustomAttributes(typeof(GetLastColumn), new DesignerAttribute(typeof(GetLastColumnDesigner)));
            builder.AddCustomAttributes(typeof(GetLastColumn), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetLastRow), categoryAttribute);
            builder.AddCustomAttributes(typeof(GetLastRow), new DesignerAttribute(typeof(GetLastRowDesigner)));
            builder.AddCustomAttributes(typeof(GetLastRow), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetSheets), categoryAttribute);
            builder.AddCustomAttributes(typeof(GetSheets), new DesignerAttribute(typeof(GetSheetsDesigner)));
            builder.AddCustomAttributes(typeof(GetSheets), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetActiveCell), categoryAttribute);
            builder.AddCustomAttributes(typeof(GetActiveCell), new DesignerAttribute(typeof(GetValueDesigner)));
            builder.AddCustomAttributes(typeof(GetActiveCell), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GoTo), categoryAttribute);
            builder.AddCustomAttributes(typeof(GoTo), new DesignerAttribute(typeof(GoToDesigner)));
            builder.AddCustomAttributes(typeof(GoTo), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(InsertIDelete), categoryAttribute);
            builder.AddCustomAttributes(typeof(InsertIDelete), new DesignerAttribute(typeof(InsertIDeleteDesigner)));
            builder.AddCustomAttributes(typeof(InsertIDelete), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(Merge), categoryAttribute);
            builder.AddCustomAttributes(typeof(Merge), new DesignerAttribute(typeof(MergeDesigner)));
            builder.AddCustomAttributes(typeof(Merge), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ReadCell), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadCell), new DesignerAttribute(typeof(ReadCellDesigner)));
            builder.AddCustomAttributes(typeof(ReadCell), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(RefreshAll), categoryAttribute);
            builder.AddCustomAttributes(typeof(RefreshAll), new DesignerAttribute(typeof(RefreshAllDesigner)));
            builder.AddCustomAttributes(typeof(RefreshAll), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ActivateSheet), categoryAttribute);
            builder.AddCustomAttributes(typeof(ActivateSheet), new DesignerAttribute(typeof(ActivateSheetDesigner)));
            builder.AddCustomAttributes(typeof(ActivateSheet), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(SaveWorkbook), categoryAttribute);
            builder.AddCustomAttributes(typeof(SaveWorkbook), new DesignerAttribute(typeof(SaveDesigner)));
            builder.AddCustomAttributes(typeof(SaveWorkbook), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(SelectRange), categoryAttribute);
            builder.AddCustomAttributes(typeof(SelectRange), new DesignerAttribute(typeof(SelectRangeDesigner)));
            builder.AddCustomAttributes(typeof(SelectRange), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(SetZoom), categoryAttribute);
            builder.AddCustomAttributes(typeof(SetZoom), new DesignerAttribute(typeof(SetZoomDesigner)));
            builder.AddCustomAttributes(typeof(SetZoom), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(WriteCell), categoryAttribute);
            builder.AddCustomAttributes(typeof(WriteCell), new DesignerAttribute(typeof(WriteCellDesigner)));
            builder.AddCustomAttributes(typeof(WriteCell), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(CloseSession), categoryAttribute);
            builder.AddCustomAttributes(typeof(WriteCell), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DeleteDuplicateColumns), categoryAttribute);
            builder.AddCustomAttributes(typeof(DeleteDuplicateColumns), new DesignerAttribute(typeof(DeleteDuplicateColumnsDesigner)));
            builder.AddCustomAttributes(typeof(DeleteDuplicateColumns), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DeleteEmptyColumns), categoryAttribute);
            builder.AddCustomAttributes(typeof(DeleteEmptyColumns), new DesignerAttribute(typeof(DeleteEmptyColumnsDesigner)));
            builder.AddCustomAttributes(typeof(DeleteEmptyColumns), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DeleteEmptyRows), categoryAttribute);
            builder.AddCustomAttributes(typeof(DeleteEmptyRows), new DesignerAttribute(typeof(DeleteEmptyRowsDesigner)));
            builder.AddCustomAttributes(typeof(DeleteEmptyRows), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ConvertRangeToHTMLCode), categoryAttribute);
            builder.AddCustomAttributes(typeof(ConvertRangeToHTMLCode), new DesignerAttribute(typeof(ConvertRangeToHTMLCodeDesigner)));
            builder.AddCustomAttributes(typeof(ConvertRangeToHTMLCode), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(SetWidthAndHeight), categoryAttribute);
            builder.AddCustomAttributes(typeof(SetWidthAndHeight), new DesignerAttribute(typeof(SetWidthAndHeightDesigner)));
            builder.AddCustomAttributes(typeof(SetWidthAndHeight), new HelpKeywordAttribute(""));

            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
