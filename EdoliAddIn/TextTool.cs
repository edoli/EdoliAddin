using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using Expressive;

namespace EdoliAddIn
{
    public static class TextTool
    {
        public static void IncreaseNumber()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ReplaceSelectedText(s => (float.Parse(s) + 1).ToString());
        }

        public static void DecreaseNumber()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ReplaceSelectedText(s => (float.Parse(s) - 1).ToString());
        }

        public static void EvaluateExpression()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ReplaceSelectedText(s => {
                return new Expression(s, ExpressiveOptions.IgnoreCaseForParsing).Evaluate().ToString();
            });
        }

        public static String GetSelectedText()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection != null && selection.TextRange != null)
            {
                return selection.TextRange.Text;
            }
            return "";
        }

        public static void ReplaceSelectedText(Func<string, string> replacer)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            var isShapeSelection = (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes);

            if (isShapeSelection)
            {
                var shapeRange = selection.ShapeRange;
                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    var text = shape.TextFrame.TextRange.Text;
                    shape.TextFrame.TextRange.Text = replacer(text);
                }

            }
            else
            {
                if (selection != null && selection.TextRange != null)
                {
                    try
                    {
                        if (selection.TextRange.Text == "")
                        {
                            var shapes = Util.ListSelectedShapes();
                            var shape = shapes[0];
                            var text = shape.TextFrame.TextRange.Text;
                            shape.TextFrame.TextRange.Text = replacer(text);
                        } else
                        {
                            var text = selection.TextRange.Text;
                            selection.TextRange.Text = replacer(text);
                        }
                    }
                    catch
                    {

                    }
                }
            }
        }
    }
}
