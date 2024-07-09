using Microsoft.Office.Interop.PowerPoint;
using System;

namespace EdoliAddIn
{
    class AnimationTool
    {
        public static void SetNameOfActive(String name)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            if (shapes.Count > 0)
            {
                var shape = shapes[0];
                shape.Name = name;
            }
        }

        public static void FollowAnimation()
        { 
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            var slide = Util.CurrentSlide();

            if (shapes.Count > 1)
            {
                var path = shapes[shapes.Count - 1];
                var nodes = path.Nodes;

                if (nodes.Count == 0)
                {
                    return;
                }

                var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;

                float slideHeight = currentPresentation.PageSetup.SlideHeight;
                float slideWidth = currentPresentation.PageSetup.SlideWidth;

                var firstNode = nodes[1];
                var segmentType = firstNode.SegmentType;
                var x1 = firstNode.Points[1, 1] / slideWidth;
                var y1 = firstNode.Points[1, 2] / slideHeight;

                string pathVml = "M 0 0";

                if (segmentType == Microsoft.Office.Core.MsoSegmentType.msoSegmentCurve)
                {
                    for (int i = 1; i < nodes.Count; i++)
                    {
                        // From second node
                        var node = nodes[i + 1];
                        if (i % 3 == 1)
                        {
                            pathVml += " C ";
                        }
                        else
                        {

                            pathVml += " ";
                        }
                        pathVml += (node.Points[1, 1] / slideWidth - x1) + " " + (node.Points[1, 2] / slideHeight - y1);
                    }
                } 
                else
                {
                    for (int i = 1; i < nodes.Count; i++)
                    {
                        // From second node
                        var node = nodes[i + 1];
                        pathVml += " L " + (node.Points[1, 1] / slideWidth - x1) + " " + (node.Points[1, 2] / slideHeight - y1);
                    }
                }

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    var sequence = slide.TimeLine.MainSequence;

                    var motionEffect = sequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectPathZigzag);
                    motionEffect.Behaviors[1].MotionEffect.Path = pathVml;
                }
            }
        }
    }
}
