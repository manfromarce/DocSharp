using System.Collections.Generic;

namespace DocSharp.Binary.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(94)]
    class NotchedRightArrowType : ShapeType
    {
        public NotchedRightArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m@0,l@0@1,0@1@5,10800,0@2@0@2@0,21600,21600,10800xe"; 
            this.Formulas = new List<string>();

            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("sum height 0 #1"); 
            this.Formulas.Add("sum 10800 0 #1"); 
            this.Formulas.Add("sum width 0 #0"); 
            this.Formulas.Add("prod @4 @3 10800"); 
            this.Formulas.Add("sum width 0 @5");

            this.AdjustmentValues = "16200,5400";
            this.ConnectorLocations = "@0,0;@5,10800;@0,21600;21600,10800";
            this.ConnectorAngles = "270,180,90,0"; 

            this.TextboxRectangle = "@5,@1,@6,@2";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1",
                xrange = "0,21600",
                yrange = "0,10800"
            };

            this.Handles.Add(HandleOne);


        }
    }
}
