using System.Collections.Generic;

namespace DocSharp.Binary.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(188)]
    public class DoubleWaveType : ShapeType
    {
        public DoubleWaveType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@43@0c@42@1@41@3@40@0@39@1@38@3@37@0l@30@4c@31@5@32@6@33@4@34@5@35@6@36@4xe";
            this.AdjustmentValues = "1404,10800";
            this.ConnectorLocations = "@40,@0;@51,10800;@33,@4;@50,10800";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@46,@48,@47,@49";

            this.Formulas = new List<string>();           
            this.Formulas.Add("val #0"); 
            this.Formulas.Add("prod @0 41 9");
            this.Formulas.Add("prod @0 23 9");
            this.Formulas.Add("sum 0 0 @2");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum 21600 0 @1 ");
            this.Formulas.Add("sum 21600 0 @3 ");
            this.Formulas.Add("sum #1 0 10800 ");
            this.Formulas.Add("sum 21600 0 #1 ");
            this.Formulas.Add("prod @8 1 3 ");
            this.Formulas.Add("prod @8 2 3 ");
            this.Formulas.Add("prod @8 4 3 ");
            this.Formulas.Add("prod @8 5 3 ");
            this.Formulas.Add("prod @8 2 1 ");
            this.Formulas.Add("sum 21600 0 @9 ");
            this.Formulas.Add("sum 21600 0 @10 ");
            this.Formulas.Add("sum 21600 0 @8 ");
            this.Formulas.Add("sum 21600 0 @11 ");
            this.Formulas.Add("sum 21600 0 @12 ");
            this.Formulas.Add("sum 21600 0 @13 ");
            this.Formulas.Add("prod #1 1 3 ");
            this.Formulas.Add("prod #1 2 3 ");
            this.Formulas.Add("prod #1 4 3 ");
            this.Formulas.Add("prod #1 5 3 ");
            this.Formulas.Add("prod #1 2 1 ");
            this.Formulas.Add("sum 21600 0 @20"); 
            this.Formulas.Add("sum 21600 0 @21 ");
            this.Formulas.Add("sum 21600 0 @22 ");
            this.Formulas.Add("sum 21600 0 @23 ");
            this.Formulas.Add("sum 21600 0 @24 ");
            this.Formulas.Add("if @7 @19 0 ");
            this.Formulas.Add("if @7 @18 @20 ");
            this.Formulas.Add("if @7 @17 @21 ");
            this.Formulas.Add("if @7 @16 #1 ");
            this.Formulas.Add("if @7 @15 @22 ");
            this.Formulas.Add("if @7 @14 @23 ");
            this.Formulas.Add("if @7 21600 @24 ");
            this.Formulas.Add("if @7 0 @29 ");
            this.Formulas.Add("if @7 @9 @28 ");
            this.Formulas.Add("if @7 @10 @27 ");
            this.Formulas.Add("if @7 @8 @8 ");
            this.Formulas.Add("if @7 @11 @26 ");
            this.Formulas.Add("if @7 @12 @25 ");
            this.Formulas.Add("if @7 @13 21600 ");
            this.Formulas.Add("sum @36 0 @30 ");
            this.Formulas.Add("sum @4 0 @0 ");
            this.Formulas.Add("max @30 @37 ");
            this.Formulas.Add("min @36 @43 ");
            this.Formulas.Add("prod @0 2 1 ");
            this.Formulas.Add("sum 21600 0 @48"); 
            this.Formulas.Add("mid @36 @43 ");
            this.Formulas.Add("mid @30 @37");


            this.Handles = new List<Handle>();
            var handleOne = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,2229"
            };
            this.Handles.Add(handleOne);

            var handleTwo = new Handle
            {
                position = "#1,bottomRight",
                xrange = "8640,12960"
            };
            this.Handles.Add(handleTwo); 

        }
    }
}
