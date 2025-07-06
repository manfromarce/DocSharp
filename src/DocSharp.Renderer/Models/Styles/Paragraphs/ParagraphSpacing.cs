namespace DocSharp.Renderer.Models.Styles
{ 
    internal class ParagraphSpacing
    {
        public static readonly ParagraphSpacing Default = new ParagraphSpacing(new AutoLineSpacing(), 0, 10);

        public ParagraphSpacing(LineSpacing line, double before, double after)
        {
            this.Line = line;
            this.Before = before;
            this.After = after;
        }

        public LineSpacing Line { get; }
        public double Before { get; }
        public double After { get; }
    }
}
