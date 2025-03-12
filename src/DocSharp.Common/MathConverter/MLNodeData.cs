using System.Collections.Generic;

namespace DocSharp.MathConverter;

// LaTeX data for inheriting classes. Stored in 'latex_dict.py' in Python version.
internal abstract partial class MLNodeBase
{
    // OMML namespace
    protected const string OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    // Define characters that need to be escaped in LaTeX
    protected static readonly char[] CHARS = { '{', '}', '_', '^', '#', '&', '$', '%', '~' };
    protected const string BACKSLASH = "\\";
    protected const string ALN = "&";

    // Mapping of Unicode characters to LaTeX commands for accents and symbols
    protected static readonly Dictionary<string, string> CHR = new Dictionary<string, string>
        {
            // Unicode : Latex Math Symbols
            // Top accents
            { "\u0300", "\\grave{{{0}}}" },
            { "\u0301", "\\acute{{{0}}}" },
            { "\u0302", "\\hat{{{0}}}" },
            { "\u0303", "\\tilde{{{0}}}" },
            { "\u0304", "\\bar{{{0}}}" },
            { "\u0305", "\\overbar{{{0}}}" },
            { "\u0306", "\\breve{{{0}}}" },
            { "\u0307", "\\dot{{{0}}}" },
            { "\u0308", "\\ddot{{{0}}}" },
            { "\u0309", "\\ovhook{{{0}}}" },
            { "\u030a", "\\ocirc{{{0}}}" },
            { "\u030c", "\\check{{{0}}}" },
            { "\u0310", "\\candra{{{0}}}" },
            { "\u0312", "\\oturnedcomma{{{0}}}" },
            { "\u0315", "\\ocommatopright{{{0}}}" },
            { "\u031a", "\\droang{{{0}}}" },
            { "\u0338", "\\not{{{0}}}" },
            { "\u20d0", "\\leftharpoonaccent{{{0}}}" },
            { "\u20d1", "\\rightharpoonaccent{{{0}}}" },
            { "\u20d2", "\\vertoverlay{{{0}}}" },
            { "\u20d6", "\\overleftarrow{{{0}}}" },
            { "\u20d7", "\\vec{{{0}}}" },
            { "\u20db", "\\dddot{{{0}}}" },
            { "\u20dc", "\\ddddot{{{0}}}" },
            { "\u20e1", "\\overleftrightarrow{{{0}}}" },
            { "\u20e7", "\\annuity{{{0}}}" },
            { "\u20e9", "\\widebridgeabove{{{0}}}" },
            { "\u20f0", "\\asteraccent{{{0}}}" },
            
            // Bottom accents
            { "\u0330", "\\wideutilde{{{0}}}" },
            { "\u0331", "\\underbar{{{0}}}" },
            { "\u20e8", "\\threeunderdot{{{0}}}" },
            { "\u20ec", "\\underrightharpoondown{{{0}}}" },
            { "\u20ed", "\\underleftharpoondown{{{0}}}" },
            { "\u20ee", "\\underledtarrow{{{0}}}" },
            { "\u20ef", "\\underrightarrow{{{0}}}" },
            
            // Over | group
            { "\u23b4", "\\overbracket{{{0}}}" },
            { "\u23dc", "\\overparen{{{0}}}" },
            { "\u23de", "\\overbrace{{{0}}}" },
            
            // Under| group
            { "\u23b5", "\\underbracket{{{0}}}" },
            { "\u23dd", "\\underparen{{{0}}}" },
            { "\u23df", "\\underbrace{{{0}}}" },
        };

    // Mapping of big operators to LaTeX commands
    protected static readonly Dictionary<string, string> CHR_BO = new Dictionary<string, string>
        {
            { "\u2140", "\\Bbbsum" },
            { "\u220f", "\\prod" },
            { "\u2210", "\\coprod" },
            { "\u2211", "\\sum" },
            { "\u222b", "\\int" },
            { "\u222c", "\\iint" },		//mxd. Double integral
            { "\u222d", "\\iiint" },	//mxd. Triple integral
            { "\u222e", "\\oint" },		//mxd. Contour integral
            { "\u22c0", "\\bigwedge" },
            { "\u22c1", "\\bigvee" },
            { "\u22c2", "\\bigcap" },
            { "\u22c3", "\\bigcup" },
            { "\u2a00", "\\bigodot" },
            { "\u2a01", "\\bigoplus" },
            { "\u2a02", "\\bigotimes" },
        };

    // Mapping of various symbols and Greek letters to LaTeX commands
    protected static readonly Dictionary<string, string> T = new Dictionary<string, string>
        {
            // Greek lowercase letters
            { "\U0001d6fc", "\\alpha " },
            { "\U0001d6fd", "\\beta " },
            { "\U0001d6fe", "\\gamma " },
            { "\U0001d6ff", "\\theta " },
            { "\U0001d700", "\\epsilon " },
            { "\U0001d701", "\\zeta " },
            { "\U0001d702", "\\eta " },
            { "\U0001d703", "\\theta " },
            { "\U0001d704", "\\iota " },
            { "\U0001d705", "\\kappa " },
            { "\U0001d706", "\\lambda " },
            { "\U0001d707", "\\m " },
            { "\U0001d708", "\\n " },
            { "\U0001d709", "\\xi " },
            { "\U0001d70a", "\\omicron " },
            { "\U0001d70b", "\\pi " },
            { "\U0001d70c", "\\rho " },
            { "\U0001d70d", "\\varsigma " },
            { "\U0001d70e", "\\sigma " },
            { "\U0001d70f", "\\ta " },
            { "\U0001d710", "\\upsilon " },
            { "\U0001d711", "\\phi " },
            { "\U0001d712", "\\chi " },
            { "\U0001d713", "\\psi " },
            { "\U0001d714", "\\omega " },
            { "\U0001d715", "\\partial " },
            { "\U0001d716", "\\varepsilon " },
            { "\U0001d717", "\\vartheta " },
            { "\U0001d718", "\\varkappa " },
            { "\U0001d719", "\\varphi " },
            { "\U0001d71a", "\\varrho " },
            { "\U0001d71b", "\\varpi " },

            //mxd. Also greek lowercase letters (https://unicodeplus.com/script/Grek)...
            { "\u03b1", "\\alpha " },
            { "\u03b2", "\\beta " },
            { "\u03b3", "\\gamma " },
            { "\u03b4", "\\theta " },
            { "\u03b5", "\\epsilon " },
            { "\u03b6", "\\zeta " },
            { "\u03b7", "\\eta " },
            { "\u03b8", "\\theta " },
            { "\u03b9", "\\iota " },
            { "\u03ba", "\\kappa " },
            { "\u03bb", "\\lambda " },
            { "\u03bc", "\\m " },
            { "\u03bd", "\\n " },
            { "\u03be", "\\xi " },
            { "\u03bf", "\\omicron " },
            { "\u03c0", "\\pi " },
            { "\u03c1", "\\rho " },
            { "\u03c2", "\\varsigma " },
            { "\u03c3", "\\sigma " },
            { "\u03c4", "\\ta " },
            { "\u03c5", "\\upsilon " },
            { "\u03c6", "\\phi " },
            { "\u03c7", "\\chi " },
            { "\u03c8", "\\psi " },
            { "\u03c9", "\\omega " },

            //mxd. Prime
            { "\u0027", "\\prime " },
            
            //mxd. Moodle math operators
            { "\u22c5", "\\cdot " },
            { "\u00d7", "\\times " },
            { "\u002a", "\\ast " },
            { "\u00f7", "\\div " },
            { "\u22c4", "\\diamond " },
            { "\u2295", "\\oplus " },
            { "\u2296", "\\ominus " },
            { "\u2297", "\\otimes " },
            { "\u2298", "\\oslash " },
            { "\u2299", "\\odot " },
            { "\u2218", "\\circ " },
            { "\u2219", "\\bullet " },
            { "\u224d", "\\asymp " },
            { "\u2261", "\\equiv " },
            { "\u2286", "\\subseteq " },
            { "\u2287", "\\supseteq " },
            { "\u2aaf", "\\preceq " },
            { "\u2ab0", "\\succeq " },
            { "\u227c", "\\preccurlyeq " },
            { "\u227d", "\\succcurlyeq " },
            { "\u223c", "\\sim " },
            { "\u2243", "\\simeq " },
            { "\u2248", "\\approx " },
            { "\u2282", "\\subset " },
            { "\u2283", "\\supset " },
            { "\u227a", "\\prec " }, //mxd. Exported as '\prcue' from MSWord...
            { "\u227b", "\\succ " },
            { "\u2200", "\\forall " },
            { "\u2203", "\\exists " },

            // Relation symbols
            { "\u2190", "\\leftarrow " },
            { "\u2191", "\\uparrow " },
            { "\u2192", "\\rightarrow " },
            { "\u2193", "\\downright " },
            { "\u2194", "\\leftrightarrow " },
            { "\u2195", "\\updownarrow " },
            { "\u2196", "\\nwarrow " },
            { "\u2197", "\\nearrow " },
            { "\u2198", "\\searrow " },
            { "\u2199", "\\swarrow " },
            { "\u22ee", "\\vdots " },
            { "\u22ef", "\\cdots " },
            { "\u22f0", "\\adots " },
            { "\u22f1", "\\ddots " },
            { "\u2260", "\\ne " },	//mxd. Called '\neq' in Moodle
            { "\u2264", "\\leq " },
            { "\u2265", "\\geq " },
            { "\u2266", "\\leqq " },
            { "\u2267", "\\geqq " },
            { "\u2268", "\\lneqq " },
            { "\u2269", "\\gneqq " },
            { "\u226a", "\\ll " },
            { "\u226b", "\\gg " },
            { "\u2208", "\\in " },
            { "\u2209", "\\notin " },
            { "\u220b", "\\ni " },
            { "\u220c", "\\nni " },

            //mxd. Double arrows
            { "\u21d0", "\\Leftarrow " },
            { "\u21d1", "\\Uparrow " },
            { "\u21d2", "\\Rightarrow " },
            { "\u21d3", "\\Downarrow " },
            { "\u21d4", "\\Leftrightarrow " },

            //mxd. Long double arrows
            { "\u27f8", "\\Longleftarrow " },
            { "\u27f9", "\\Longrightarrow " },
            { "\u27fa", "\\Longleftrightarrow " },

            // Ordinary symbols
            { "\u221e", "\\infty " },
            
            // Binary relations
            { "\u00b1", "\\pm " },
            { "\u2213", "\\mp " },

            // Italic, Latin, uppercase
            { "\U0001d434", "A" },
            { "\U0001d435", "B" },
            { "\U0001d436", "C" },
            { "\U0001d437", "D" },
            { "\U0001d438", "E" },
            { "\U0001d439", "F" },
            { "\U0001d43a", "G" },
            { "\U0001d43b", "H" },
            { "\U0001d43c", "I" },
            { "\U0001d43d", "J" },
            { "\U0001d43e", "K" },
            { "\U0001d43f", "L" },
            { "\U0001d440", "M" },
            { "\U0001d441", "N" },
            { "\U0001d442", "O" },
            { "\U0001d443", "P" },
            { "\U0001d444", "Q" },
            { "\U0001d445", "R" },
            { "\U0001d446", "S" },
            { "\U0001d447", "T" },
            { "\U0001d448", "U" },
            { "\U0001d449", "V" },
            { "\U0001d44a", "W" },
            { "\U0001d44b", "X" },
            { "\U0001d44c", "Y" },
            { "\U0001d44d", "Z" },
            
            // Italic, Latin, lowercase
            { "\U0001d44e", "a" },
            { "\U0001d44f", "b" },
            { "\U0001d450", "c" },
            { "\U0001d451", "d" },
            { "\U0001d452", "e" },
            { "\U0001d453", "f" },
            { "\U0001d454", "g" },
            { "\U0001d456", "i" },
            { "\U0001d457", "j" },
            { "\U0001d458", "k" },
            { "\U0001d459", "l" },
            { "\U0001d45a", "m" },
            { "\U0001d45b", "n" },
            { "\U0001d45c", "o" },
            { "\U0001d45d", "p" },
            { "\U0001d45e", "q" },
            { "\U0001d45f", "r" },
            { "\U0001d460", "s" },
            { "\U0001d461", "t" },
            { "\U0001d462", "u" },
            { "\U0001d463", "v" },
            { "\U0001d464", "w" },
            { "\U0001d465", "x" },
            { "\U0001d466", "y" },
            { "\U0001d467", "z" },
        };

    // Mapping of functions to their LaTeX representations
    protected static readonly Dictionary<string, string> FUNC = new Dictionary<string, string>
        {
            { "sin", "\\sin({fe})" },
            { "cos", "\\cos({fe})" },
            { "tan", "\\tan({fe})" },
            { "arcsin", "\\arcsin({fe})" },
            { "arccos", "\\arccos({fe})" },
            { "arctan", "\\arctan({fe})" },
            { "arccot", "\\arccot({fe})" },
            { "sinh", "\\sinh({fe})" },
            { "cosh", "\\cosh({fe})" },
            { "tanh", "\\tanh({fe})" },
            { "coth", "\\coth({fe})" },
            { "sec", "\\sec({fe})" },
            { "csc", "\\csc({fe})" },
        };

    protected const string FUNC_PLACE = "{fe}";
    protected const string BRK = "\\\\";

    protected static readonly Dictionary<string, string> CHR_DEFAULT = new Dictionary<string, string>
        {
            { "ACC_VAL", "\\hat{{{0}}}" },
        };

    protected static readonly Dictionary<string, string> POS = new Dictionary<string, string>
        {
            { "top", "\\overline{{{0}}}" }, // Not sure about this
            { "bot", "\\underline{{{0}}}" },
        };

    protected static readonly Dictionary<string, string> POS_DEFAULT = new Dictionary<string, string>
        {
            { "BAR_VAL", "\\overline{{{0}}}" },
        };

    protected const string SUB = "_{{{0}}}";
    protected const string SUP = "^{{{0}}}";

    // Mapping of fraction types to their LaTeX representations
    protected static readonly Dictionary<string, string> F = new Dictionary<string, string>
        {
            { "bar", F_DEFAULT },
            //{ "bar", "\\frac{{{0}}}{{{1}}}" },
            { "skw", "^{{{0}}}/_{{{1}}}" },
            { "noBar", "\\genfrac{{}}{{}}{{0pt}}{{}}{{{0}}}{{{1}}}" },
            { "lin", "{{{0}}}/{{{1}}}" },
        };

    protected const string F_DEFAULT = "\\dfrac{{{0}}}{{{1}}}";
    //protected const string F_DEFAULT = "\\frac{{{0}}}{{{1}}}";

    protected const string D = "\\left{0}{1}\\right{2}";
    protected static readonly Dictionary<string, string> D_DEFAULT = new Dictionary<string, string>
        {
            { "left", "(" },
            { "right", ")" },
            { "null", "." },
        };

    protected const string RAD = "\\sqrt[{0}]{{{1}}}";
    protected const string RAD_DEFAULT = "\\sqrt{{{0}}}";

    protected const string ARR = "\\begin{{array}}{{c}}{0}\\end{{array}}";

    // Mapping of limit functions to their LaTeX representations
    protected static readonly Dictionary<string, string> LIM_FUNC = new Dictionary<string, string>
        {
            { "lim", "\\lim_{{{0}}}" },
            { "max", "\\max_{{{0}}}" },
            { "min", "\\min_{{{0}}}" },
        };

    protected static readonly string[] LIM_TO = { "\\rightarrow", "\\to" };
    protected const string LIM_UPP = "\\overset{{{0}}}{{{1}}}";
    protected const string M = "\\begin{{matrix}}{0}\\end{{matrix}}";
}
