using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public static class StringHelpers
{
    public static string ToStringInvariant(int i)
    {
        return i.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(double d)
    {
        return d.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(decimal d)
    {
        return d.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(float f)
    {
        return f.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(long l)
    {
        return l.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(short s)
    {
        return s.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(ushort us)
    {
        return us.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(uint ui)
    {
        return ui.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(ulong ul)
    {
        return ul.ToString(CultureInfo.InvariantCulture);
    }

    public static string GetLeadingSpaces(string s)
    {
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != ' ')
            {
                return s.Substring(0, i);
            }
        }
        return s;
    }

    public static string GetTrailingSpaces(string s)
    {
        for (int i = s.Length - 1; i >= 0; i--)
        {
            if (s[i] != ' ')
            {
                if (i < s.Length - 1)
                {
                    return s.Substring(i + 1);
                }
                else
                {
                    return string.Empty;
                }
            }
        }
        return s;
    }

    // Convert Wingdings char to Unicode char or emoji.
    // When a standard emoji with similar appearance and meaning exist (even if not identical), it should be preferred
    // over other Unicode chars, as it will be displayed as colored by browsers and is less likely to be missing on the system.
    // However, Visual Studio displays emojis as black and white anyway so it should be checked on
    // https://emojipedia.org, Windows 11 notepad or Visual Studio Code.
    public static string WingdingsToUnicode(char wingdings)
    {
        // https://www.alanwood.net/demos/wingdings.html
        switch (wingdings)
        {
            case '\u0020': return " ";
            case '\u0021': return "✏️";
            case '\u0022': return "✂️";
            case '\u0023': return "✁";
            case '\u0024': return "👓";
            case '\u0025': return "🔔";
            case '\u0026': return "📖";
            case '\u0027': return "🕯️";
            case '\u0028': return "☎️";
            case '\u0029': return "✆";
            case '\u002A': return "✉️";
            case '\u002B': return "🖃";
            case '\u002C': return "📪";
            case '\u002D': return "📫";
            case '\u002E': return "📬";
            case '\u002F': return "📭";

            case '\u0030': return "📁";
            case '\u0031': return "📂";
            case '\u0032': return "📄";
            case '\u0033': return "🗏";
            case '\u0034': return "🗐";
            case '\u0035': return "🗄️";
            case '\u0036': return "⌛";
            case '\u0037': return "⌨️";
            case '\u0038': return "🖰";
            case '\u0039': return "🖲";
            case '\u003A': return "💻";
            case '\u003B': return "🖴";
            case '\u003C': return "💾";
            case '\u003D': return "🖬";
            case '\u003E': return "✇";
            case '\u003F': return "✍";

            case '\u0040': return "🖎";
            case '\u0041': return "✌";
            case '\u0042': return "👌";
            case '\u0043': return "👍";
            case '\u0044': return "👎";
            case '\u0045': return "👈";
            case '\u0046': return "👉";
            case '\u0047': return "☝";
            case '\u0048': return "👇";
            case '\u0049': return "🖐";
            case '\u004A': return "🙂";
            case '\u004B': return "😐";
            case '\u004C': return "🙁";
            case '\u004D': return "💣";
            case '\u004E': return "☠️";
            case '\u004F': return "🏳";

            case '\u0050': return "🏱";
            case '\u0051': return "✈️";
            case '\u0052': return "☀️";
            case '\u0053': return "💧";
            case '\u0054': return "❄️";
            case '\u0055': return "🕆";
            case '\u0056': return "✞";
            case '\u0057': return "🕈";
            case '\u0058': return "✠";
            case '\u0059': return "✡";
            case '\u005A': return "☪";
            case '\u005B': return "☯";
            case '\u005C': return "ॐ";
            case '\u005D': return "☸";
            case '\u005E': return "♈";
            case '\u005F': return "♉";

            case '\u0060': return "♊";
            case '\u0061': return "♋";
            case '\u0062': return "♌";
            case '\u0063': return "♍";
            case '\u0064': return "♎";
            case '\u0065': return "♏";
            case '\u0066': return "♐";
            case '\u0067': return "♑";
            case '\u0068': return "♒";
            case '\u0069': return "♓";
            case '\u006A': return "🙰";
            case '\u006B': return "🙵";
            case '\u006C': return "●";
            case '\u006D': return "🔾";
            case '\u006E': return "■";
            case '\u006F': return "□";

            case '\u0070': return "🞐";
            case '\u0071': return "❑";
            case '\u0072': return "❒";
            case '\u0073': return "⬧";
            case '\u0074': return "⧫";
            case '\u0075': return "◆";
            case '\u0076': return "❖";
            case '\u0077': return "⬥";
            case '\u0078': return "❎";
            case '\u0079': return "⮹";
            case '\u007A': return "⌘";
            case '\u007B': return "🏵";
            case '\u007C': return "🏵️";
            case '\u007D': return "🙶";
            case '\u007E': return "🙷";

            case '\u0080': return "⓪";
            case '\u0081': return "①";
            case '\u0082': return "②";
            case '\u0083': return "③";
            case '\u0084': return "④";
            case '\u0085': return "⑤";
            case '\u0086': return "⑥";
            case '\u0087': return "⑦";
            case '\u0088': return "⑧";
            case '\u0089': return "⑨";
            case '\u008A': return "⑩";
            case '\u008B': return "⓿";
            case '\u008C': return "❶";
            case '\u008D': return "❷";
            case '\u008E': return "❸";
            case '\u008F': return "❹";

            case '\u0090': return "❺";
            case '\u0091': return "❻";
            case '\u0092': return "❼";
            case '\u0093': return "❽";
            case '\u0094': return "❾";
            case '\u0095': return "❿";
            case '\u0096': return "🙢";
            case '\u0097': return "🙠";
            case '\u0098': return "🙡";
            case '\u0099': return "🙣";
            case '\u009A': return "🙞";
            case '\u009B': return "🙜";
            case '\u009C': return "🙝";
            case '\u009D': return "🙟";
            case '\u009E': return "·";
            case '\u009F': return "•";

            case '\u00A0': return "▪";
            case '\u00A1': return "⚪";
            case '\u00A2': return "🞆";
            case '\u00A3': return "🞈";
            case '\u00A4': return "◉";
            case '\u00A5': return "🎯";
            case '\u00A6': return "🔿";
            case '\u00A7': return "▪";
            case '\u00A8': return "◻";
            case '\u00A9': return "🟂";
            case '\u00AA': return "✦";
            case '\u00AB': return "⭐";
            case '\u00AC': return "✶";
            case '\u00AD': return "✴";
            case '\u00AE': return "✹";
            case '\u00AF': return "✵";

            case '\u00B0': return "⯐";
            case '\u00B1': return "⌖";
            case '\u00B2': return "⟡";
            case '\u00B3': return "⌑";
            case '\u00B4': return "⯑";
            case '\u00B5': return "✪";
            case '\u00B6': return "✰";
            case '\u00B7': return "🕐";
            case '\u00B8': return "🕑";
            case '\u00B9': return "🕒";
            case '\u00BA': return "🕓";
            case '\u00BB': return "🕔";
            case '\u00BC': return "🕕";
            case '\u00BD': return "🕖";
            case '\u00BE': return "🕗";
            case '\u00BF': return "🕘";

            case '\u00C0': return "🕙";
            case '\u00C1': return "🕚";
            case '\u00C2': return "🕛";
            case '\u00C3': return "⮰";
            case '\u00C4': return "⮱";
            case '\u00C5': return "⮲";
            case '\u00C6': return "⮳";
            case '\u00C7': return "⮴";
            case '\u00C8': return "⮵";
            case '\u00C9': return "⮶";
            case '\u00CA': return "⮷";
            case '\u00CB': return "🙪";
            case '\u00CC': return "🙫";
            case '\u00CD': return "🙕";
            case '\u00CE': return "🙔";
            case '\u00CF': return "🙗";

            case '\u00D0': return "🙖";
            case '\u00D1': return "🙐";
            case '\u00D2': return "🙑";
            case '\u00D3': return "🙒";
            case '\u00D4': return "🙓";
            case '\u00D5': return "⌫";
            case '\u00D6': return "⌦";
            case '\u00D7': return "⮘";
            case '\u00D8': return "⮚";
            case '\u00D9': return "⮙";
            case '\u00DA': return "⮛";
            case '\u00DB': return "⮈";
            case '\u00DC': return "⮊";
            case '\u00DD': return "⮉";
            case '\u00DE': return "⮋";
            case '\u00DF': return "🡨";

            case '\u00E0': return "🡪";
            case '\u00E1': return "🡩";
            case '\u00E2': return "🡫";
            case '\u00E3': return "🡬";
            case '\u00E4': return "🡭";
            case '\u00E5': return "🡯";
            case '\u00E6': return "🡮";
            case '\u00E7': return "🡸";
            case '\u00E8': return "🡺";
            case '\u00E9': return "🡹";
            case '\u00EA': return "🡻";
            case '\u00EB': return "🡼";
            case '\u00EC': return "🡽";
            case '\u00ED': return "🡿";
            case '\u00EE': return "🡾";
            case '\u00EF': return "⇦";

            case '\u00F0': return "⇨";
            case '\u00F1': return "⇧";
            case '\u00F2': return "⇩";
            case '\u00F3': return "⬄";
            case '\u00F4': return "⇳";
            case '\u00F5': return "⬀";
            case '\u00F6': return "⬁";
            case '\u00F7': return "⬃";
            case '\u00F8': return "⬂";
            case '\u00F9': return "🢬";
            case '\u00FA': return "🢭";
            case '\u00FB': return "❌";
            case '\u00FC': return "✔️";
            case '\u00FD': return "❎";
            case '\u00FE': return "✅";
            case '\u00FF': return "🪟"; // Window emoji (may not be displayed by Visual Studio)
            default: return "";
        }
    }

    public static string Wingdings2ToUnicode(char wingdings)
    {
        // https://www.alanwood.net/demos/wingdings-2.html
        switch (wingdings)
        {
            case '\u0020': return " ";
            case '\u0021': return "🖊️";
            case '\u0022': return "🖋️";
            case '\u0023': return "🖌️";
            case '\u0024': return "🖍️";
            case '\u0025': return "✂️";
            case '\u0026': return "✂️";
            case '\u0027': return "☎️";
            case '\u0028': return "📞";
            case '\u0029': return "🗅";
            case '\u002A': return "🗆";
            case '\u002B': return "🗇";
            case '\u002C': return "🗈";
            case '\u002D': return "🗉";
            case '\u002E': return "🗊";
            case '\u002F': return "🗋";

            case '\u0030': return "🗌";
            case '\u0031': return "🗍";
            case '\u0032': return "📋";
            case '\u0033': return "🗑️";
            case '\u0034': return "🗔";
            case '\u0035': return "🖵";
            case '\u0036': return "🖨️";
            case '\u0037': return "📠";
            case '\u0038': return "💿";
            case '\u0039': return "🖭";
            case '\u003A': return "🖯";
            case '\u003B': return "🖱";
            case '\u003C': return "👍";
            case '\u003D': return "👎";
            case '\u003E': return "🖘";
            case '\u003F': return "🖙";

            case '\u0040': return "🖚";
            case '\u0041': return "🖛";
            case '\u0042': return "👈";
            case '\u0043': return "👉";
            case '\u0044': return "🖜";
            case '\u0045': return "🖝";
            case '\u0046': return "🖞";
            case '\u0047': return "🖟";
            case '\u0048': return "🖠";
            case '\u0049': return "🖡";
            case '\u004A': return "👆";
            case '\u004B': return "👇";
            case '\u004C': return "🖢";
            case '\u004D': return "🖣";
            case '\u004E': return "🖑";
            case '\u004F': return "❌";

            case '\u0050': return "✔️";
            case '\u0051': return "🗵";
            case '\u0052': return "✅";
            case '\u0053': return "❎";
            case '\u0054': return "❎";
            case '\u0055': return "⮾";
            case '\u0056': return "⮿";
            case '\u0057': return "🚫";
            case '\u0058': return "🚫";
            case '\u0059': return "🙱";
            case '\u005A': return "🙴";
            case '\u005B': return "🙲";
            case '\u005C': return "🙳";
            case '\u005D': return "‽";
            case '\u005E': return "🙹";
            case '\u005F': return "🙺";
            
            case '\u0060': return "🙻";
            case '\u0061': return "🙦";
            case '\u0062': return "🙤";
            case '\u0063': return "🙥";
            case '\u0064': return "🙧";
            case '\u0065': return "🙚";
            case '\u0066': return "🙘";
            case '\u0067': return "🙙";
            case '\u0068': return "🙛";
            case '\u0069': return "⓪";
            case '\u006A': return "①";
            case '\u006B': return "②";
            case '\u006C': return "③";
            case '\u006D': return "④";
            case '\u006E': return "⑤";
            case '\u006F': return "⑥";

            case '\u0070': return "⑦";
            case '\u0071': return "⑧";
            case '\u0072': return "⑨";
            case '\u0073': return "⑩";
            case '\u0074': return "⓿";
            case '\u0075': return "❶";
            case '\u0076': return "❷";
            case '\u0077': return "❸";
            case '\u0078': return "❹";
            case '\u0079': return "❺";
            case '\u007A': return "❻";
            case '\u007B': return "❼";
            case '\u007C': return "❽";
            case '\u007D': return "❾";
            case '\u007E': return "❿";

            case '\u0080': return "☉";
            case '\u0081': return "🌕";
            case '\u0082': return "☽";
            case '\u0083': return "☾";
            case '\u0084': return "⸿";
            case '\u0085': return "✝";
            case '\u0086': return "🕇";
            case '\u0087': return "🕜";
            case '\u0088': return "🕝";
            case '\u0089': return "🕞";
            case '\u008A': return "🕟";
            case '\u008B': return "🕠";
            case '\u008C': return "🕡";
            case '\u008D': return "🕢";
            case '\u008E': return "🕣";
            case '\u008F': return "🕤";

            case '\u0090': return "🕥";
            case '\u0091': return "🕦";
            case '\u0092': return "🕧";
            case '\u0093': return "🙨";
            case '\u0094': return "🙩";
            case '\u0095': return "•";
            case '\u0096': return "●";
            case '\u0097': return "⚫";
            case '\u0098': return "⬤";
            case '\u0099': return "🞅";
            case '\u009A': return "🞆";
            case '\u009B': return "🞇";
            case '\u009C': return "🞈";
            case '\u009D': return "🞊";
            case '\u009E': return "⦿";
            case '\u009F': return "◾";

            case '\u00A0': return "■";
            case '\u00A1': return "◼";
            case '\u00A2': return "⬛";
            case '\u00A3': return "⬜";
            case '\u00A4': return "🞑";
            case '\u00A5': return "🞒";
            case '\u00A6': return "🞓";
            case '\u00A7': return "🞔";
            case '\u00A8': return "▣";
            case '\u00A9': return "🞕";
            case '\u00AA': return "🞖";
            case '\u00AB': return "🞗";
            case '\u00AC': return "⬩";
            case '\u00AD': return "⬥";
            case '\u00AE': return "◆";
            case '\u00AF': return "◇";

            case '\u00B0': return "🞚";
            case '\u00B1': return "◈";
            case '\u00B2': return "🞛";
            case '\u00B3': return "🞜";
            case '\u00B4': return "🞝";
            case '\u00B5': return "⬪";
            case '\u00B6': return "⬧";
            case '\u00B7': return "⧫";
            case '\u00B8': return "◊";
            case '\u00B9': return "🞠";
            case '\u00BA': return "◖";
            case '\u00BB': return "◗";
            case '\u00BC': return "⯊";
            case '\u00BD': return "⯋";
            case '\u00BE': return "◼";
            case '\u00BF': return "⬥";

            case '\u00C0': return "⬟";
            case '\u00C1': return "⯂";
            case '\u00C2': return "⬣";
            case '\u00C3': return "⬢";
            case '\u00C4': return "⯃";
            case '\u00C5': return "⯄";
            case '\u00C6': return "🞡";
            case '\u00C7': return "🞢";
            case '\u00C8': return "🞣";
            case '\u00C9': return "🞤";
            case '\u00CA': return "🞥";
            case '\u00CB': return "🞦";
            case '\u00CC': return "🞧";
            case '\u00CD': return "🞨";
            case '\u00CE': return "🞩";
            case '\u00CF': return "🞪";

            case '\u00D0': return "🞫";
            case '\u00D1': return "🞬";
            case '\u00D2': return "🞭";
            case '\u00D3': return "🞮";
            case '\u00D4': return "🞯";
            case '\u00D5': return "🞰";
            case '\u00D6': return "🞱";
            case '\u00D7': return "🞲";
            case '\u00D8': return "🞳";
            case '\u00D9': return "🞴";
            case '\u00DA': return "🞵";
            case '\u00DB': return "🞶";
            case '\u00DC': return "🞷";
            case '\u00DD': return "🞸";
            case '\u00DE': return "🞹";
            case '\u00DF': return "🞺";

            case '\u00E0': return "🞻";
            case '\u00E1': return "🞼";
            case '\u00E2': return "🞽";
            case '\u00E3': return "🞾";
            case '\u00E4': return "🞿";
            case '\u00E5': return "🟀";
            case '\u00E6': return "🟂";
            case '\u00E7': return "🟄";
            case '\u00E8': return "✦";
            case '\u00E9': return "🟉";
            case '\u00EA': return "⭐";
            case '\u00EB': return "✶";
            case '\u00EC': return "🟋";
            case '\u00ED': return "✷";
            case '\u00EE': return "🟏";
            case '\u00EF': return "🟒";

            case '\u00F0': return "✹";
            case '\u00F1': return "🟃";
            case '\u00F2': return "🟇";
            case '\u00F3': return "✯";
            case '\u00F4': return "🟍";
            case '\u00F5': return "🟔";
            case '\u00F6': return "⯌";
            case '\u00F7': return "⯍";
            case '\u00F8': return "※";
            case '\u00F9': return "⁂";
            default: return "";
        }
    }

    public static string Wingdings3ToUnicode(char wingdings)
    {
        // https://www.alanwood.net/demos/wingdings-3.html
        switch (wingdings)
        {
            case '\u0020': return " ";
            case '\u0021': return "⭠";
            case '\u0022': return "⭢";
            case '\u0023': return "⭡";
            case '\u0024': return "⭣";
            case '\u0025': return "⭦";
            case '\u0026': return "⭧";
            case '\u0027': return "⭩";
            case '\u0028': return "⭨";
            case '\u0029': return "⭰";
            case '\u002A': return "⭲";
            case '\u002B': return "⭱";
            case '\u002C': return "⭳";
            case '\u002D': return "⭶";
            case '\u002E': return "⭸";
            case '\u002F': return "⭻";

            case '\u0030': return "⭽";
            case '\u0031': return "⭤";
            case '\u0032': return "⭥";
            case '\u0033': return "⭪";
            case '\u0034': return "⭬";
            case '\u0035': return "⭫";
            case '\u0036': return "⭭";
            case '\u0037': return "⭍";
            case '\u0038': return "⮠";
            case '\u0039': return "⮡";
            case '\u003A': return "⮢";
            case '\u003B': return "⮣";
            case '\u003C': return "⮤";
            case '\u003D': return "⮥";
            case '\u003E': return "⮦";
            case '\u003F': return "⮧";
            
            case '\u0040': return "⮐";
            case '\u0041': return "⮑";
            case '\u0042': return "⮒";
            case '\u0043': return "⮓";
            case '\u0044': return "⮀";
            case '\u0045': return "⮃";
            case '\u0046': return "⭾";
            case '\u0047': return "⭿";
            case '\u0048': return "⮄";
            case '\u0049': return "⮆";
            case '\u004A': return "⮅";
            case '\u004B': return "⮇";
            case '\u004C': return "⮏";
            case '\u004D': return "⮍";
            case '\u004E': return "⮎";
            case '\u004F': return "⮌";

            case '\u0050': return "⭮";
            case '\u0051': return "⭯";
            case '\u0052': return "⎋";
            case '\u0053': return "⌤";
            case '\u0054': return "⌃";
            case '\u0055': return "⌥";
            case '\u0056': return "⎵";
            case '\u0057': return "⍽";
            case '\u0058': return "⇪";
            case '\u0059': return "⮸";
            case '\u005A': return "🢠";
            case '\u005B': return "🢡";
            case '\u005C': return "🢢";
            case '\u005D': return "🢣";
            case '\u005E': return "🢤";
            case '\u005F': return "🢥";

            case '\u0060': return "🢦";
            case '\u0061': return "🢧";
            case '\u0062': return "🢨";
            case '\u0063': return "🢩";
            case '\u0064': return "🢪";
            case '\u0065': return "🢫";
            case '\u0066': return "←";
            case '\u0067': return "→";
            case '\u0068': return "↑";
            case '\u0069': return "↓";
            case '\u006A': return "↖";
            case '\u006B': return "↗";
            case '\u006C': return "↙";
            case '\u006D': return "↘";
            case '\u006E': return "🡘";
            case '\u006F': return "🡙";

            case '\u0070': return "▲";
            case '\u0071': return "▼";
            case '\u0072': return "△";
            case '\u0073': return "▽";
            case '\u0074': return "◄";
            case '\u0075': return "►";
            case '\u0076': return "◁";
            case '\u0077': return "▷";
            case '\u0078': return "◣";
            case '\u0079': return "◢";
            case '\u007A': return "◤";
            case '\u007B': return "◥";
            case '\u007C': return "🞀";
            case '\u007D': return "🞂";
            case '\u007E': return "🞁";
            
            case '\u0080': return "🞃";
            case '\u0081': return "▲";
            case '\u0082': return "▼";
            case '\u0083': return "◀";
            case '\u0084': return "▶";
            case '\u0085': return "⮜";
            case '\u0086': return "⮞";
            case '\u0087': return "⮝";
            case '\u0088': return "⮟";
            case '\u0089': return "🠐";
            case '\u008A': return "🠒";
            case '\u008B': return "🠑";
            case '\u008C': return "🠓";
            case '\u008D': return "🠔";
            case '\u008E': return "🠖";
            case '\u008F': return "🠕";
            
            case '\u0090': return "🠗";
            case '\u0091': return "🠘";
            case '\u0092': return "🠚";
            case '\u0093': return "🠙";
            case '\u0094': return "🠛";
            case '\u0095': return "🠜";
            case '\u0096': return "🠞";
            case '\u0097': return "🠝";
            case '\u0098': return "🠟";
            case '\u0099': return "🠀";
            case '\u009A': return "🠂";
            case '\u009B': return "🠁";
            case '\u009C': return "🠃";
            case '\u009D': return "🠄";
            case '\u009E': return "🠆";
            case '\u009F': return "🠅";

            case '\u00a0': return "🠇";
            case '\u00a1': return "🠈";
            case '\u00a2': return "🠊";
            case '\u00a3': return "🠉";
            case '\u00a4': return "🠋";
            case '\u00a5': return "🠠";
            case '\u00a6': return "🠢";
            case '\u00a7': return "🠤";
            case '\u00a8': return "🠦";
            case '\u00a9': return "🠨";
            case '\u00aa': return "🠪";
            case '\u00ab': return "🠬";
            case '\u00ac': return "🢜";
            case '\u00ad': return "🢝";
            case '\u00ae': return "🢞";
            case '\u00af': return "🢟";

            case '\u00b0': return "🠮";
            case '\u00b1': return "🠰";
            case '\u00b2': return "🠲";
            case '\u00b3': return "🠴";
            case '\u00b4': return "🠶";
            case '\u00b5': return "🠸";
            case '\u00b6': return "🠺";
            case '\u00b7': return "🠹";
            case '\u00b8': return "🠻";
            case '\u00b9': return "🢘";
            case '\u00ba': return "🢚";
            case '\u00bb': return "🢙";
            case '\u00bc': return "🢛";
            case '\u00bd': return "🠼";
            case '\u00be': return "🠾";
            case '\u00bf': return "🠽";

            case '\u00c0': return "🠿";
            case '\u00c1': return "🡀";
            case '\u00c2': return "🡂";
            case '\u00c3': return "🡁";
            case '\u00c4': return "🡃";
            case '\u00c5': return "🡄";
            case '\u00c6': return "🡆";
            case '\u00c7': return "🡅";
            case '\u00c8': return "🡇";
            case '\u00c9': return "⮨";
            case '\u00ca': return "⮩";
            case '\u00cb': return "⮪";
            case '\u00cc': return "⮫";
            case '\u00cd': return "⮬";
            case '\u00ce': return "⮭";
            case '\u00cf': return "⮮";
            
            case '\u00d0': return "⮯";
            case '\u00d1': return "🡠";
            case '\u00d2': return "🡢";
            case '\u00d3': return "🡡";
            case '\u00d4': return "🡣";
            case '\u00d5': return "🡤";
            case '\u00d6': return "🡥";
            case '\u00d7': return "🡧";
            case '\u00d8': return "🡦";
            case '\u00d9': return "🡰";
            case '\u00da': return "🡲";
            case '\u00db': return "🡱";
            case '\u00dc': return "🡳";
            case '\u00dd': return "🡴";
            case '\u00de': return "🡵";
            case '\u00df': return "🡷";
            
            case '\u00e0': return "🡶";
            case '\u00e1': return "🢀";
            case '\u00e2': return "🢂";
            case '\u00e3': return "🢁";
            case '\u00e4': return "🢃";
            case '\u00e5': return "🢄";
            case '\u00e6': return "🢅";
            case '\u00e7': return "🢇";
            case '\u00e8': return "🢆";
            case '\u00e9': return "🢐";
            case '\u00ea': return "🢒";
            case '\u00eb': return "🢑";
            case '\u00ec': return "🢓";
            case '\u00ed': return "🢔";
            case '\u00ee': return "🢖";
            case '\u00ef': return "🢕";
            
            case '\u00f0': return "🢗";
            default: return "";
        }
    }

    public static string WebdingsToUnicode(char wingdings)
    {
        switch (wingdings)
        {
            case '\u0020': return " ";
            case '\u0021': return "🕷";
            case '\u0022': return "🕸";
            case '\u0023': return "🕲"; 
            case '\u0024': return "🕶";
            case '\u0025': return "🏆";
            case '\u0026': return "🏅";
            case '\u0027': return "🖇";
            case '\u0028': return "🗨";
            case '\u0029': return "💬";
            case '\u002A': return "🗰"; 
            case '\u002B': return "🗱"; 
            case '\u002C': return "🌶";
            case '\u002D': return "🎗";
            case '\u002E': return "🙾"; 
            case '\u002F': return "🙼"; 
            
            case '\u0030': return "🗕"; 
            case '\u0031': return "🗖"; 
            case '\u0032': return "🗗"; 
            case '\u0033': return "◀";
            case '\u0034': return "▶";
            case '\u0035': return "▲";
            case '\u0036': return "▼";
            case '\u0037': return "⏪";
            case '\u0038': return "⏩";
            case '\u0039': return "⏮";
            case '\u003A': return "⏭";
            case '\u003B': return "⏸";
            case '\u003C': return "⏹";
            case '\u003D': return "⏺";
            case '\u003E': return "🗚"; 
            case '\u003F': return "🗳";

            case '\u0040': return "🛠";
            case '\u0041': return "🏗";
            case '\u0042': return "🏘";
            case '\u0043': return "🏙";
            case '\u0044': return "🏚";
            case '\u0045': return "🏜";
            case '\u0046': return "🏭";
            case '\u0047': return "🏛";
            case '\u0048': return "🏠";
            case '\u0049': return "🏖";
            case '\u004A': return "🏝";
            case '\u004B': return "🛣";
            case '\u004C': return "🔍";
            case '\u004D': return "🏔";
            case '\u004E': return "👁";
            case '\u004F': return "👂";
            
            case '\u0050': return "🏞";
            case '\u0051': return "🏕";
            case '\u0052': return "🛤";
            case '\u0053': return "🏟";
            case '\u0054': return "🛳";
            case '\u0055': return "🔊";
            case '\u0056': return "📢";
            case '\u0057': return "🕨"; 
            case '\u0058': return "🔈";
            case '\u0059': return "🎔"; 
            case '\u005A': return "💐";
            case '\u005B': return "🗬"; 
            case '\u005C': return "🙽"; 
            case '\u005D': return "💭";
            case '\u005E': return "🗪"; 
            case '\u005F': return "🗫"; 

            case '\u0060': return "🔄";
            case '\u0061': return "✔";
            case '\u0062': return "🚲";
            case '\u0063': return "□";
            case '\u0064': return "🛡";
            case '\u0065': return "📦";
            case '\u0066': return "🚒";
            case '\u0067': return "⬛";
            case '\u0068': return "🚑";
            case '\u0069': return "ℹ";
            case '\u006A': return "🛩";
            case '\u006B': return "🛰";
            case '\u006C': return "🟈"; 
            case '\u006D': return "🕴";
            case '\u006E': return "⚫";
            case '\u006F': return "🛥";

            case '\u0070': return "🚔";
            case '\u0071': return "🔃";
            case '\u0072': return "❌";
            case '\u0073': return "❓";
            case '\u0074': return "🚆";
            case '\u0075': return "🚇";
            case '\u0076': return "🚍";
            case '\u0077': return "⛳";
            case '\u0078': return "🚫";
            case '\u0079': return "⛔";
            case '\u007A': return "🚭";
            case '\u007B': return "🗮"; 
            case '\u007C': return "|";
            case '\u007D': return "🗯";
            case '\u007E': return "⚡";

            case '\u0080': return "🚹";
            case '\u0081': return "🚺";
            case '\u0082': return "🛉"; 
            case '\u0083': return "🛊"; 
            case '\u0084': return "🚼";
            case '\u0085': return "👽";
            case '\u0086': return "🏋";
            case '\u0087': return "⛷";
            case '\u0088': return "🏂";
            case '\u0089': return "🏌";
            case '\u008A': return "🏊";
            case '\u008B': return "🏄";
            case '\u008C': return "🏍";
            case '\u008D': return "🏎";
            case '\u008E': return "🚘";
            case '\u008F': return "📈";

            case '\u0090': return "🛢";
            case '\u0091': return "💰";
            case '\u0092': return "🏷";
            case '\u0093': return "💳";
            case '\u0094': return "👪";
            case '\u0095': return "🗡";
            case '\u0096': return "💋";
            case '\u0097': return "🗣";
            case '\u0098': return "⭐";
            case '\u0099': return "🖄"; 
            case '\u009A': return "📨";
            case '\u009B': return "✉";
            case '\u009C': return "🖆"; 
            case '\u009D': return "📄";
            case '\u009E': return "🖺"; 
            case '\u009F': return "🖻"; 

            case '\u00A0': return "🕵";
            case '\u00A1': return "🕰";
            case '\u00A2': return "🖼";
            case '\u00A3': return "🖼";
            case '\u00A4': return "📋";
            case '\u00A5': return "🗒";
            case '\u00A6': return "🗓";
            case '\u00A7': return "📖";
            case '\u00A8': return "📚";
            case '\u00A9': return "🗞";
            case '\u00AA': return "📰";
            case '\u00AB': return "🗃";
            case '\u00AC': return "🗂";
            case '\u00AD': return "🖼";
            case '\u00AE': return "🎭";
            case '\u00AF': return "🎵";

            case '\u00B0': return "🎹";
            case '\u00B1': return "🎙";
            case '\u00B2': return "🎧";
            case '\u00B3': return "💿";
            case '\u00B4': return "🎞";
            case '\u00B5': return "📷";
            case '\u00B6': return "🎟";
            case '\u00B7': return "🎬";
            case '\u00B8': return "📽";
            case '\u00B9': return "📹";
            case '\u00BA': return "📾"; 
            case '\u00BB': return "📻";
            case '\u00BC': return "🎚";
            case '\u00BD': return "🎛";
            case '\u00BE': return "📺";
            case '\u00BF': return "💻";

            case '\u00C0': return "🖥";
            case '\u00C1': return "🖦";
            case '\u00C2': return "🖧"; 
            case '\u00C3': return "🕹";
            case '\u00C4': return "🎮";
            case '\u00C5': return "📞";
            case '\u00C6': return "🕼"; 
            case '\u00C7': return "📟";
            case '\u00C8': return "📱";
            case '\u00C9': return "☎";
            case '\u00CA': return "🖨";
            case '\u00CB': return "🖩"; 
            case '\u00CC': return "📁";
            case '\u00CD': return "💾";
            case '\u00CE': return "🗜";
            case '\u00CF': return "🔒";

            case '\u00D0': return "🔓";
            case '\u00D1': return "🗝";
            case '\u00D2': return "📥";
            case '\u00D3': return "📤";
            case '\u00D4': return "🕳";
            //case '\u00D5': return "🌣";
            case '\u00D5': return "☀";
            case '\u00D6': return "🌤";
            case '\u00D7': return "🌥";
            case '\u00D8': return "🌦";
            case '\u00D9': return "☁";
            case '\u00DA': return "🌨";
            case '\u00DB': return "🌧";
            case '\u00DC': return "🌩";
            case '\u00DD': return "🌪";
            case '\u00DE': return "🌬";
            case '\u00DF': return "🌫";

            case '\u00E0': return "🌜";
            case '\u00E1': return "🌡";
            case '\u00E2': return "🛋";
            case '\u00E3': return "🛏";
            case '\u00E4': return "🍽";
            case '\u00E5': return "🍸";
            case '\u00E6': return "🛎";
            case '\u00E7': return "🛍";
            case '\u00E8': return "🅿️";
            case '\u00E9': return "♿";
            case '\u00EA': return "🔺";
            case '\u00EB': return "📌";
            case '\u00EC': return "🎓";
            case '\u00ED': return "🗤"; 
            case '\u00EE': return "🗥"; 
            case '\u00EF': return "🗦"; 

            case '\u00F0': return "🗧"; 
            case '\u00F1': return "✈";
            case '\u00F2': return "🐿";
            case '\u00F3': return "🐦";
            case '\u00F4': return "🐟";
            case '\u00F5': return "🐕";
            case '\u00F6': return "🐈";
            case '\u00F7': return "🚀";
            case '\u00F8': return "🚀";
            case '\u00F9': return "🚀";
            case '\u00FA': return "🚀";
            case '\u00FB': return "🗺";
            case '\u00FC': return "🌍";
            case '\u00FD': return "🌏";
            case '\u00FE': return "🌎";
            case '\u00FF': return "🕊";
            default: return "";
        }
    }
}
