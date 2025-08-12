using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal static class RtfBorderMapper
{
    internal static string GetBorderType(BorderValues borderValue)
    {
        if (borderValue == BorderValues.Single)
            return @"\brdrs";
        else if (borderValue == BorderValues.Thick)
            return @"\brdrth";
        else if (borderValue == BorderValues.Double)
            return @"\brdrdb";
        else if (borderValue == BorderValues.Dotted)
            return @"\brdrdot";
        else if (borderValue == BorderValues.Dashed)
            return @"\brdrdash";
        else if (borderValue == BorderValues.Triple)
            return @"\brdrtriple";
        else if (borderValue == BorderValues.Wave)
            return @"\brdrwavy";
        else if (borderValue == BorderValues.DoubleWave)
            return @"\brdrwavydb";
        else if (borderValue == BorderValues.Outset)
            return @"\brdroutset";
        else if (borderValue == BorderValues.Inset)
            return @"\brdrinset";
        else if (borderValue == BorderValues.DashSmallGap)
            return @"\brdrdashsm";
        else if (borderValue == BorderValues.DotDash)
            return @"\brdrdashd";
        else if (borderValue == BorderValues.DotDotDash)
            return @"\brdrdashdd";
        else if (borderValue == BorderValues.ThickThinSmallGap)
            return @"\brdrtnthsg";
        else if (borderValue == BorderValues.ThickThinMediumGap)
            return @"\brdrtnthmg";
        else if (borderValue == BorderValues.ThickThinLargeGap)
            return @"\brdrtnthlg";
        else if (borderValue == BorderValues.ThinThickSmallGap)
            return @"\brdrthtnsg";
        else if (borderValue == BorderValues.ThinThickMediumGap)
            return @"\brdrthtnmg";
        else if (borderValue == BorderValues.ThinThickLargeGap)
            return @"\brdrthtnlg";
        else if (borderValue == BorderValues.ThinThickThinSmallGap)
            return @"\brdrtnthtnsg";
        else if (borderValue == BorderValues.ThinThickThinMediumGap)
            return @"\brdrtnthtnmg";
        else if (borderValue == BorderValues.ThinThickThinLargeGap)
            return @"\brdrtnthtnlg";
        else if (borderValue == BorderValues.ThreeDEmboss)
            return @"\brdremboss";
        else if (borderValue == BorderValues.ThreeDEngrave)
            return @"\brdrengrave";
        else if (borderValue == BorderValues.DashDotStroked)
            return @"\brdrdashdotstr";
        else if (borderValue == BorderValues.Nil)
            // return @"\brdrnil"; 
            // In DOCX nil means no border.
            // "brdrnil" in RTF is interpreted differently by MS Word (probably like a default border) causing some issues,
            // although in LibreOffice and OnlyOffice is interpreted as no border.
            // We should use brdrnone instead.
            return @"\brdrnone";
        else if (borderValue == BorderValues.None)
            return @"\brdrnone";

        else if (borderValue == BorderValues.Apples)
            return @"\brdrart1";
        else if (borderValue == BorderValues.ArchedScallops)
            return @"\brdrart2";
        else if (borderValue == BorderValues.BabyPacifier)
            return @"\brdrart3";
        else if (borderValue == BorderValues.BabyRattle)
            return @"\brdrart4";
        else if (borderValue == BorderValues.Balloons3Colors)
            return @"\brdrart5";
        else if (borderValue == BorderValues.BalloonsHotAir)
            return @"\brdrart6";
        else if (borderValue == BorderValues.BasicBlackDashes)
            return @"\brdrart7";
        else if (borderValue == BorderValues.BasicBlackDots)
            return @"\brdrart8";
        else if (borderValue == BorderValues.BasicBlackSquares)
            return @"\brdrart9";
        else if (borderValue == BorderValues.BasicThinLines)
            return @"\brdrart10";
        else if (borderValue == BorderValues.BasicWhiteDashes)
            return @"\brdrart11";
        else if (borderValue == BorderValues.BasicWhiteDots)
            return @"\brdrart12";
        else if (borderValue == BorderValues.BasicWhiteSquares)
            return @"\brdrart13";
        else if (borderValue == BorderValues.BasicWideInline)
            return @"\brdrart14";
        else if (borderValue == BorderValues.BasicWideMidline)
            return @"\brdrart15";
        else if (borderValue == BorderValues.BasicWideOutline)
            return @"\brdrart16";
        else if (borderValue == BorderValues.Bats)
            return @"\brdrart17";
        else if (borderValue == BorderValues.Birds)
            return @"\brdrart18";
        else if (borderValue == BorderValues.BirdsFlight)
            return @"\brdrart19";
        else if (borderValue == BorderValues.Cabins)
            return @"\brdrart20";
        else if (borderValue == BorderValues.CakeSlice)
            return @"\brdrart21";
        else if (borderValue == BorderValues.CandyCorn)
            return @"\brdrart22";
        else if (borderValue == BorderValues.CelticKnotwork)
            return @"\brdrart23";
        else if (borderValue == BorderValues.CertificateBanner)
            return @"\brdrart24";
        else if (borderValue == BorderValues.ChainLink)
            return @"\brdrart25";
        else if (borderValue == BorderValues.ChampagneBottle)
            return @"\brdrart26";
        else if (borderValue == BorderValues.CheckedBarBlack)
            return @"\brdrart27";
        else if (borderValue == BorderValues.CheckedBarColor)
            return @"\brdrart28";
        else if (borderValue == BorderValues.Checkered)
            return @"\brdrart29";
        else if (borderValue == BorderValues.ChristmasTree)
            return @"\brdrart30";
        else if (borderValue == BorderValues.CirclesLines)
            return @"\brdrart31";
        else if (borderValue == BorderValues.CirclesRectangles)
            return @"\brdrart32";
        else if (borderValue == BorderValues.ClassicalWave)
            return @"\brdrart33";
        else if (borderValue == BorderValues.Clocks)
            return @"\brdrart34";
        else if (borderValue == BorderValues.Compass)
            return @"\brdrart35";
        else if (borderValue == BorderValues.Confetti)
            return @"\brdrart36";
        else if (borderValue == BorderValues.ConfettiGrays)
            return @"\brdrart37";
        else if (borderValue == BorderValues.ConfettiOutline)
            return @"\brdrart38";
        else if (borderValue == BorderValues.ConfettiStreamers)
            return @"\brdrart39";
        else if (borderValue == BorderValues.ConfettiWhite)
            return @"\brdrart40";
        else if (borderValue == BorderValues.CornerTriangles)
            return @"\brdrart41";
        else if (borderValue == BorderValues.CouponCutoutDashes)
            return @"\brdrart42";
        else if (borderValue == BorderValues.CouponCutoutDots)
            return @"\brdrart43";
        else if (borderValue == BorderValues.CrazyMaze)
            return @"\brdrart44";
        else if (borderValue == BorderValues.CreaturesButterfly)
            return @"\brdrart45";
        else if (borderValue == BorderValues.CreaturesFish)
            return @"\brdrart46";
        else if (borderValue == BorderValues.CreaturesInsects)
            return @"\brdrart47";
        else if (borderValue == BorderValues.CreaturesLadyBug)
            return @"\brdrart48";
        else if (borderValue == BorderValues.CrossStitch)
            return @"\brdrart49";
        else if (borderValue == BorderValues.Cup)
            return @"\brdrart50";
        else if (borderValue == BorderValues.DecoArch)
            return @"\brdrart51";
        else if (borderValue == BorderValues.DecoArchColor)
            return @"\brdrart52";
        else if (borderValue == BorderValues.DecoBlocks)
            return @"\brdrart53";
        else if (borderValue == BorderValues.DiamondsGray)
            return @"\brdrart54";
        else if (borderValue == BorderValues.DoubleD)
            return @"\brdrart55";
        else if (borderValue == BorderValues.DoubleDiamonds)
            return @"\brdrart56";
        else if (borderValue == BorderValues.Earth1)
            return @"\brdrart57";
        else if (borderValue == BorderValues.Earth2)
            return @"\brdrart58";
        else if (borderValue == BorderValues.EclipsingSquares1)
            return @"\brdrart59";
        else if (borderValue == BorderValues.EclipsingSquares2)
            return @"\brdrart60";
        else if (borderValue == BorderValues.EggsBlack)
            return @"\brdrart61";
        else if (borderValue == BorderValues.Fans)
            return @"\brdrart62";
        else if (borderValue == BorderValues.Film)
            return @"\brdrart63";
        else if (borderValue == BorderValues.Firecrackers)
            return @"\brdrart64";
        else if (borderValue == BorderValues.FlowersBlockPrint)
            return @"\brdrart65";
        else if (borderValue == BorderValues.FlowersDaisies)
            return @"\brdrart66";
        else if (borderValue == BorderValues.FlowersModern1)
            return @"\brdrart67";
        else if (borderValue == BorderValues.FlowersModern2)
            return @"\brdrart68";
        else if (borderValue == BorderValues.FlowersPansy)
            return @"\brdrart69";
        else if (borderValue == BorderValues.FlowersRedRose)
            return @"\brdrart70";
        else if (borderValue == BorderValues.FlowersRoses)
            return @"\brdrart71";
        else if (borderValue == BorderValues.FlowersTeacup)
            return @"\brdrart72";
        else if (borderValue == BorderValues.FlowersTiny)
            return @"\brdrart73";
        else if (borderValue == BorderValues.Gems)
            return @"\brdrart74";
        else if (borderValue == BorderValues.GingerbreadMan)
            return @"\brdrart75";
        else if (borderValue == BorderValues.Gradient)
            return @"\brdrart76";
        else if (borderValue == BorderValues.Handmade1)
            return @"\brdrart77";
        else if (borderValue == BorderValues.Handmade2)
            return @"\brdrart78";
        else if (borderValue == BorderValues.HeartBalloon)
            return @"\brdrart79";
        else if (borderValue == BorderValues.HeartGray)
            return @"\brdrart80";
        else if (borderValue == BorderValues.Hearts)
            return @"\brdrart81";
        else if (borderValue == BorderValues.HeebieJeebies)
            return @"\brdrart82";
        else if (borderValue == BorderValues.Holly)
            return @"\brdrart83";
        else if (borderValue == BorderValues.HouseFunky)
            return @"\brdrart84";
        else if (borderValue == BorderValues.Hypnotic)
            return @"\brdrart85";
        else if (borderValue == BorderValues.IceCreamCones)
            return @"\brdrart86";
        else if (borderValue == BorderValues.LightBulb)
            return @"\brdrart87";
        else if (borderValue == BorderValues.Lightning1)
            return @"\brdrart88";
        else if (borderValue == BorderValues.Lightning2)
            return @"\brdrart89";
        else if (borderValue == BorderValues.MapleLeaf)
            return @"\brdrart91";
        else if (borderValue == BorderValues.MapleMuffins)
            return @"\brdrart92";
        else if (borderValue == BorderValues.MapPins)
            return @"\brdrart90";
        else if (borderValue == BorderValues.Marquee)
            return @"\brdrart93";
        else if (borderValue == BorderValues.MarqueeToothed)
            return @"\brdrart94";
        else if (borderValue == BorderValues.Moons)
            return @"\brdrart95";
        else if (borderValue == BorderValues.Mosaic)
            return @"\brdrart96";
        else if (borderValue == BorderValues.MusicNotes)
            return @"\brdrart97";
        else if (borderValue == BorderValues.Northwest)
            return @"\brdrart98";
        else if (borderValue == BorderValues.Ovals)
            return @"\brdrart99";
        else if (borderValue == BorderValues.Packages)
            return @"\brdrart100";
        else if (borderValue == BorderValues.PalmsBlack)
            return @"\brdrart101";
        else if (borderValue == BorderValues.PalmsColor)
            return @"\brdrart102";
        else if (borderValue == BorderValues.PaperClips)
            return @"\brdrart103";
        else if (borderValue == BorderValues.Papyrus)
            return @"\brdrart104";
        else if (borderValue == BorderValues.PartyFavor)
            return @"\brdrart105";
        else if (borderValue == BorderValues.PartyGlass)
            return @"\brdrart106";
        else if (borderValue == BorderValues.Pencils)
            return @"\brdrart107";
        else if (borderValue == BorderValues.People)
            return @"\brdrart108";
        else if (borderValue == BorderValues.PeopleHats)
            return @"\brdrart110";
        else if (borderValue == BorderValues.PeopleWaving)
            return @"\brdrart109";
        else if (borderValue == BorderValues.Poinsettias)
            return @"\brdrart111";
        else if (borderValue == BorderValues.PostageStamp)
            return @"\brdrart112";
        else if (borderValue == BorderValues.Pumpkin1)
            return @"\brdrart113";
        else if (borderValue == BorderValues.PushPinNote1)
            return @"\brdrart115";
        else if (borderValue == BorderValues.PushPinNote2)
            return @"\brdrart114";
        else if (borderValue == BorderValues.Pyramids)
            return @"\brdrart116";
        else if (borderValue == BorderValues.PyramidsAbove)
            return @"\brdrart117";
        else if (borderValue == BorderValues.Quadrants)
            return @"\brdrart118";
        else if (borderValue == BorderValues.Rings)
            return @"\brdrart119";
        else if (borderValue == BorderValues.Safari)
            return @"\brdrart120";
        else if (borderValue == BorderValues.Sawtooth)
            return @"\brdrart121";
        else if (borderValue == BorderValues.SawtoothGray)
            return @"\brdrart122";
        else if (borderValue == BorderValues.ScaredCat)
            return @"\brdrart123";
        else if (borderValue == BorderValues.Seattle)
            return @"\brdrart124";
        else if (borderValue == BorderValues.ShadowedSquares)
            return @"\brdrart125";
        else if (borderValue == BorderValues.SharksTeeth)
            return @"\brdrart126";
        else if (borderValue == BorderValues.ShorebirdTracks)
            return @"\brdrart127";
        else if (borderValue == BorderValues.Skyrocket)
            return @"\brdrart128";
        else if (borderValue == BorderValues.SnowflakeFancy)
            return @"\brdrart129";
        else if (borderValue == BorderValues.Snowflakes)
            return @"\brdrart130";
        else if (borderValue == BorderValues.Sombrero)
            return @"\brdrart131";
        else if (borderValue == BorderValues.Southwest)
            return @"\brdrart132";
        else if (borderValue == BorderValues.Stars)
            return @"\brdrart133";
        else if (borderValue == BorderValues.Stars3d)
            return @"\brdrart135";
        else if (borderValue == BorderValues.StarsBlack)
            return @"\brdrart136";
        else if (borderValue == BorderValues.StarsShadowed)
            return @"\brdrart137";
        else if (borderValue == BorderValues.StarsTop)
            return @"\brdrart134";
        else if (borderValue == BorderValues.Sun)
            return @"\brdrart138";
        else if (borderValue == BorderValues.Swirligig)
            return @"\brdrart139";
        else if (borderValue == BorderValues.TornPaper)
            return @"\brdrart140";
        else if (borderValue == BorderValues.TornPaperBlack)
            return @"\brdrart141";
        else if (borderValue == BorderValues.Trees)
            return @"\brdrart142";
        else if (borderValue == BorderValues.TriangleParty)
            return @"\brdrart143";
        else if (borderValue == BorderValues.Triangles)
            return @"\brdrart144";
        else if (borderValue == BorderValues.Tribal1)
            return @"\brdrart145";
        else if (borderValue == BorderValues.Tribal2)
            return @"\brdrart146";
        else if (borderValue == BorderValues.Tribal3)
            return @"\brdrart147";
        else if (borderValue == BorderValues.Tribal4)
            return @"\brdrart148";
        else if (borderValue == BorderValues.Tribal5)
            return @"\brdrart149";
        else if (borderValue == BorderValues.Tribal6)
            return @"\brdrart150";
        else if (borderValue == BorderValues.TwistedLines1)
            return @"\brdrart151";
        else if (borderValue == BorderValues.TwistedLines2)
            return @"\brdrart152";
        else if (borderValue == BorderValues.Vine)
            return @"\brdrart153";
        else if (borderValue == BorderValues.Waveline)
            return @"\brdrart154";
        else if (borderValue == BorderValues.WeavingAngles)
            return @"\brdrart155";
        else if (borderValue == BorderValues.WeavingBraid)
            return @"\brdrart156";
        else if (borderValue == BorderValues.WeavingRibbon)
            return @"\brdrart157";
        else if (borderValue == BorderValues.WeavingStrips)
            return @"\brdrart158";
        else if (borderValue == BorderValues.WhiteFlowers)
            return @"\brdrart159";
        else if (borderValue == BorderValues.Woodwork)
            return @"\brdrart160";
        else if (borderValue == BorderValues.XIllusions)
            return @"\brdrart161";
        else if (borderValue == BorderValues.ZanyTriangles)
            return @"\brdrart162";
        else if (borderValue == BorderValues.ZigZag)
            return @"\brdrart163";
        else if (borderValue == BorderValues.ZigZagStitch)
            return @"\brdrart164";

        else
            // Assume single
            return @"\brdrs";
        
    }
}
