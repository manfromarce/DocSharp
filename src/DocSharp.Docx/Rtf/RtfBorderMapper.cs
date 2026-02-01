using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Rtf;

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

    internal static BorderValues GetBorderType(string borderValue)
    {
        if (borderValue == "brdrs")
            return BorderValues.Single;
        else if (borderValue == "brdrth")
            return BorderValues.Thick;
        else if (borderValue == "brdrdb")
            return BorderValues.Double;
        else if (borderValue == "brdrdot")
            return BorderValues.Dotted;
        else if (borderValue == "brdrdash")
            return BorderValues.Dashed;
        else if (borderValue == "brdrtriple")
            return BorderValues.Triple;
        else if (borderValue == "brdrwavy")
            return BorderValues.Wave;
        else if (borderValue == "brdrwavydb")
            return BorderValues.DoubleWave;
        else if (borderValue == "brdroutset")
            return BorderValues.Outset;
        else if (borderValue == "brdrinset")
            return BorderValues.Inset;
        else if (borderValue == "brdrdashsm")
            return BorderValues.DashSmallGap;
        else if (borderValue == "brdrdashd")
            return BorderValues.DotDash;
        else if (borderValue == "brdrdashdd")
            return BorderValues.DotDotDash;
        else if (borderValue == "brdrtnthsg")
            return BorderValues.ThickThinSmallGap;
        else if (borderValue == "brdrtnthmg")
            return BorderValues.ThickThinMediumGap;
        else if (borderValue == "brdrtnthlg")
            return BorderValues.ThickThinLargeGap;
        else if (borderValue == "brdrthtnsg")
            return BorderValues.ThinThickSmallGap;
        else if (borderValue == "brdrthtnmg")
            return BorderValues.ThinThickMediumGap;
        else if (borderValue == "brdrthtnlg")
            return BorderValues.ThinThickLargeGap;
        else if (borderValue == "brdrtnthtnsg")
            return BorderValues.ThinThickThinSmallGap;
        else if (borderValue == "brdrtnthtnmg")
            return BorderValues.ThinThickThinMediumGap;
        else if (borderValue == "brdrtnthtnlg")
            return BorderValues.ThinThickThinLargeGap;
        else if (borderValue == "brdremboss")
            return BorderValues.ThreeDEmboss;
        else if (borderValue == "brdrengrave")
            return BorderValues.ThreeDEngrave;
        else if (borderValue == "brdrdashdotstr")
            return BorderValues.DashDotStroked;
        else if (borderValue == "brdrnil")
            return BorderValues.Nil;
        else if (borderValue == "brdrnone")
            return BorderValues.None;
        else if (borderValue == "brdrart1")
            return BorderValues.Apples;
        else if (borderValue == "brdrart2")
            return BorderValues.ArchedScallops;
        else if (borderValue == "brdrart3")
            return BorderValues.BabyPacifier;
        else if (borderValue == "brdrart4")
            return BorderValues.BabyRattle;
        else if (borderValue == "brdrart5")
            return BorderValues.Balloons3Colors;
        else if (borderValue == "brdrart6")
            return BorderValues.BalloonsHotAir;
        else if (borderValue == "brdrart7")
            return BorderValues.BasicBlackDashes;
        else if (borderValue == "brdrart8")
            return BorderValues.BasicBlackDots;
        else if (borderValue == "brdrart9")
            return BorderValues.BasicBlackSquares;
        else if (borderValue == "brdrart10")
            return BorderValues.BasicThinLines;
        else if (borderValue == "brdrart11")
            return BorderValues.BasicWhiteDashes;
        else if (borderValue == "brdrart12")
            return BorderValues.BasicWhiteDots;
        else if (borderValue == "brdrart13")
            return BorderValues.BasicWhiteSquares;
        else if (borderValue == "brdrart14")
            return BorderValues.BasicWideInline;
        else if (borderValue == "brdrart15")
            return BorderValues.BasicWideMidline;
        else if (borderValue == "brdrart16")
            return BorderValues.BasicWideOutline;
        else if (borderValue == "brdrart17")
            return BorderValues.Bats;
        else if (borderValue == "brdrart18")
            return BorderValues.Birds;
        else if (borderValue == "brdrart19")
            return BorderValues.BirdsFlight;
        else if (borderValue == "brdrart20")
            return BorderValues.Cabins;
        else if (borderValue == "brdrart21")
            return BorderValues.CakeSlice;
        else if (borderValue == "brdrart22")
            return BorderValues.CandyCorn;
        else if (borderValue == "brdrart23")
            return BorderValues.CelticKnotwork;
        else if (borderValue == "brdrart24")
            return BorderValues.CertificateBanner;
        else if (borderValue == "brdrart25")
            return BorderValues.ChainLink;
        else if (borderValue == "brdrart26")
            return BorderValues.ChampagneBottle;
        else if (borderValue == "brdrart27")
            return BorderValues.CheckedBarBlack;
        else if (borderValue == "brdrart28")
            return BorderValues.CheckedBarColor;
        else if (borderValue == "brdrart29")
            return BorderValues.Checkered;
        else if (borderValue == "brdrart30")
            return BorderValues.ChristmasTree;
        else if (borderValue == "brdrart31")
            return BorderValues.CirclesLines;
        else if (borderValue == "brdrart32")
            return BorderValues.CirclesRectangles;
        else if (borderValue == "brdrart33")
            return BorderValues.ClassicalWave;
        else if (borderValue == "brdrart34")
            return BorderValues.Clocks;
        else if (borderValue == "brdrart35")
            return BorderValues.Compass;
        else if (borderValue == "brdrart36")
            return BorderValues.Confetti;
        else if (borderValue == "brdrart37")
            return BorderValues.ConfettiGrays;
        else if (borderValue == "brdrart38")
            return BorderValues.ConfettiOutline;
        else if (borderValue == "brdrart39")
            return BorderValues.ConfettiStreamers;
        else if (borderValue == "brdrart40")
            return BorderValues.ConfettiWhite;
        else if (borderValue == "brdrart41")
            return BorderValues.CornerTriangles;
        else if (borderValue == "brdrart42")
            return BorderValues.CouponCutoutDashes;
        else if (borderValue == "brdrart43")
            return BorderValues.CouponCutoutDots;
        else if (borderValue == "brdrart44")
            return BorderValues.CrazyMaze;
        else if (borderValue == "brdrart45")
            return BorderValues.CreaturesButterfly;
        else if (borderValue == "brdrart46")
            return BorderValues.CreaturesFish;
        else if (borderValue == "brdrart47")
            return BorderValues.CreaturesInsects;
        else if (borderValue == "brdrart48")
            return BorderValues.CreaturesLadyBug;
        else if (borderValue == "brdrart49")
            return BorderValues.CrossStitch;
        else if (borderValue == "brdrart50")
            return BorderValues.Cup;
        else if (borderValue == "brdrart51")
            return BorderValues.DecoArch;
        else if (borderValue == "brdrart52")
            return BorderValues.DecoArchColor;
        else if (borderValue == "brdrart53")
            return BorderValues.DecoBlocks;
        else if (borderValue == "brdrart54")
            return BorderValues.DiamondsGray;
        else if (borderValue == "brdrart55")
            return BorderValues.DoubleD;
        else if (borderValue == "brdrart56")
            return BorderValues.DoubleDiamonds;
        else if (borderValue == "brdrart57")
            return BorderValues.Earth1;
        else if (borderValue == "brdrart58")
            return BorderValues.Earth2;
        else if (borderValue == "brdrart59")
            return BorderValues.EclipsingSquares1;
        else if (borderValue == "brdrart60")
            return BorderValues.EclipsingSquares2;
        else if (borderValue == "brdrart61")
            return BorderValues.EggsBlack;
        else if (borderValue == "brdrart62")
            return BorderValues.Fans;
        else if (borderValue == "brdrart63")
            return BorderValues.Film;
        else if (borderValue == "brdrart64")
            return BorderValues.Firecrackers;
        else if (borderValue == "brdrart65")
            return BorderValues.FlowersBlockPrint;
        else if (borderValue == "brdrart66")
            return BorderValues.FlowersDaisies;
        else if (borderValue == "brdrart67")
            return BorderValues.FlowersModern1;
        else if (borderValue == "brdrart68")
            return BorderValues.FlowersModern2;
        else if (borderValue == "brdrart69")
            return BorderValues.FlowersPansy;
        else if (borderValue == "brdrart70")
            return BorderValues.FlowersRedRose;
        else if (borderValue == "brdrart71")
            return BorderValues.FlowersRoses;
        else if (borderValue == "brdrart72")
            return BorderValues.FlowersTeacup;
        else if (borderValue == "brdrart73")
            return BorderValues.FlowersTiny;
        else if (borderValue == "brdrart74")
            return BorderValues.Gems;
        else if (borderValue == "brdrart75")
            return BorderValues.GingerbreadMan;
        else if (borderValue == "brdrart76")
            return BorderValues.Gradient;
        else if (borderValue == "brdrart77")
            return BorderValues.Handmade1;
        else if (borderValue == "brdrart78")
            return BorderValues.Handmade2;
        else if (borderValue == "brdrart79")
            return BorderValues.HeartBalloon;
        else if (borderValue == "brdrart80")
            return BorderValues.HeartGray;
        else if (borderValue == "brdrart81")
            return BorderValues.Hearts;
        else if (borderValue == "brdrart82")
            return BorderValues.HeebieJeebies;
        else if (borderValue == "brdrart83")
            return BorderValues.Holly;
        else if (borderValue == "brdrart84")
            return BorderValues.HouseFunky;
        else if (borderValue == "brdrart85")
            return BorderValues.Hypnotic;
        else if (borderValue == "brdrart86")
            return BorderValues.IceCreamCones;
        else if (borderValue == "brdrart87")
            return BorderValues.LightBulb;
        else if (borderValue == "brdrart88")
            return BorderValues.Lightning1;
        else if (borderValue == "brdrart89")
            return BorderValues.Lightning2;
        else if (borderValue == "brdrart91")
            return BorderValues.MapleLeaf;
        else if (borderValue == "brdrart92")
            return BorderValues.MapleMuffins;
        else if (borderValue == "brdrart90")
            return BorderValues.MapPins;
        else if (borderValue == "brdrart93")
            return BorderValues.Marquee;
        else if (borderValue == "brdrart94")
            return BorderValues.MarqueeToothed;
        else if (borderValue == "brdrart95")
            return BorderValues.Moons;
        else if (borderValue == "brdrart96")
            return BorderValues.Mosaic;
        else if (borderValue == "brdrart97")
            return BorderValues.MusicNotes;
        else if (borderValue == "brdrart98")
            return BorderValues.Northwest;
        else if (borderValue == "brdrart99")
            return BorderValues.Ovals;
        else if (borderValue == "brdrart100")
            return BorderValues.Packages;
        else if (borderValue == "brdrart101")
            return BorderValues.PalmsBlack;
        else if (borderValue == "brdrart102")
            return BorderValues.PalmsColor;
        else if (borderValue == "brdrart103")
            return BorderValues.PaperClips;
        else if (borderValue == "brdrart104")
            return BorderValues.Papyrus;
        else if (borderValue == "brdrart105")
            return BorderValues.PartyFavor;
        else if (borderValue == "brdrart106")
            return BorderValues.PartyGlass;
        else if (borderValue == "brdrart107")
            return BorderValues.Pencils;
        else if (borderValue == "brdrart108")
            return BorderValues.People;
        else if (borderValue == "brdrart110")
            return BorderValues.PeopleHats;
        else if (borderValue == "brdrart109")
            return BorderValues.PeopleWaving;
        else if (borderValue == "brdrart111")
            return BorderValues.Poinsettias;
        else if (borderValue == "brdrart112")
            return BorderValues.PostageStamp;
        else if (borderValue == "brdrart113")
            return BorderValues.Pumpkin1;
        else if (borderValue == "brdrart115")
            return BorderValues.PushPinNote1;
        else if (borderValue == "brdrart114")
            return BorderValues.PushPinNote2;
        else if (borderValue == "brdrart116")
            return BorderValues.Pyramids;
        else if (borderValue == "brdrart117")
            return BorderValues.PyramidsAbove;
        else if (borderValue == "brdrart118")
            return BorderValues.Quadrants;
        else if (borderValue == "brdrart119")
            return BorderValues.Rings;
        else if (borderValue == "brdrart120")
            return BorderValues.Safari;
        else if (borderValue == "brdrart121")
            return BorderValues.Sawtooth;
        else if (borderValue == "brdrart122")
            return BorderValues.SawtoothGray;
        else if (borderValue == "brdrart123")
            return BorderValues.ScaredCat;
        else if (borderValue == "brdrart124")
            return BorderValues.Seattle;
        else if (borderValue == "brdrart125")
            return BorderValues.ShadowedSquares;
        else if (borderValue == "brdrart126")
            return BorderValues.SharksTeeth;
        else if (borderValue == "brdrart127")
            return BorderValues.ShorebirdTracks;
        else if (borderValue == "brdrart128")
            return BorderValues.Skyrocket;
        else if (borderValue == "brdrart129")
            return BorderValues.SnowflakeFancy;
        else if (borderValue == "brdrart130")
            return BorderValues.Snowflakes;
        else if (borderValue == "brdrart131")
            return BorderValues.Sombrero;
        else if (borderValue == "brdrart132")
            return BorderValues.Southwest;
        else if (borderValue == "brdrart133")
            return BorderValues.Stars;
        else if (borderValue == "brdrart135")
            return BorderValues.Stars3d;
        else if (borderValue == "brdrart136")
            return BorderValues.StarsBlack;
        else if (borderValue == "brdrart137")
            return BorderValues.StarsShadowed;
        else if (borderValue == "brdrart134")
            return BorderValues.StarsTop;
        else if (borderValue == "brdrart138")
            return BorderValues.Sun;
        else if (borderValue == "brdrart139")
            return BorderValues.Swirligig;
        else if (borderValue == "brdrart140")
            return BorderValues.TornPaper;
        else if (borderValue == "brdrart141")
            return BorderValues.TornPaperBlack;
        else if (borderValue == "brdrart142")
            return BorderValues.Trees;
        else if (borderValue == "brdrart143")
            return BorderValues.TriangleParty;
        else if (borderValue == "brdrart144")
            return BorderValues.Triangles;
        else if (borderValue == "brdrart145")
            return BorderValues.Tribal1;
        else if (borderValue == "brdrart146")
            return BorderValues.Tribal2;
        else if (borderValue == "brdrart147")
            return BorderValues.Tribal3;
        else if (borderValue == "brdrart148")
            return BorderValues.Tribal4;
        else if (borderValue == "brdrart149")
            return BorderValues.Tribal5;
        else if (borderValue == "brdrart150")
            return BorderValues.Tribal6;
        else if (borderValue == "brdrart151")
            return BorderValues.TwistedLines1;
        else if (borderValue == "brdrart152")
            return BorderValues.TwistedLines2;
        else if (borderValue == "brdrart153")
            return BorderValues.Vine;
        else if (borderValue == "brdrart154")
            return BorderValues.Waveline;
        else if (borderValue == "brdrart155")
            return BorderValues.WeavingAngles;
        else if (borderValue == "brdrart156")
            return BorderValues.WeavingBraid;
        else if (borderValue == "brdrart157")
            return BorderValues.WeavingRibbon;
        else if (borderValue == "brdrart158")
            return BorderValues.WeavingStrips;
        else if (borderValue == "brdrart159")
            return BorderValues.WhiteFlowers;
        else if (borderValue == "brdrart160")
            return BorderValues.Woodwork;
        else if (borderValue == "brdrart161")
            return BorderValues.XIllusions;
        else if (borderValue == "brdrart162")
            return BorderValues.ZanyTriangles;
        else if (borderValue == "brdrart163")
            return BorderValues.ZigZag;
        else if (borderValue == "brdrart164")
            return BorderValues.ZigZagStitch;

        return BorderValues.Single;
    }
}
