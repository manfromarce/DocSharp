using System;
using DocSharp.Binary.CommonTranslatorLib;

namespace DocSharp.Binary.DocFileFormat
{
    public class SinglePropertyModifier : IVisitable
    {
        public enum OperationCode
        { 
            //Paragraph SPRMs
            sprmPIstd=0x4600,
            sprmPIstdPermute=0xC601,
            sprmPIncLvl=0x2602,
            sprmPJc=0x2461,
            sprmPJc80=0x2403,
            sprmPFSideBySide=0x2404,
            sprmPFKeep=0x2405,
            sprmPFKeepFollow=0x2406,
            sprmPFPageBreakBefore=0x2407,
            sprmPBrcl=0x2408,
            sprmPBrcp=0x2409,
            sprmPIlvl=0x260A,
            sprmPIlfo=0x460B,
            sprmPFNoLineNumb=0x240C,
            sprmPChgTabsPapx=0xC60D,
            sprmPDxaLeft=0x845e,
            sprmPDxaLeft80=0x840f,
            sprmPDxaLeft1=0x8460,
            sprmPDxaLeft180=0x8411,
            sprmPDxaRight=0x845d,
            sprmPDxaRight80=0x840e,
            sprmPDxcLeft=0x4456,
            sprmPDxcLeft1=0x4457,
            sprmPDxcRight=0x4455,
            sprmPNest=0x465f,
            sprmPNest80=0x4610,
            sprmPDyaLine=0x6412,
            sprmPDyaBefore=0xA413,
            sprmPDyaAfter=0xA414,
            sprmPFDyaAfterAuto=0x245c,
            sprmPFDyaBeforeAuto=0x245b,
            sprmPDylAfter=0x4459,
            sprmPDylBefore=0x4458,
            sprmPChgTabs=0xC615,
            sprmPFInTable=0x2416,
            sprmPFTtp=0x2417,
            sprmPDxaAbs=0x8418,
            sprmPDyaAbs=0x8419,
            sprmPDxaWidth=0x841A,
            sprmPPc=0x261B,
            sprmPBrcTop10=0x461C,
            sprmPBrcLeft10=0x461D,
            sprmPBrcBottom10=0x461E,
            sprmPBrcRight10=0x461F,
            sprmPBrcBetween10=0x4620,
            sprmPBrcBar10=0x4621,
            sprmPDxaFromText10=0x4622,
            sprmPWr=0x2423,
            sprmPBrcBar=0xc653,
            sprmPBrcBar70=0x4629,
            sprmPBrcBar80=0x6629,
            sprmPBrcBetween=0xc652,
            sprmPBrcBetween70=0x4428,
            sprmPBrcBetween80=0x6428,
            sprmPBrcBottom=0xc650,
            sprmPBrcBottom70=0x4426,
            sprmPBrcBottom80=0x6426,
            sprmPBrcLeft=0xc64f,
            sprmPBrcLeft70=0x4425,
            sprmPBrcLeft80=0x6425,
            sprmPBrcRight=0xc651,
            sprmPBrcRight70=0x4427,
            sprmPBrcRight80=0x6427,
            sprmPBrcTop=0xc64e,
            sprmPBrcTop70=0x4424,
            sprmPBrcTop80=0x6424,
            sprmPFNoAutoHyph=0x242A,
            sprmPWHeightAbs=0x442B,
            sprmPDcs=0x442C,
            sprmPShd80=0x442D,
            sprmPShd=0xc64d,
            sprmPDyaFromText=0x842E,
            sprmPDxaFromText=0x842F,
            sprmPFLocked=0x2430,
            sprmPFWidowControl=0x2431,
            sprmPRuler=0xC632,
            sprmPFKinsoku=0x2433,
            sprmPFWordWrap=0x2434,
            sprmPFOverflowPunct=0x2435,
            sprmPFTopLinePunct=0x2436,
            sprmPFAutoSpaceDE=0x2437,
            sprmPFAutoSpaceDN=0x2438,
            sprmPWAlignFont=0x4439,
            sprmPFrameTextFlow=0x443A,
            sprmPISnapBaseLine=0x243B,
            sprmPAnld80=0xC63E,
            sprmPAnldCv=0x6654,
            sprmPPropRMark=0xC63F,
            sprmPOutLvl=0x2640,
            sprmPFBiDi=0x2441,
            sprmPFNumRMIns=0x2443,
            sprmPNumRM=0xC645,
            sprmPHugePapx=0x6645,
            sprmPFUsePgsuSettings=0x2447,
            sprmPFAdjustRight=0x2448,
            sprmPDtap=0x664a,
            sprmPFInnerTableCell=0x244b,
            sprmPFInnerTtp=0x244c,
            sprmPFNoAllowOverlap=0x2462,
            sprmPItap=0x6649,
            sprmPWall=0x2664,
            sprmPIpgp=0x6465,
            sprmPCnf=0xc666,
            sprmPRsid=0x6467,
            sprmPIstdList=0x4468,
            sprmPIstdListPermute=0xc669,
            sprmPDyaBeforeNotCp0=0xa46a,
            sprmPTableProps=0x646b,
            sprmPTIstdInfo=0xc66c,
            sprmPFContextualSpacing=0x246d,
            sprmPRpf=0x246e,
            sprmPPropRMark90=0xc66f,

            //Character SPRMs
            sprmCFRMarkDel=0x0800,
            sprmCFRMark=0x0801,
            sprmCFFldVanish=0x0802,
            sprmCFSdtVanish=0x2A90,
            sprmCPicLocation=0x6A03,
            sprmCIbstRMark=0x4804,
            sprmCDttmRMark=0x6805,
            sprmCFData=0x0806,
            sprmCIdslRMark=0x4807,
            sprmCChs=0xEA08,
            sprmCSymbol=0x6A09,
            sprmCFOle2=0x080A,
            sprmCIdCharType=0x480B,
            sprmCHighlight=0x2A0C,
            sprmCObjLocation=0x680E,
            sprmCObjpLocation=0x680e,
            sprmCFFtcAsciSymb=0x2A10,
            sprmCIstd=0x4A30,
            sprmCIstdPermute=0xCA31,
            sprmCDefault=0x2A32,
            sprmCPlain=0x2A33,
            sprmCKcd=0x2A34,
            sprmCFBold=0x0835,
            sprmCFItalic=0x0836,
            sprmCFStrike=0x0837,
            sprmCFOutline=0x0838,
            sprmCFShadow=0x0839,
            sprmCFSmallCaps=0x083A,
            sprmCFCaps=0x083B,
            sprmCFVanish=0x083C,
            sprmCFtcDefault=0x4A3D,
            sprmCKul=0x2A3E,
            sprmCSizePos=0xEA3F,
            sprmCDxaSpace=0x8840,
            sprmCLid=0x4A41,
            sprmCIco=0x2A42,
            sprmCHps=0x4A43,
            sprmCHpsInc=0x2A44,
            sprmCHpsPos=0x4845,
            sprmCHpsPosAdj=0x2A46,
            sprmCMajority=0xCA47,
            sprmCIss=0x2A48,
            sprmCHpsNew50=0xCA49,
            sprmCHpsInc1=0xCA4A,
            sprmCHpsKern=0x484B,
            sprmCMajority50=0xCA4C,
            sprmCHpsMul=0x4A4D,
            sprmCHresi=0x484e,
            sprmCRgFtc0=0x4A4F,
            sprmCRgFtc1=0x4A50,
            sprmCRgFtc2=0x4A51,
            sprmCCharScale=0x4852,
            sprmCFDStrike=0x2A53,
            sprmCFImprint=0x0854,
            sprmCFSpec=0x0855,
            sprmCFObj=0x0856,
            sprmCPropRMark1=0xCA57,
            sprmCFEmboss=0x0858,
            sprmCSfxText=0x2859,
            sprmCFBiDi=0x085A,
            sprmCFDiacColor=0x085B,
            sprmCFBoldBi=0x085C,
            sprmCFItalicBi=0x085D,
            sprmCFtcBi=0x4A5E,
            sprmCLidBi=0x485F,
            sprmCIcoBi=0x4A60,
            sprmCHpsBi=0x4A61,
            sprmCDispFldRMark=0xCA62,
            sprmCIbstRMarkDel=0x4863,
            sprmCDttmRMarkDel=0x6864,
            SprmCBrc80=0x6865,
            sprmCBrc=0xca72,
            sprmCShd80=0x4866,
            sprmCShd=0xca71,
            sprmCIdslRMarkDel=0x4867,
            sprmCFUsePgsuSettings=0x0868,
            sprmCCpg=0x486B,
            sprmCRgLid0_80=0x486D,
            sprmCRgLid0=0x4873,
            sprmCRgLid1_80=0x486E,
            sprmCRgLid1=0x4874,
            sprmCIdctHint=0x286F,
            sprmCCv=0x6870,
            sprmCCvPermute=0xca7c,
            sprmCCvUl=0x6877,
            sprmCFBoldPresent=0x287d,
            sprmCFELayout=0xca78,
            sprmCFItalicPresent=0x287e,
            sprmCFitText=0xca76,
            sprmCFLangApplied=0x2a7a,
            sprmCFNoProof=0x875,
            sprmCFWebHidden=0x811,
            sprmCHsp=0x6a12,
            sprmCLbcCRJ=0x2879,
            sprmCNewIbstRM=0xca13,
            sprmCTransNoProof0=0x287f,
            sprmCTransNoProof1=0x2880,
            sprmCFRMMove=0x2814,
            sprmCRsidProp=0x6815,
            sprmCRsidText=0x6816,
            sprmCRsidRMDel=0x6817,
            sprmCFSpecVanish=0x0818,
            sprmCFComplexScripts=0x0882,
            sprmCWall=0x2a83,
            sprmCPbi=0xca84,
            sprmCCnf=0xca85,
            sprmCNeedFontFixup=0x2a86,
            sprmCPbiIBullet=0x6887,
            sprmCPbiGrf=0x4888,
            sprmCPropRMark2=0xca89,

            //Picture SPRMs
            sprmPicBrcl=0x2E00,
            sprmPicScale=0xCE01,
            sprmPicBrcTop80=0x6C02,
            sprmPicBrcBottom=0xce0a,
            sprmPicBrcBottom70=0x4c04,
            sprmPicBrcLeft80=0x6C03,
            sprmPicBrcLeft=0xce09,
            sprmPicBrcLeft70=0x4c03,
            sprmPicBrcBottom80=0x6C04,
            sprmPicBrcRight=0xce0b,
            sprmPicBrcRight70=0x4c05,
            sprmPicBrcRight80=0x6C05,
            sprmPicBrcTop=0xce08,
            sprmPicBrcTop70=0x4c02,
            sprmPicSpare4=0xce06,
            sprmCFOle2WasHere=0xce07,

            //Section SPRMs
            sprmScnsPgn=0x3000,
            sprmSiHeadingPgn=0x3001,
            sprmSOlstAnm=0xD202,
            sprmSOlstAnm80=0xd202,
            sprmSOlstCv=0xd238,
            sprmSDxaColWidth=0xF203,
            sprmSDxaColSpacing=0xF204,
            sprmSFEvenlySpaced=0x3005,
            sprmSFProtected=0x3006,
            sprmSDmBinFirst=0x5007,
            sprmSDmBinOther=0x5008,
            sprmSBkc=0x3009,
            sprmSFTitlePage=0x300A,
            sprmSCcolumns=0x500B,
            sprmSDxaColumns=0x900C,
            sprmSFAutoPgn=0x300D,
            sprmSNfcPgn=0x300E,
            sprmSDyaPgn=0xB00F,
            sprmSDxaPgn=0xB010,
            sprmSFPgnRestart=0x3011,
            sprmSFEndnote=0x3012,
            sprmSLnc=0x3013,
            sprmSGprfIhdt=0x3014,
            sprmSNLnnMod=0x5015,
            sprmSDxaLnn=0x9016,
            sprmSDyaHdrTop=0xB017,
            sprmSDyaHdrBottom=0xB018,
            sprmSLBetween=0x3019,
            sprmSVjc=0x301A,
            sprmSLnnMin=0x501B,
            sprmSPgnStart=0x501C,
            sprmSBOrientation=0x301D,
            sprmSXaPage=0xB01F,
            sprmSYaPage=0xB020,
            sprmSDxaLeft=0xB021,
            sprmSDxaRight=0xB022,
            sprmSDyaTop=0x9023,
            sprmSDyaBottom=0x9024,
            sprmSDzaGutter=0xB025,
            sprmSDmPaperReq=0x5026,
            sprmSPropRMark1=0xD227,
            sprmSFBiDi=0x3228,
            sprmSFFacingCol=0x3229,
            sprmSFRTLGutter=0x322A,
            sprmSBrcTop80=0x702B,
            sprmSBrcTop=0xd234,
            sprmSBrcLeft80=0x702C,
            sprmSBrcLeft=0xd235,
            sprmSBrcBottom80=0x702d,
            sprmSBrcBottom=0xd236,
            sprmSBrcRight80=0x702e,
            sprmSBrcRight=0xd237,
            sprmSPgbProp=0x522F,
            sprmSDxtCharSpace=0x7030,
            sprmSDyaLinePitch=0x9031,
            sprmSClm=0x5032,
            sprmSTextFlow=0x5033,
            sprmSWall=0x3239,
            sprmSRsid=0x703a,
            sprmSFpc=0x303b,
            sprmSRncFtn=0x303c,
            sprmSEpc=0x303d,
            sprmSRncEdn=0x303e,
            sprmSNFtn=0x503f,
            sprmSNfcFtnRef=0x5040,
            sprmSNEdn=0x5041,
            sprmSNfcEdnRef=0x5042,
            sprmSPropRMark2=0xd243,

            //Table SPRMs
            sprmTDefTable=0xD608,
            sprmTDefTable10=0xD606,
            sprmTDefTableShd97=0xD609,
            sprmTDefTableShd=0xd612,
            sprmTDefTableShd2nd=0xd616,
            sprmTDefTableShd3rd=0xd60c,
            sprmTDelete=0x5622,
            sprmTDiagLine=0xd630,
            sprmTDiagLine80=0xd62a,
            sprmTDxaCol=0x7623,
            sprmTDxaGapHalf=0x9602,
            sprmTDxaLeft=0x9601,
            sprmTDyaRowHeight=0x9407,
            sprmTFBiDi80=0x560b,
            sprmTFCantSplit=0x3403,
            sprmTHTMLProps=0x740C,
            sprmTInsert=0x7621,
            sprmTJc=0x5400,
            sprmTMerge=0x5624,
            sprmTSetBrc80=0xD620,
            sprmTSetBrc10=0xD626,
            sprmTSetBrc=0xd62f,
            sprmTSetShd80=0x7627,
            sprmTSetShdOdd80=0x7628,
            sprmTSetShd=0xd62d,
            sprmTSetShdOdd=0xd62e,
            sprmTSetShdTable=0xd660,
            sprmTSplit=0x5625,
            sprmTTableBorders=0xd613,
            sprmTTableBorders80=0xd605,
            sprmTTableHeader=0x3404,
            sprmTTextFlow=0x7629,
            sprmTTlp=0x740A,
            sprmTVertAlign=0xD62C,
            sprmTVertMerge=0xD62B,
            sprmTFCellNoWrap=0xd639,
            sprmTFitText=0xf636,
            sprmTFKeepFollow=0x3619,
            sprmTFNeverBeenAutofit=0x3663,
            sprmTFNoAllowOverlap=0x3465,
            sprmTPc=0x360d,
            sprmTBrcBottomCv=0xd61c,
            sprmTBrcLeftCv=0xd61b,
            sprmTBrcRightCv=0xd61d,
            sprmTBrcTopCv=0xd61a,
            sprmTCellBrcType=0xd662,
            sprmTCellPadding=0xd632,
            sprmTCellPaddingDefault=0xd634,
            sprmTCellPaddingOuter=0xd638,
            sprmTCellSpacing=0xd631,
            sprmTCellSpacingDefault=0xd633,
            sprmTCellSpacingOuter=0xd637,
            sprmTCellWidth=0xd635,
            sprmTDxaAbs=0x940e,
            sprmTDxaFromText=0x9410,
            sprmTDxaFromTextRight=0x941e,
            sprmTDyaAbs=0x940f,
            sprmTDyaFromText=0x9411,
            sprmTDyaFromTextBottom=0x941f,
            sprmTFAutofit=0x3615,
            sprmTTableWidth=0xf614,
            sprmTWidthAfter=0xf618,
            sprmTWidthBefore=0xf617,
            sprmTWidthIndent=0xf661,
            sprmTIstd=0x563a,
            sprmTSetShdRaw=0xd63b,
            sprmTSetShdOddRaw=0xd63c,
            sprmTIstdPermute=0xd63d,
            sprmTCellPaddingStyle=0xd63e,
            sprmTFCantSplit90=0x3466,
            sprmTPropRMark=0xd667,
            sprmTWall=0x3668,
            sprmTIpgp=0x7469,
            sprmTCnf=0xd66a,
            sprmTSetShdTableDef=0xd66b,
            sprmTDiagLine2nd=0xd66c,
            sprmTDiagLine3rd=0xd66d,
            sprmTDiagLine4th=0xd66e,
            sprmTDiagLine5th=0xd66f,
            sprmTDefTableShdRaw=0xd670,
            sprmTDefTableShdRaw2nd=0xd671,
            sprmTDefTableShdRaw3rd=0xd672,
            sprmTSetShdRowFirst=0xd673,
            sprmTSetShdRowLast=0xd674,
            sprmTSetShdColFirst=0xd675,
            sprmTSetShdColLast=0xd676,
            sprmTSetShdBand1=0xd677,
            sprmTSetShdBand2=0xd678,
            sprmTRsid=0x7479,
            sprmTCellWidthStyle=0xf47a,
            sprmTCellPaddingStyleBad=0xd67b,
            sprmTCellVertAlignStyle=0x347c,
            sprmTCellNoWrapStyle=0x347d,
            sprmTCellFitTextStyle=0x347e,
            sprmTCellBrcTopStyle=0xd47f,
            sprmTCellBrcBottomStyle=0xd680,
            sprmTCellBrcLeftStyle=0xd681,
            sprmTCellBrcRightStyle=0xd682,
            sprmTCellBrcInsideHStyle=0xd683,
            sprmTCellBrcInsideVStyle=0xd684,
            sprmTCellBrcTL2BRStyle=0xd685,
            sprmTCellBrcTR2BLStyle=0xd686,
            sprmTCellShdStyle=0xd687,
            sprmTCHorzBands=0x3488,
            sprmTCVertBands=0x3489,
            sprmTJcRow=0x548a,
            sprmTTableBrcTop=0xd68b,
            sprmTTableBrcLeft=0xd68c,
            sprmTTableBrcBottom=0xd68d,
            sprmTTableBrcRight=0xd68e,
            sprmTTableBrcInsideH=0xd68f,
            sprmTTableBrcInsideV=0xd690,
            sprmTFBiDi=0x560b,
            sprmTFBiDi90=0x5664
        }

        /// <summary>
        /// Identifies the type of a SPRM
        /// </summary>
        public enum SprmType
        {
            PAP = 1,
            CHP = 2,
            PIC = 3,
            SEP = 4,
            TAP = 5
        }

        /// <summary>
        /// The operation code identifies the property of the 
        /// PAP/CHP/PIC/SEP/TAP which sould be modified
        /// </summary>
        public OperationCode OpCode;

        /// <summary>
        /// This SPRM requires special handling
        /// </summary>
        public bool fSpec;

        /// <summary>
        /// The type of the SPRM
        /// </summary>
        public SprmType Type;

        /// <summary>
        /// The arguments which is applied to the property
        /// </summary>
        public byte[] Arguments;

        /// <summary>
        /// parses the byte to retrieve a SPRM
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public SinglePropertyModifier(byte[] bytes)
        {
            //first 2 bytes are the operation code ...
            this.OpCode = (OperationCode)System.BitConverter.ToUInt16(bytes, 0);

            //... whereof bit 9 is fSpec ...
            uint j = (uint)this.OpCode << 22;
            j = j >> 31;
            if (j == 1)
                this.fSpec = true;
            else
                this.fSpec = false;

            //... and bits 10,11,12 are the type ...
            uint i = (uint)this.OpCode << 19;
            i = i >> 29;
            this.Type = (SprmType)i;

            //... and last 3 bits are the spra
            byte spra = (byte)((int)this.OpCode >> 13);
            byte opSize = GetOperandSize(spra);
            if (opSize == 255)
            {
                switch (this.OpCode)
                {
                    case OperationCode.sprmTDefTable:
                    case OperationCode.sprmTDefTable10:
                        //the variable length stand in the bytes 2 and 3
                        short opSizeTable = System.BitConverter.ToInt16(bytes, 2);
                        //and the arguments start at the byte after that (byte3)
                        this.Arguments = new byte[opSizeTable-1];
                        //Arguments start at byte 4
                        Array.Copy(bytes, 4, this.Arguments, 0, this.Arguments.Length);
                        break;
                    case OperationCode.sprmPChgTabs:
                        this.Arguments = new byte[bytes[2]];
                        Array.Copy(bytes, 3, this.Arguments, 0, this.Arguments.Length);
                        break;
                    default:
                        //the variable length stand in the byte after the opcode (byte2)
                        opSize = bytes[2];
                        //and the arguments start at the byte after that (byte3)
                        this.Arguments = new byte[opSize];
                        Array.Copy(bytes, 3, this.Arguments, 0, this.Arguments.Length);
                        break;
                }
            }
            else
            {
                this.Arguments = new byte[opSize];
                Array.Copy(bytes, 2, this.Arguments, 0, this.Arguments.Length);
            }
        }

        /// <summary>
        /// Get be used to get the size of the sprm's operand.
        /// Returns 0 if the Operation failed and 255 if the size is variable
        /// </summary>
        /// <param name="spra">the 3 bits for spra (as byte)</param>
        /// <returns>the size (as byte)</returns>
        public static byte GetOperandSize(byte spra)
        {
            switch (spra)
            {
                case 0: return 1;
                case 1: return 1;
                case 2: return 2;
                case 3: return 4;
                case 4: return 2;
                case 5: return 2;
                case 6: return 255;
                case 7: return 3;
                default: return 0;
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            (mapping as IMapping<SinglePropertyModifier>)?.Apply(this);
        }

        #endregion
    }
}
