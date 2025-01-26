using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class ListHelpers
{
    public static AbstractNum AddOrderedListAbstractNumbering(this WordprocessingDocument document)
    {
        var numbering = document.GetOrCreateNumbering();

        var abstractNumId = numbering.Elements<AbstractNum>().Count() + 1;

        var abstractNum = new AbstractNum(
            new Level(
                new NumberingFormat() { Val = NumberFormatValues.Decimal },
                new LevelText() { Val = "%1." }
            )
            { LevelIndex = 0, StartNumberingValue = new StartNumberingValue() { Val = 1 } }
        )
        {
            AbstractNumberId = abstractNumId,
            MultiLevelType = new MultiLevelType { Val = MultiLevelValues.SingleLevel }
        };

        numbering.AddAbstractNumbering(abstractNum);

        return abstractNum;
    }

    public static AbstractNum AddBulletListAbstractNumbering(this WordprocessingDocument document)
    {
        var numbering = document.GetOrCreateNumbering();

        var abstractNumberId = numbering.Elements<AbstractNum>().Count() + 1;

        var abstractNum = new AbstractNum(
            new Level(
                new NumberingFormat() { Val = NumberFormatValues.Bullet },
                new LevelText() { Val = "·" }
            )
            { LevelIndex = 0 }
        )
        { AbstractNumberId = abstractNumberId };

        numbering.AddAbstractNumbering(abstractNum);
        return abstractNum;
    }

    public static NumberingInstance AddOrderedListNumbering(this WordprocessingDocument document, int abstractNumId, int? startFrom = null)
    {
        var numbering = document.GetOrCreateNumbering();
        var numId = numbering.Elements<NumberingInstance>().Count() + 1;
        var numberingInstance = new NumberingInstance(
            new AbstractNumId() { Val = abstractNumId }
        )
        { NumberID = numId };
        numbering.AddNumberingInstance(numberingInstance);

        if (startFrom != null)
        {
            var levelOverride = new LevelOverride
            {
                LevelIndex = 0,
                StartOverrideNumberingValue = new StartOverrideNumberingValue() { Val = startFrom }
            };
            numberingInstance.AppendChild(levelOverride);
        }


        return numberingInstance;
    }

    public static NumberingInstance AddBulletedListNumbering(this WordprocessingDocument document,
        AbstractNum? abstractNum = null)
    {
        var numbering = document.GetOrCreateNumbering();

        if (abstractNum == null)
        {
            var abstractNumberId = numbering.Elements<AbstractNum>().Count() + 1;

            abstractNum = new AbstractNum(
                new Level(
                    new NumberingFormat() { Val = NumberFormatValues.Bullet },
                    new LevelText() { Val = "·" }
                )
                { LevelIndex = 0 }
            )
            { AbstractNumberId = abstractNumberId };

            numbering.AddAbstractNumbering(abstractNum);
        }


        var numId = numbering.Elements<NumberingInstance>().Count() + 1;

        var numberingInstance = new NumberingInstance(
            new AbstractNumId() { Val = abstractNum.AbstractNumberId }
        )
        { NumberID = numId };

        numbering.AddNumberingInstance(numberingInstance);

        return numberingInstance;
    }

    public static void AddAbstractNumbering(this Numbering numbering, AbstractNum abstractNum)
    {
        numbering.InsertAfterLastOfType(abstractNum);
        numbering.Save();
    }

    public static void AddNumberingInstance(this Numbering numbering, NumberingInstance numberingInstance)
    {
        numbering.InsertAfterLastOfType(numberingInstance);
        numbering.Save();
    }

    public static Numbering GetOrCreateNumbering(this WordprocessingDocument document)
    {
        Debug.Assert(document.MainDocumentPart != null, "document.MainDocumentPart != null");

        if (document.MainDocumentPart.NumberingDefinitionsPart == null)
        {
            var part = document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
        }

        var numbering = document.MainDocumentPart.NumberingDefinitionsPart!.Numbering;
        return numbering;
    }

}
