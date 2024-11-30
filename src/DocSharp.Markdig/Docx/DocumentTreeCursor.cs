using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Docx;

public class DocumentTreeCursor
{
    public OpenXmlCompositeElement Container { get; private set; }
    public OpenXmlElement? InsertAfter { get; private set; }

    private Stack<Tuple<OpenXmlCompositeElement, OpenXmlElement?>> _position = new();

    public DocumentTreeCursor(OpenXmlCompositeElement container, OpenXmlElement? insertAfter)
    {
        Container = container;
        InsertAfter = insertAfter;
    }

    public void GoInto(OpenXmlCompositeElement container)
    {
        _position.Push(new Tuple<OpenXmlCompositeElement, OpenXmlElement?>(Container, InsertAfter));
        Container = container;
        InsertAfter = null;
    }

    public void PopAndAdvanceAfter(OpenXmlElement? element)
    {
        var tuple = _position.Pop();
        if (element != null && tuple.Item1 != element.Parent)
        {
            throw new ArgumentException($"Element parent does not match container: expected {tuple.Item1}, got {element.Parent}");
        }
        (Container, InsertAfter) = (tuple.Item1, element);
    }

    public void SetAfter(OpenXmlElement? insertAfter)
    {
        if (insertAfter == null)
        {
            InsertAfter = null;
            return;
        }
        
        if (insertAfter.Parent is not OpenXmlCompositeElement parent)
        {
            throw new ArgumentException("Target position parent is not a composite element");
        }
        
        InsertAfter = insertAfter;
        Container = parent;
    }

    public void AssertContainerType<T>() where T : OpenXmlCompositeElement
    {
        if (Container is T) return;
        
        throw new ConstraintException(
            $"Wrong container type: expected {typeof(T).Name}, got {Container.GetType().Name}");
    }
    
    public void AssertContainerTypeNot<T>() where T : OpenXmlCompositeElement
    {
        if (Container is not T) return;
        
        throw new ConstraintException(
            $"Wrong container type: expected anything except {typeof(T).Name}, got {Container.GetType().Name}");
    }
    public void Write(OpenXmlElement element)
    {
        if (InsertAfter == null)
        {
            Container.AppendChild(element);
        }
        else
        {
            Container.InsertAfter(element, InsertAfter);
        }
    }
}
