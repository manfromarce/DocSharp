namespace DocSharp.Binary.CommonTranslatorLib;

public interface IVisitable
{
    void Convert<T>(T mapping);
}