namespace DocSharp.Binary.CommonTranslatorLib
{
    public interface IMapping<T> where T : IVisitable
    {
        void Apply(T visited);
    }
}