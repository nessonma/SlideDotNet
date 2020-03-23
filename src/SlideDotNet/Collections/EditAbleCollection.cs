namespace SlideDotNet.Collections
{
    public abstract class EditAbleCollection<T> : LibraryCollection<T>
    {
        public abstract void Remove(T innerRow);
    }
}