using System;

namespace FelipeCastro.Utils.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class CollumnAttribute : Attribute
    {
        public string Name { get; }

        public CollumnAttribute(string name)
        {
            Name = name;
        }
    }
}