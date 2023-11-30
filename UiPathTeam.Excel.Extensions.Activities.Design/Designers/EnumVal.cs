using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace UiPathTeam.Excel.Extensions.Activities.Design.Designers
{
    public class EnumVal
    {
        public string Name { get; private set; }
        public Enum Value { get; private set; }

        protected internal EnumVal(string name, Enum value)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException(nameof(name));
            }
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            Name = name;
            Value = value;
        }
    }

    public class EnumVal<T> : EnumVal
    {
        protected EnumVal(string name, Enum value) : base(name, value)
        {
        }

        public static List<EnumVal> GetDirectionValues()
        {
            List<EnumVal> DirectionValues = new List<EnumVal>();

            Type enumType = typeof(T);
            Array enumValues = Enum.GetValues(enumType);

            foreach (Enum value in enumValues)
            {
                string name = enumType.GetEnumName(value);
                FieldInfo field = enumType.GetField(name);
                DescriptionAttribute descriptionAttribute = field?.GetCustomAttribute<DescriptionAttribute>();

                DirectionValues.Add(new EnumVal(descriptionAttribute?.Description ?? name, value));
            }

            return DirectionValues;
        }
    }
}
