using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Common.Helpers
{
    public static class EnumHelper
    {
        public static List<SelectListItem> GetEnumSelectListItems<T>(bool addEmptyItem = false)
        {
            var attributes = typeof(T).GetMembers()
                .SelectMany(e => e.GetCustomAttributes(typeof(DescriptionAttribute), true).Cast<DescriptionAttribute>())
                .ToList();

            List<SelectListItem> selectList = new List<SelectListItem>();

            if (addEmptyItem)
            {
                selectList.Add(new SelectListItem() { Value = "", Text = "--Избери--" });
            }

            foreach (Enum item in Enum.GetValues(typeof(T)))
            {
                selectList.Add(new SelectListItem() { Value = ((int)(ValueType)item).ToString(), Text = item.GetDescription() });
            }

            return selectList;
        }

        public static List<SelectListItem> FixEmptyItemValue(List<SelectListItem> selectListItems)
        {
            if (selectListItems != null)
            {
                foreach (var item in selectListItems.Where(e => e.Value == null))
                {
                    item.Value = string.Empty;
                }
            }

            return selectListItems;
        }

        public static string GetDescription(this Enum enumVal)
        {
            DescriptionAttribute attr = enumVal.GetAttributeOfType<DescriptionAttribute>();

            return attr.Description;
        }
        /// <summary>
        /// Gets an attribute on an enum field value
        /// </summary>
        /// <typeparam name="T">The type of the attribute you want to retrieve</typeparam>
        /// <param name="enumVal">The enum value</param>
        /// <returns>The attribute of type T that exists on the enum value</returns>
        /// <example>string desc = myEnumVariable.GetAttributeOfType<DescriptionAttribute>().Description;</example>
        public static T GetAttributeOfType<T>(this Enum enumVal) where T : System.Attribute
        {
            var type = enumVal.GetType();
            var memInfo = type.GetMember(enumVal.ToString());
            var attributes = memInfo[0].GetCustomAttributes(typeof(T), false);
            return (attributes.Length > 0) ? (T)attributes[0] : null;
        }
    }
}
