using System;
using System.ComponentModel;

namespace EPayments.Model.Enums
{
    [Flags]
    public enum InternalAdminUserPermissionEnum
    {
        [Description("Преглед на справки")]
        ViewReferences = 1,
        [Description("Преглед на справки за разпределение")]
        DistributionReferences = 2,
        [Description("Управление")]
        Modify = 4
    }
}
