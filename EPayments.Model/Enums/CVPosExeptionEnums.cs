using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum CVPosExeptionEnums
    {
        [Description("OherEception")]
        Other = 0,
        [Description("FaultEception")]
        Fault = 1,
        [Description("EndpointNotFoundException")]
        EndpointNotFound = 2,
        [Description("CommunicationException")]
        Communication =3,
        [Description("CertificatException")]
        Certificate = 4,
    }
}
