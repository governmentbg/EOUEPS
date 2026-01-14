using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace EPayments.Common.Saml
{
    public enum EAuthResponseStatus
    {
        CanceledByUser,
        AuthenticationFailed,
        InvalidResponseXML,
        InvalidSignature,
        MissingEGN,
        NotDetectedQES,
        Success,
        InvalidMetadata
    }
}
