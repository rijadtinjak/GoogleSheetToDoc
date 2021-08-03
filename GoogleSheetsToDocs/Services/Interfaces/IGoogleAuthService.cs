using Google.Apis.Auth.OAuth2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GoogleSheetsToDocs.Services.Interfaces
{
    public interface IGoogleAuthService
    {
        UserCredential GetUserCredential();
        
    }
}
