using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using GoogleSheetsToDocs.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsToDocs.Services.Implementations
{
    public class GoogleAuthService : IGoogleAuthService
    {
       private readonly string[] Scopes = {
            SheetsService.Scope.SpreadsheetsReadonly,
            DocsService.Scope.Documents
        };

        public UserCredential GetUserCredential()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            return credential;
        }
    }
}
