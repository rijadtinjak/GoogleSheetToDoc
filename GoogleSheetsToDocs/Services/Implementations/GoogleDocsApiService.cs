using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using GoogleSheetsToDocs.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsToDocs.Services.Implementations
{
    public class GoogleDocsApiService : IGoogleDocsApiService
    {
        private readonly DocsService googleDocsService;
        private readonly IGoogleAuthService googleAuthService;

        public GoogleDocsApiService(IGoogleAuthService googleAuthService)
        {
            this.googleAuthService = googleAuthService;

            UserCredential credential = this.googleAuthService.GetUserCredential();

            googleDocsService = new DocsService(new BaseClientService.Initializer
            {
                ApplicationName = "GoogleSheetToDoc app",
                HttpClientInitializer = credential
            });
        }

        public async Task<IList<Google.Apis.Docs.v1.Data.Response>> BatchUpdate(string documentId, BatchUpdateDocumentRequest body)
        {
            var request = this.googleDocsService.Documents.BatchUpdate(body, documentId);

            BatchUpdateDocumentResponse response = await request.ExecuteAsync();
            return response.Replies;
        }

        public async Task<Document> Get(string documentId, string fields)
        {
            var request = this.googleDocsService.Documents.Get(documentId);

            if (fields != null)
                request.Fields = fields;

            return await request.ExecuteAsync();
        }
    }
}
