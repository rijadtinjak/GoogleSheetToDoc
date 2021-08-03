using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;
using Google.Apis.Docs.v1.Data;

namespace GoogleSheetsToDocs.Services.Interfaces
{
    public interface IGoogleDocsApiService
    {
        Task<IList<Google.Apis.Docs.v1.Data.Response>> BatchUpdate(string documentId, BatchUpdateDocumentRequest body);
        Task<Document> Get(string documentId, string fields);
    }
}
