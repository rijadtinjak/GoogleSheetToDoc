using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsToDocs.Services.Interfaces
{
    public interface IGoogleSheetsApiService
    {

        Task<IList<IList<object>>> GetSheetValuesRange(string SheetId, string range, ValueRenderOptionEnum? valueRenderOption = null);

        Task<Sheet> GetSheetRange(string SheetId, string range, bool includeGridData);
    }
}
