using GoogleSheetsToDocs.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GoogleSheetsToDocs.Services.Interfaces
{
    public interface IConvertService
    {
        Task ExtractDataFromSheet(SheetVM VM);
        Task CreateDoc(SheetVM vM);
    }
}
