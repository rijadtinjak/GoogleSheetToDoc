using Google.Apis.Docs.v1.Data;
using GoogleSheetsToDocs.Models;
using GoogleSheetsToDocs.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static GoogleSheetsToDocs.Models.Claim;

namespace GoogleSheetsToDocs.Services.Implementations
{
    public class ConvertService : IConvertService
    {
        private readonly IGoogleSheetsApiService googleSheetsApiService;
        private readonly IGoogleDocsApiService googleDocsApiService;

        public ConvertService(IGoogleSheetsApiService googleSheetsApiService, IGoogleDocsApiService googleDocsApiService)
        {
            this.googleSheetsApiService = googleSheetsApiService;
            this.googleDocsApiService = googleDocsApiService;
        }

        public async Task ExtractDataFromSheet(SheetVM VM)
        {
            var coverPageTask = GetCoverPage(VM.SheetId);
            var abstractTask = GetAbstract(VM.SheetId);
            var classificationCodesTask = GetClassificationCodes(VM.SheetId);
            var claimsTask = GetClaims(VM.SheetId);
            var applicationEventsTask = GetApplicationEvents(VM.SheetId);

            await Task.WhenAny(coverPageTask, abstractTask, classificationCodesTask, claimsTask, applicationEventsTask);

            VM.CoverPage = coverPageTask.Result;
            VM.Abstract = abstractTask.Result;
            VM.ClassificationCodes = classificationCodesTask.Result;
            VM.ApplicationEvents = applicationEventsTask.Result;
            VM.Claims = claimsTask.Result;
        }

        private async Task<CoverPage> GetCoverPage(string SheetId)
        {
            var CoverPage = new CoverPage()
            {
                PreparedBy = "Agent Name"
            };

            var matrix = await googleSheetsApiService.GetSheetValuesRange(SheetId, "Cover Page!B1:B3");

            CoverPage.Title = matrix[0][0].ToString();
            CoverPage.PatentNo = matrix[1][0].ToString();
            CoverPage.Inventor = matrix[2][0].ToString();

            return CoverPage;
        }


        private async Task<string> GetAbstract(string SheetId)
        {
            var matrix = await googleSheetsApiService.GetSheetValuesRange(SheetId, "Abstract!B5");

            return matrix[0][0].ToString();
        }

        private async Task<List<ClassificationCode>> GetClassificationCodes(string SheetId)
        {
            var List = new List<ClassificationCode>();

            var matrix = await googleSheetsApiService.GetSheetValuesRange(SheetId, "Classification Codes!B4:C");

            if (matrix != null)
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    List.Add(new ClassificationCode
                    {
                        USC_Code = matrix[i][0].ToString(),
                        Definition = matrix[i][1].ToString()
                    });
                }
            }

            return List;
        }

        private async Task<List<Claim>> GetClaims(string SheetId)
        {
            var List = new List<Claim>();

            var matrix = await googleSheetsApiService.GetSheetValuesRange(SheetId, "Claims!B4:E");

            if (matrix != null)
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    for (int j = 0; j < matrix[i].Count; j++)
                    {
                        var indent = matrix[i][j].ToString();

                        if (!string.IsNullOrWhiteSpace(indent))
                        {
                            var level = (IndentLevel)j;
                            var text = matrix[i][j + 1].ToString();

                            if (level == IndentLevel.L1)
                            {
                                List.Add(new Claim
                                {
                                    Level = level,
                                    Content = text
                                });
                            }
                            else if (level == IndentLevel.L2)
                            {
                                var indentInfo = GetListIndent(indent);
                                List[indentInfo.L1Parent - 1].Children.Add(new Claim
                                {
                                    Level = level,
                                    Content = text
                                });
                            }
                            else if (level == IndentLevel.L3)
                            {
                                var indentInfo = GetListIndent(indent);
                                List[indentInfo.L1Parent - 1].Children[indentInfo.L2Parent - 1].Children.Add(new Claim
                                {
                                    Level = level,
                                    Content = text
                                });
                            }


                            break;
                        }
                    }
                }
            }

            return List;
        }

        static ListIndentInfo GetListIndent(string indent)
        {
            var info = new ListIndentInfo();

            var splitStr = indent.Split(".");

            if (splitStr.Length > 1)
            {
                info.L2Parent = int.Parse(splitStr[1]);
            }

            if (splitStr.Length > 0)
            {
                info.L1Parent = int.Parse(splitStr[0]);
            }

            return info;
        }

        private async Task<List<ApplicationEvent>> GetApplicationEvents(string SheetId)
        {
            var List = new List<ApplicationEvent>();

            var sheetRange = await googleSheetsApiService.GetSheetRange(SheetId, "Application Events!A5:B", true);

            var sheetData = sheetRange.Data[0];
            foreach (var row in sheetData.RowData)
            {
                if (row.Values.Count != 2)
                    continue;

                var date = row.Values[0];
                var status = row.Values[1];

                List.Add(new ApplicationEvent
                {
                    Date = DateTime.Parse(date.FormattedValue),
                    Status = status.FormattedValue,
                    Hyperlink = status.Hyperlink
                });
            }

            return List;
        }

        public async Task CreateDoc(SheetVM VM)
        {
            var DocID = VM.DocId;

            await ClearDocumentBody(DocID);

            var Requests = new List<Request>();

            int index = 1, offset = 0;

            AddInlineImage(Requests, "https://drive.google.com/uc?export=view&id=1oVrLElJ_q9ErqssSIgxukpAe80UwwWPC", ref index);

            AddLineBreak(Requests, ref index, 3);

            AddHeading(VM.CoverPage.Title, Requests, ref index);

            offset += 1;
            int NewlineCount = 12;

            AddLineBreak(Requests, ref index, NewlineCount);
            UpdateParagraphStyle(Requests, new ParagraphStyle
            {
                NamedStyleType = "NORMAL_TEXT"
            }, index - NewlineCount + offset, index + offset);

            var Font14 = new TextStyle
            {
                FontSize = new Dimension
                {
                    Magnitude = 14,
                    Unit = "PT"
                }
            };
            var Font14Bold = new TextStyle
            {
                FontSize = new Dimension
                {
                    Magnitude = 14,
                    Unit = "PT"
                },
                Bold = true
            };
            var Font10 = new TextStyle
            {
                FontSize = new Dimension
                {
                    Magnitude = 10,
                    Unit = "PT"
                }
            };
            var Font10Bold = new TextStyle
            {
                FontSize = new Dimension
                {
                    Magnitude = 10,
                    Unit = "PT"
                },
                Bold = true
            };

            AddText(Requests, "Prepared by: ", ref index, Font14);
            AddText(Requests, VM.CoverPage.PreparedBy, ref index, Font14Bold, true);

            AddText(Requests, "Inventor: ", ref index, Font14);
            AddText(Requests, VM.CoverPage.Inventor, ref index, Font14Bold, true);

            AddText(Requests, "Patent No: ", ref index, Font14);
            AddText(Requests, VM.CoverPage.PatentNo, ref index, Font14Bold, true);

            AddPageBreak(Requests, ref index);

            var text = "Abstract";

            AddText(Requests, text, ref index, new TextStyle { Bold = true }, true);
            UpdateParagraphStyle(Requests, new ParagraphStyle
            {
                NamedStyleType = "HEADING_2"
            }, index - text.Length - 1 + offset, index - 1 + offset);


            Requests.Add(new Request
            {
                InsertTable = new InsertTableRequest
                {
                    Columns = 1,
                    Rows = 2,
                    Location = new Location
                    {
                        Index = index++
                    }
                },

            });

            var tableIdx = index;
            UpdateTableHeadingBackground(Requests, tableIdx, 0, 0, 159, 197, 232);

            var table0Row0CellContent = tableIdx + 3 * 1;
            var table0HeadingText = "Abstract";
            var table0Row1CellContent = tableIdx + 3 * 2 + table0HeadingText.Length;

            AddText(Requests, table0HeadingText, ref table0Row0CellContent, new TextStyle
            {
                Bold = true
            }, false, false);

            AddText(Requests, VM.Abstract, ref table0Row1CellContent, new TextStyle
            {
                FontSize = new Dimension
                {
                    Magnitude = 10,
                    Unit = "PT"
                },
                WeightedFontFamily = new WeightedFontFamily
                {
                    FontFamily = "Roboto"
                }
            }, false, false);


            await googleDocsApiService.BatchUpdate(DocID, new BatchUpdateDocumentRequest
            {
                Requests = Requests
            });


            var Document = await googleDocsApiService.Get(DocID, "body.content(startIndex,endIndex)");
            index = Document.Body.Content[Document.Body.Content.Count - 1].EndIndex.Value - 1;
            Requests.Clear();

            AddPageBreak(Requests, ref index);

            text = "Relevant Classification Codes";

            AddText(Requests, text, ref index, new TextStyle { Bold = true }, true);
            UpdateParagraphStyle(Requests, new ParagraphStyle
            {
                NamedStyleType = "HEADING_2"
            }, index - text.Length - 1 + offset, index - 1 + offset);


            Requests.Add(new Request
            {
                InsertTable = new InsertTableRequest
                {
                    Columns = 2,
                    Rows = VM.ClassificationCodes.Count + 1,
                    Location = new Location
                    {
                        Index = index++
                    }
                },

            });

            tableIdx = index;
            var table1Row0Idx = tableIdx + 1;
            for (int i = 0; i < 2; i++)
            {
                UpdateTableHeadingBackground(Requests, tableIdx, i, 0, 207, 226, 243);

            }

            var table1HeadingText0 = "USC Codes";
            var table1HeadingText1 = "Definitions";

            var table1Row0Col0CellContent = table1Row0Idx + 2 * 1;
            var table1Row0Col1CellContent = table1Row0Idx + 2 * 2 + table1HeadingText0.Length;

            AddText(Requests, table1HeadingText0, ref table1Row0Col0CellContent, Font10Bold, false, false);


            AddText(Requests, table1HeadingText1, ref table1Row0Col1CellContent, Font10Bold, false, false);

            var currentIdx = table1Row0Col1CellContent + table1HeadingText1.Length;

            for (int i = 0; i < VM.ClassificationCodes.Count; i++)
            {
                currentIdx += 3;

                AddText(Requests, VM.ClassificationCodes[i].USC_Code, ref currentIdx, Font10, false, false);

                currentIdx += VM.ClassificationCodes[i].USC_Code.Length;
                currentIdx += 2;

                AddText(Requests, VM.ClassificationCodes[i].Definition, ref currentIdx, Font10, false, false);

                currentIdx += VM.ClassificationCodes[i].Definition.Length;


            }

            await googleDocsApiService.BatchUpdate(DocID, new BatchUpdateDocumentRequest
            {
                Requests = Requests
            });

            Document = await googleDocsApiService.Get(DocID, "body.content(startIndex,endIndex)");
            index = Document.Body.Content[Document.Body.Content.Count - 1].EndIndex.Value - 1;
            Requests.Clear();

            AddPageBreak(Requests, ref index);

            text = "Claims";

            AddText(Requests, text, ref index, new TextStyle { Bold = true }, true);
            UpdateParagraphStyle(Requests, new ParagraphStyle
            {
                NamedStyleType = "HEADING_2"
            }, index - text.Length - 1 + offset, index - 1 + offset);

            var listStartIndex = index;

            AddClaims(Requests, VM.Claims, ref index, Font10Bold);

            Requests.Add(new Request
            {
                CreateParagraphBullets = new CreateParagraphBulletsRequest
                {
                    BulletPreset = "NUMBERED_DECIMAL_NESTED",
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = listStartIndex,
                        EndIndex = index
                    }
                }
            });

            await googleDocsApiService.BatchUpdate(DocID, new BatchUpdateDocumentRequest
            {
                Requests = Requests
            });

            Document = await googleDocsApiService.Get(DocID, "body.content(startIndex,endIndex)");
            index = Document.Body.Content[Document.Body.Content.Count - 1].EndIndex.Value - 1;
            Requests.Clear();

            text = "Application Events";

            AddText(Requests, text, ref index, new TextStyle { Bold = true }, true);
            UpdateParagraphStyle(Requests, new ParagraphStyle
            {
                NamedStyleType = "HEADING_2"
            }, index - text.Length - 1 + offset, index - 1 + offset);


            Requests.Add(new Request
            {
                InsertTable = new InsertTableRequest
                {
                    Columns = 2,
                    Rows = VM.ApplicationEvents.Count + 1,
                    Location = new Location
                    {
                        Index = index++
                    }
                },

            });

            tableIdx = index;
            var table2Row0Idx = tableIdx + 1;
            for (int i = 0; i < 2; i++)
            {
                UpdateTableHeadingBackground(Requests, tableIdx, i, 0, 159, 197, 232);
            }

            var table2HeadingText0 = "Date";
            var table2HeadingText1 = "Status";

            var table2Row0Col0CellContent = table2Row0Idx + 2 * 1;
            var table2Row0Col1CellContent = table2Row0Idx + 2 * 2 + table2HeadingText0.Length;

            AddText(Requests, table2HeadingText0, ref table2Row0Col0CellContent, Font10Bold, false, false);

            AddText(Requests, table2HeadingText1, ref table2Row0Col1CellContent, Font10Bold, false, false);

            currentIdx = table2Row0Col1CellContent + table2HeadingText1.Length;

            var GreyColor = new RgbColor
            {
                Red = 95 / (float)255,
                Green = 99 / (float)255,
                Blue = 104 / (float)255
            };
            var HyperlinkColor = new RgbColor
            {
                Red = 17 / (float)255,
                Green = 85 / (float)255,
                Blue = 204 / (float)255
            };

            for (int i = 0; i < VM.ApplicationEvents.Count; i++)
            {
                currentIdx += 3;

                var FormattedDate = VM.ApplicationEvents[i].Date.ToString("yyyy-MM-dd");
                AddText(Requests, FormattedDate, ref currentIdx, Font10, false, false);

                currentIdx += FormattedDate.Length;
                currentIdx += 2;

                bool HasHyperlink = VM.ApplicationEvents[i].Hyperlink != null;

                AddText(Requests, VM.ApplicationEvents[i].Status, ref currentIdx, new TextStyle
                {
                    Link = HasHyperlink ? new Link
                    {
                        Url = VM.ApplicationEvents[i].Hyperlink,
                    } : null,
                    FontSize = new Dimension
                    {
                        Magnitude = 10,
                        Unit = "PT"
                    },
                    ForegroundColor = new OptionalColor
                    {
                        Color = new Color
                        {
                            RgbColor = HasHyperlink ? HyperlinkColor : GreyColor
                        }
                    },
                    Underline = HasHyperlink
                }, false, false);

                currentIdx += VM.ApplicationEvents[i].Status.Length;
            }

            await googleDocsApiService.BatchUpdate(DocID, new BatchUpdateDocumentRequest
            {
                Requests = Requests
            });

        }

        private void AddClaims(List<Request> requests, List<Claim> claims, ref int index, TextStyle textStyle = null)
        {
            foreach (var claim in claims)
            {
                AddText(requests, new string('\t', (int)claim.Level) + claim.Content, ref index, textStyle, true);
                if(claim.Children != null)
                {
                    AddClaims(requests, claim.Children, ref index, textStyle);
                }
            }
        }

        private static void UpdateTableHeadingBackground(List<Request> Requests, int tableIdx, int column, int row, int r, int g, int b)
        {
            Requests.Add(new Request
            {
                UpdateTableCellStyle = new UpdateTableCellStyleRequest
                {
                    Fields = "backgroundColor",
                    TableRange = new TableRange
                    {
                        TableCellLocation = new TableCellLocation
                        {
                            TableStartLocation = new Location
                            {
                                Index = tableIdx
                            },
                            ColumnIndex = column,
                            RowIndex = row
                        },
                        ColumnSpan = 1,
                        RowSpan = 1
                    },
                    TableCellStyle = new TableCellStyle
                    {
                        BackgroundColor = new OptionalColor()
                        {
                            Color = new Color()
                            {
                                RgbColor = new RgbColor()
                                {
                                    Red = r / (float)255,
                                    Green = g / (float)255,
                                    Blue = b / (float)255
                                }
                            }
                        }
                    }
                }
            });
        }

        private static void AddInlineImage(List<Request> Requests, string ImageUri, ref int index)
        {
            Requests.Add(new Request
            {
                InsertInlineImage = new InsertInlineImageRequest
                {
                    Uri = ImageUri,
                    Location = new Location
                    {
                        Index = index++
                    }
                }
            });
        }

        private static void AddPageBreak(List<Request> Requests, ref int index)
        {
            Requests.Add(new Request
            {
                InsertPageBreak = new InsertPageBreakRequest
                {
                    Location = new Location
                    {
                        Index = index++
                    }
                }
            });
        }

        private void UpdateParagraphStyle(List<Request> Requests, ParagraphStyle ParagraphStyle, int StartIndex, int EndIndex)
        {
            Requests.Add(new Request
            {
                UpdateParagraphStyle = new UpdateParagraphStyleRequest
                {
                    ParagraphStyle = ParagraphStyle,
                    Fields = "*",
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = StartIndex,
                        EndIndex = EndIndex
                    }
                }
            });
        }

        private void UpdateTextStyle(List<Request> Requests, TextStyle TextStyle, int StartIndex, int EndIndex)
        {
            Requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    TextStyle = TextStyle,
                    Fields = "*",
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = StartIndex,
                        EndIndex = EndIndex
                    }
                }
            });
        }

        private static void AddText(List<Request> Requests, string text, ref int index, TextStyle TextStyle = null, bool addNewLine = false, bool updateIndex = true)
        {
            if (addNewLine)
                text += "\n";

            Requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = text,
                    Location = new Location
                    {
                        Index = (updateIndex ? index++ : index)
                    }
                }
            });
            if (updateIndex)
                index += text.Length - 1;

            if (TextStyle != null)
            {
                var Range = new Google.Apis.Docs.v1.Data.Range();
                if (updateIndex)
                {
                    Range.StartIndex = index - text.Length;
                    Range.EndIndex = index;
                }
                else
                {
                    Range.StartIndex = index;
                    Range.EndIndex = index + text.Length;
                }

                Requests.Add(new Request
                {
                    UpdateTextStyle = new UpdateTextStyleRequest
                    {
                        Fields = "*",
                        TextStyle = TextStyle,
                        Range = Range
                    }
                });
            }
        }

        private static void AddHeading(string headingText, List<Request> Requests, ref int index)
        {
            Requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = headingText,
                    Location = new Location
                    {
                        Index = index
                    }
                }
            });

            Requests.Add(new Request
            {
                UpdateParagraphStyle = new UpdateParagraphStyleRequest
                {
                    ParagraphStyle = new ParagraphStyle
                    {
                        NamedStyleType = "TITLE",
                        Alignment = "CENTER",
                    },
                    Fields = "*",
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = index,
                        EndIndex = index + headingText.Length
                    }
                }
            });

            index += headingText.Length;
        }

        private static void AddLineBreak(List<Request> Requests, ref int index, int count = 1)
        {
            for (int i = 0; i < count; i++)
            {
                Requests.Add(new Request
                {
                    InsertText = new InsertTextRequest
                    {
                        Text = "\n",
                        Location = new Location
                        {
                            Index = index++
                        }
                    }
                });
            }

        }

        private async Task ClearDocumentBody(string DocID)
        {
            var Document = await googleDocsApiService.Get(DocID, "body.content(startIndex,endIndex)");

            if (Document.Body.Content.Count == 2)
                return;

            var DeleteRequest = new List<Request>()
            {
                new Request
                {
                    DeleteContentRange = new DeleteContentRangeRequest
                    {
                        Range = new Google.Apis.Docs.v1.Data.Range
                        {
                            StartIndex = 1,
                            EndIndex = Document.Body.Content[Document.Body.Content.Count - 1].EndIndex - 1
                        }
                    }
                }
            };

            var DeleteResponse = await this.googleDocsApiService.BatchUpdate(DocID, new Google.Apis.Docs.v1.Data.BatchUpdateDocumentRequest
            {
                Requests = DeleteRequest
            });
        }
    }

    public class ListIndentInfo
    {
        public int L1Parent { get; set; }
        public int L2Parent { get; set; }
    }
}
