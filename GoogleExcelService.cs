using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleExcel.Models;
using HttpHelper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GoogleExcel
{
    public class GoogleExcelService
    {
        private struct AuthorizeResponseModel
        {
            [JsonProperty("access_token")]
            public string AccessToken;

            [JsonProperty("token_type")]
            public string TokenType;

            [JsonProperty("expires_in")]
            public int ExpiresIn;
        }

        private static async Task<string> authorizePost(string assertion)
        {
            Dictionary<string, string> form = new Dictionary<string, string>();
            form.Add("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer");
            form.Add("assertion", assertion);

            HttpHelper.Model.ResponseModel resp = await HttpRequest.New()
                .SetForm(form)
                .To(TOKEN_URL)
                .Post();
            return resp.Content;
        }

        private const string JWT_HEADER = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9";
        private const string TOKEN_URL = "https://www.googleapis.com/oauth2/v4/token";
        private SheetsService _service;
        private string _accessToken;
        private string _privateKey;
        private ClaimsModel _claims;

        public GoogleExcelService(string applicationName, string email, Scope scope, string privateKey, int expires = 3600)
             : this(applicationName, privateKey, new ClaimsModel(email, scope, expires))
        { }

        public GoogleExcelService(string applicationName, string privateKey, ClaimsModel claims)
        {
            Init(applicationName, privateKey, claims);
        }

        /// <summary>
        /// use google json file to init
        /// </summary>
        /// <param name="applicationName"></param>
        /// <param name="fileName"></param>
        /// <param name="scope"></param>
        /// <param name="expires"></param>
        public GoogleExcelService(string applicationName, string fileName, Scope scope, int expires = 3600)
        {
            GoogleJsonSecret fileData = FileUtil.JsonFile<GoogleJsonSecret>(fileName);

            string email = fileData.ClientEmail;
            Init(applicationName, fileData.PrivateKey, new ClaimsModel(email, scope, expires));
        }

        public void Init(string applicationName, string privateKey, ClaimsModel claims)
        {
            _privateKey = privateKey;
            SetClaims(claims);

            // Create Google Sheets API service.
            _service = new SheetsService(
                new BaseClientService.Initializer()
                {
                    ApplicationName = applicationName
                }
            );
        }

        public void SetClaims(ClaimsModel claims)
        {
            _claims = claims;
        }

        public async Task FormatUpdate(string spreadsheetId, RangePosition range, CellFormat userEnteredFormat)
        {
            //get sheet id by sheet name
            int sheetId = await getSheetID(spreadsheetId, range.TableName);

            //define cell color
            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();

            //create the update request for cells from the first row
            Request updateCellsRequest = new Request()
            {
                RepeatCell = new RepeatCellRequest()
                {
                    Range = range.GetGridRange(sheetId),
                    Cell = new CellData()
                    {
                        UserEnteredFormat = userEnteredFormat
                    },
                    Fields = "UserEnteredFormat(BackgroundColor,TextFormat)"
                }
            };
            bussr.Requests = new List<Request>();
            bussr.Requests.Add(updateCellsRequest);

            SpreadsheetsResource.BatchUpdateRequest bur = _service.Spreadsheets.BatchUpdate(bussr, spreadsheetId);

            await autoAuthorize(bur);
        }

        public async Task Update(string spreadsheetId, UpdateRange updateRange)
        {
            string rangeStr = updateRange.Range.GetStr();
            ValueRange targetRange = new ValueRange();
            targetRange.Values = updateRange.Data;
            targetRange.MajorDimension = "ROWS";

            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                    _service.Spreadsheets.Values.Update(targetRange, spreadsheetId, rangeStr);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            await autoAuthorize(request);
        }

        public async Task BatchUpdate(string spreadsheetId, IEnumerable<UpdateRange> updateDatas)
        {
            BatchUpdateValuesRequest body = new BatchUpdateValuesRequest();
            body.Data = updateDatas.Select(d => new ValueRange
            {
                Range = d.Range.GetStr(),
                Values = d.Data,
                MajorDimension = "ROWS"
            }).ToList();
            body.ValueInputOption = "USER_ENTERED";
            SpreadsheetsResource.ValuesResource.BatchUpdateRequest request =
                    _service.Spreadsheets.Values.BatchUpdate(body, spreadsheetId);

            await autoAuthorize(request);
        }

        public async Task<IList<IList<object>>> Get(string spreadsheetId, RangePosition range)
        {
            string rangeStr = range.GetStr();

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    _service.Spreadsheets.Values.Get(spreadsheetId, rangeStr);

            ValueRange response = await autoAuthorize(request);

            return response.Values;
        }

        private async Task authorize()
        {
            string jsonStr = JsonConvert.SerializeObject(_claims);
            string jwtClaimSet = jsonStr.Base64();

            string assertion = $"{JWT_HEADER}.{jwtClaimSet}";
            byte[] shaSignature = ConvertUtil.Sha256WithRSA(assertion, _privateKey);
            string jwtSignature = shaSignature.Base64();

            assertion += "." + jwtSignature;

            string responseStr = await authorizePost(assertion);
            AuthorizeResponseModel response = JsonConvert.DeserializeObject<AuthorizeResponseModel>(responseStr);
            _accessToken = response.AccessToken;
        }

        private async Task<int> getSheetID(string spreadsheetId, string tableName)
        {
            //get sheet id by sheet name
            SpreadsheetsResource.GetRequest requs = _service.Spreadsheets.Get(spreadsheetId);
            Spreadsheet spr = await autoAuthorize(requs);
            if (spr == null)
            {
                throw new Exception("找不到此ID的表");
            }

            Sheet sh = spr.Sheets.Where(s => s.Properties.Title == tableName).FirstOrDefault();
            return (int)sh.Properties.SheetId;
        }

        private bool isUnauthorize(Exception e)
        {
            return e.Message.Contains("[403]");
        }

        private async Task<responseModel> autoAuthorize<responseModel>(SheetsBaseServiceRequest<responseModel> reqes)
        {
            bool hasRetried = false;
            while (true)
            {
                try
                {
                    reqes.AccessToken = _accessToken;
                    return await reqes.ExecuteAsync();
                }
                catch (Exception e)
                {
                    if (!hasRetried && isUnauthorize(e))
                    {
                        await authorize();
                        hasRetried = true;
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
        }
    }
}
