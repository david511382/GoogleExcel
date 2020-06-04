# GoogleExcel

## 服務設定

1. 到[Google developers 服務](https://console.developers.google.com/)
2. 啟用api Google Sheets API
3. 建立 服務帳戶
4. 用服務帳戶建立金鑰
5. 在 Google Sheet 分享權限給剛剛建立的服務帳戶

## Use Case

### SheetID in Google Sheet Uri

https://docs.google.com/spreadsheets/d/<SheetID>/edit#gid=0

### Google JSON格式的私密金鑰

```json
{
  "type": "service_account",
  "project_id": "project_id",
  "private_key_id": "private_key_id",
  "private_key": "private_key",
  "client_email": "client_email",
  "client_id": "client_id",
  "auth_uri": "https://auth_uri",
  "token_uri": "https://token_uri",
  "auth_provider_x509_cert_url": "https://auth_provider_x509_cert_url",
  "client_x509_cert_url": "https://client_x509_cert_url"
}
```

### Code

```C#
using GoogleExcel;
using GoogleExcel.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

class TestClass
{
    static void Main(string[] args)
    {
        const string USERNAME = "client_email";
        const string PRIVATE_KEY = "private_key";
        const string APP_NAME = "AppName";
        const string SHEET_ID = "SheetID";
        const string TABLE_NAME = "TableName";
        Scope SCOPE = Scope.SPREADSHEETS;
        GoogleExcelService service = new GoogleExcelService(APP_NAME, USERNAME, SCOPE, PRIVATE_KEY);

        RangePosition range = new RangePosition(TABLE_NAME, new Position(1, 1), new Position(10, 10));
        IList<IList<object>> result = await service.Get(SHEET_ID, range);
    }
}
```