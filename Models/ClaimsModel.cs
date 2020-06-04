using Newtonsoft.Json;
using System;

namespace GoogleExcel.Models
{
    public class ClaimsModel
    {
        /// <summary>
        /// The email address of the service account.
        /// </summary>
        [JsonProperty("iss")]
        public string Iss;

        /// <summary>
        /// https://www.googleapis.com/auth/spreadsheets.readonly
        /// https://www.googleapis.com/auth/spreadsheets
        /// https://www.googleapis.com/auth/drive.readonly
        /// https://www.googleapis.com/auth/drive.file
        /// https://www.googleapis.com/auth/drive
        /// </summary>
        [JsonProperty("scope")]
        public string Scope;

        /// <summary>
        /// A descriptor of the intended target of the assertion. When making an access token request this value is always https://www.googleapis.com/oauth2/v4/token.
        /// </summary>
        [JsonProperty("aud")]
        public string Aud;

        [JsonProperty("exp")]
        public int Exp;

        [JsonProperty("iat")]
        public int Iat;

        public ClaimsModel()
        {
            Aud = "https://www.googleapis.com/oauth2/v4/token";

            DateTime t1 = new DateTime(1970, 1, 1);
            Iat = (int)DateTime.UtcNow.Subtract(t1).TotalSeconds;
        }

        public ClaimsModel(string clientEmail, Scope scope, int exp = 3600)
            : this()
        {
            Iss = clientEmail;
            Scope = scope.Value;
            Exp = Iat + exp;
        }
    }

}
