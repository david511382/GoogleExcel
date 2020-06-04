using Newtonsoft.Json;

namespace GoogleExcel.Models
{
    public class GoogleJsonSecret
    {
        [JsonProperty("type")]
        public string Type;

        [JsonProperty("project_id")]
        public string ProjectID;

        [JsonProperty("private_key_id")]
        public string PrivateKeyID;

        [JsonProperty("private_key")]
        public string PrivateKey;

        [JsonProperty("client_email")]
        public string ClientEmail;

        [JsonProperty("client_id")]
        public string ClientID;

        [JsonProperty("auth_uri")]
        public string AuthUri;

        [JsonProperty("token_uri")]
        public string TokenUri;

        [JsonProperty("auth_provider_x509_cert_url")]
        public string AuthProviderX509CertUrl;

        [JsonProperty("client_x509_cert_url")]
        public string ClientX509CertUrl;

        public GoogleJsonSecret()
        {

        }
    }

}
