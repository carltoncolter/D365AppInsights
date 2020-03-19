using System.Net.Http;

namespace JLattimer.D365AppInsights
{
    public static class HttpHelper
    {
        public static HttpClient GetHttpClient()
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Connection", "close");

            return client;
        }
    }
}