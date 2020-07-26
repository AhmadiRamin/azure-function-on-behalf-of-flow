using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Security.Claims;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Linq;
using System.Web.Http;
using Microsoft.Identity.Client;
using System.Collections.Generic;

namespace Dynamics_crm_function
{
    public static class GetAccounts
    {
        [FunctionName("GetAccounts")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ClaimsPrincipal claimsPrincipal)
       {
            log.LogInformation("C# HTTP trigger function processed a request.");

            try
            {
                var jwtToken = req.Headers.FirstOrDefault(x => x.Key == "Authorization").Value.FirstOrDefault();
                var resource = req.Query["resource"];
                var token = GetDynamicsCrmAccessToken(jwtToken.Replace("Bearer ",""),resource);
                
                using (HttpClient httpClient = new HttpClient())
                {
                    httpClient.BaseAddress = new Uri(resource);
                    httpClient.Timeout = new TimeSpan(0, 2, 0);
                    httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
                    httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
                    httpClient.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue("application/json"));
                    httpClient.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", token);

                    HttpResponseMessage retrieveResponse =
                                await httpClient.GetAsync("/api/data/v9.1/accounts?$select=accountid,name,emailaddress1&$top=5");
                    if (retrieveResponse.IsSuccessStatusCode)
                    {
                        var results = await retrieveResponse.Content.ReadAsStringAsync();
                        var jsonObject = JsonConvert.DeserializeObject(results);
                        return new OkObjectResult(jsonObject);
                    }
                    else
                        return new BadRequestErrorMessageResult("");
                }
            }
            catch(Exception ex)
            {
                return new BadRequestErrorMessageResult(ex.Message);
            }
            
        }
        private static string GetDynamicsCrmAccessToken(string jwtToken, string crmUri)
        {
            var crmResourceUri = $"{crmUri}/.default";
            var clientId = Environment.GetEnvironmentVariable("CLIENTID");
            var clientSecret = Environment.GetEnvironmentVariable("CLIENTSECRET");
            var tenantId = Environment.GetEnvironmentVariable("TENANTID");
            List<string> scopes = new List<string>();
            scopes.Add(crmResourceUri);
            var app = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithClientSecret(clientSecret)
                .WithTenantId(tenantId)                
                .Build();
            var userAssertion = new UserAssertion(jwtToken);
            var result = app.AcquireTokenOnBehalfOf(scopes, userAssertion).ExecuteAsync().GetAwaiter().GetResult();
            return result.AccessToken;

        }
    }
}
