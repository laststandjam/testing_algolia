using System;
using Algolia.Search.Clients;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Collections;
using Azure.Security.KeyVault.Secrets;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Algolia_Product_Data_Integration
{
    public class Product
    {
        public string ObjectID { get; set; }
        public string product_id_suffixes { get; set; }
        public string product_description { get; set; }
        public string brand { get; set; }
        public string film { get; set; }
        public string adhesive { get; set; }
        public string color { get; set; }
        public string print_technology { get; set; }
        public string print_press { get; set; }
        public string market { get; set; }
        public string application { get; set; }
        public string converting_type { get; set; }
        public string master_width { get; set; }
        public string location { get; set; }
        public string prefinished_type { get; set; }
        public string discounts { get; set; }
        public string surface_container_type { get; set; }
        public string application_surface_treatment { get; set; }
        public string application_surface_texture { get; set; }
        public string application_surface_profile { get; set; }
        public string agency { get; set; }
        public string score { get; set; }
    }
    public class Function1
    {
        [FunctionName("Function1")]
        public void Run([TimerTrigger("0 3 * * * ")]TimerInfo myTimer, ILogger log)  
        {
            // This function runs once a day @ 3:00 AM - Cron Syntax in the Job Scheduler
            //0 3 * * *   - @ 3 AM  (0 0 3 * * *)
            //* * * * *   - Instantly for testing
            //0 0 * * * * - Every Hour
            //30 * * * *  - Every Hlaf Hour
            //0 */5 * * * * - Every 5 Min

            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            ArrayList objs = new ArrayList();
            
            List<Product> products = new List<Product>();

            //var cs = Environment.GetEnvironmentVariable("Website", EnvironmentVariableTarget.Process);
            using (var conn = new SqlConnection("Data Source=FLEXSQL.flexcon.com,1433; Initial Catalog=Website; Persist Security Info=True;User ID=website; Password =w3bsit3!"))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "algolia_GetDelimitedProductData_VBS";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Connection = conn;
                    conn.Open();
                    cmd.ExecuteNonQuery();

                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        Product product = new Product();

                        product.ObjectID = reader.GetInt32(0).ToString();
                        product.product_id_suffixes = reader.GetString(1).Replace("FLX", "") + ", " + reader.GetString(1).Replace("FLX", "").TrimStart(new Char[] { '0' });
                        product.product_description = reader.IsDBNull(2) ? "" : reader.GetString(2);
                        product.brand = reader.IsDBNull(3) ? "" : reader.GetString(3).TrimEnd();
                        product.film = reader.IsDBNull(4) ? "" : reader.GetString(4);
                        product.adhesive = reader.IsDBNull(5) ? "" : reader.GetString(5);
                        product.color = reader.IsDBNull(6) ? "" : reader.GetString(6);
                        product.print_technology = reader.IsDBNull(7) ? "" : reader.GetString(7);
                        product.print_press = reader.IsDBNull(8) ? "" : reader.GetString(8);
                        product.market = reader.IsDBNull(9) ? "" : reader.GetString(9);
                        product.application = reader.IsDBNull(10) ? "" : reader.GetString(10);
                        product.converting_type = reader.IsDBNull(11) ? "" : reader.GetString(11);
                        product.master_width = reader.GetDecimal(12).ToString();
                        product.location = reader.IsDBNull(13) ? "" : reader.GetString(13);
                        product.prefinished_type = reader.IsDBNull(14) ? "" : reader.GetString(14);
                        product.discounts = reader.IsDBNull(15) ? "" : reader.GetString(15);
                        product.surface_container_type = reader.IsDBNull(16) ? "" : reader.GetString(16);
                        product.application_surface_treatment = reader.IsDBNull(17) ? "" : reader.GetString(17);
                        product.application_surface_texture = reader.IsDBNull(18) ? "" : reader.GetString(18);
                        product.application_surface_profile = reader.IsDBNull(19) ? "" : reader.GetString(19);
                        product.agency = reader.IsDBNull(20) ? "" : reader.GetString(20);
                        product.score = reader.GetInt32(22).ToString();

                        products.Add(product);
                    }
                }
                catch (Exception ex)
                {
                    log.LogInformation(ex.Message);
                    //Send Error Email
                    SendOAuthMail("webdev@flexcon.com", "Algolia Product Data Import Error", ex.Message);
                }
            }

            // You need to add the value to Cinfiguration -> Applications settings in Azure. The local.settings.json will not be moved to production. 
            var kvUrl = Environment.GetEnvironmentVariable("AzureKeyVaultUrl");

            var secretsClient = new SecretClient(new Uri(kvUrl), new DefaultAzureCredential());
            var apiKey = secretsClient.GetSecret("algolia-apikey");
            var appId = secretsClient.GetSecret("algolia-appid-dev");
            var indexName = secretsClient.GetSecret("algolia-indexname-dev");

            // Connect to Aloglia 
            SearchClient client = new SearchClient(appId.Value.Value, apiKey.Value.Value);
            // Algolia Index
            SearchIndex index = client.InitIndex(indexName.Value.Value);
            // Clears all objects from your index and replaces them with a new set of objects.
            index.ReplaceAllObjects(products);

            log.LogInformation(products.Count.ToString() + " products were uploaded to Algolia");

            // Send Confirmation Email
            string content = products.Count.ToString() + " products were uploaded to Algolia at: " + DateTime.Now;
            SendOAuthMail("webdev@flexcon.com", "Algolia Product Data Import Completed", content);
        }

        // The secret will expire on expires 5/21/24
        public static async void SendOAuthMail(string toAddress, string subject, string content)
        {
            string fromAddress = "no-reply@flexcon.com";

            string tenantId = "dcbc7964-664f-42e8-a055-54d9fc9f4390";
            string clientId = "ab92f4a7-2e4e-4fc4-9a76-1714a650239a";
            string clientSecret = "4DU8Q~7ME9R_KrrUE~4sHpEicxz6460RlwPo9brD";

            ClientSecretCredential credential = new(tenantId, clientId, clientSecret);
            GraphServiceClient graphClient = new(credential);

            Microsoft.Graph.Models.Message message = new()
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = Microsoft.Graph.Models.BodyType.Text,
                    Content = content
                },
                ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = toAddress
                            }
                        }
                    }
            };

            Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody body = new()
            {
                Message = message,
                SaveToSentItems = false  // or true, as you want
            };

            try
            {
                await graphClient.Users[fromAddress]
                        .SendMail
                        .PostAsync(body);

                //Console.WriteLine("Email sent successfully!");
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }
    }
}
