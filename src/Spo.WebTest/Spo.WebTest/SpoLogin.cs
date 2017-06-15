namespace Spo.WebTest
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.WebTesting;
    using Microsoft.VisualStudio.TestTools.WebTesting.Rules;
    using OpenQA.Selenium.Remote;
    using OpenQA.Selenium.Support.UI;
    using OpenQA.Selenium.PhantomJS;
    using System.Text.RegularExpressions;
    using PCLWebUtility;
    using OpenQA.Selenium;
    using System.Reflection;
    using System.IO;

    [DeploymentItem("phantomjs.exe")]
    public class SpoLogin : WebTest
    {
        
        private RemoteWebDriver driver;
        private WebDriverWait wait { get; set; }
        private TimeSpan timeout;

        public SpoLogin()
        {
            timeout = TimeSpan.FromSeconds(30);
        }

        public override IEnumerator<WebTestRequest> GetRequestEnumerator()
        {
            this.AddCommentToResult("Browse the homepage");
            // Open the homepage
            WebTestRequest request1 = new WebTestRequest(("https://"
                            + (this.Context["Tenant"].ToString() + ".sharepoint.com")));
            request1.FollowRedirects = false;
            ExtractText extractionRule1 = new ExtractText();
            extractionRule1.StartsWith = "href=\"";
            extractionRule1.EndsWith = "\">";
            extractionRule1.IgnoreCase = false;
            extractionRule1.UseRegularExpression = false;
            extractionRule1.Required = true;
            extractionRule1.ExtractRandomMatch = false;
            extractionRule1.Index = 0;
            extractionRule1.HtmlDecode = true;
            extractionRule1.SearchInHeaders = false;
            extractionRule1.ContextParameterName = "Authenticate302";
            request1.ExtractValues += new EventHandler<ExtractionEventArgs>(extractionRule1.Extract);
            yield return request1;
            request1 = null;

            WebTestRequest request2 = new WebTestRequest(this.Context["Authenticate302"].ToString());
            request2.FollowRedirects = false;
            request2.Headers.Add(new WebTestRequestHeader("referer", ("https://"
                                + (this.Context["Tenant"].ToString() + ".sharepoint.com"))));
            ExtractText extractionRule2 = new ExtractText();
            extractionRule2.StartsWith = "href=\"";
            extractionRule2.EndsWith = "\">";
            extractionRule2.IgnoreCase = false;
            extractionRule2.UseRegularExpression = false;
            extractionRule2.Required = true;
            extractionRule2.ExtractRandomMatch = false;
            extractionRule2.Index = 0;
            extractionRule2.HtmlDecode = true;
            extractionRule2.SearchInHeaders = false;
            extractionRule2.ContextParameterName = "AuthSpForms";
            request2.ExtractValues += new EventHandler<ExtractionEventArgs>(extractionRule2.Extract);
            yield return request2;
            request2 = null;

            this.AddCommentToResult("Authentication starts");

            //foreach (WebTestRequest r in IncludeWebTest("SpoLogin", false)) { yield return r; };
            driver = new PhantomJSDriver(AssemblyDirectory);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            var baseURl = $"https://{this.Context["Tenant"].ToString()}.sharepoint.com";
            string pattern = @"(?<=CONTEXT \= \').*(?=\';)";

            driver.Navigate().GoToUrl(baseURl);            
            var match = Regex.Match(driver.PageSource, pattern);

            var validatorURl = $"https://login.microsoftonline.com/common/userrealm?user={this.Context["UserName"].ToString()}&api-version=2.1&stsRequest={match.Value}&checkForMicrosoftAccount=true";

            driver.Navigate().GoToUrl(validatorURl);

            pattern = "(?<=AuthURL\\\"\\:\\\")[^\\\"]*(?=\\\")";
            match = Regex.Match(driver.PageSource, pattern);
            if (match.Success)
            {
                FederatedLogin(WebUtility.HtmlDecode(match.Value), this.Context["UserName"].ToString(), this.Context["UserPassword"].ToString());
            }
            else
            {
                AzureDomainLogin(baseURl, this.Context["UserName"].ToString(), this.Context["UserPassword"].ToString());
            }
            // Get Hidden
            var content = driver.PageSource;
            var document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(content);
            foreach (var hidden in document.DocumentNode.Descendants("input")
                            .Where(u => u.GetAttributeValue("type", "") == "hidden"))
            {
                var attributeName = hidden.GetAttributeValue("name", "");
                var attributeValue = hidden.GetAttributeValue("value", "");

                if (!String.IsNullOrEmpty(attributeValue))
                {
                    var hiddenContextName = $"$HIDDENCtx.{attributeName}";
                    this.Context[hiddenContextName] = attributeValue;
                }
            }

            driver.Quit();

            // Check if request digest is empty, need to post token to SharePoint
            if (this.Context["$HIDDENCtx.__REQUESTDIGEST"] == null)
            {
                WebTestRequest request8 = new WebTestRequest(("https://"
                    + (this.Context["Tenant"].ToString()
                    + (".sharepoint.com" + this.Context["AuthSpForms"].ToString()))));
                request8.Method = "POST";
                request8.ExpectedResponseUrl = ("https://"
                            + (this.Context["Tenant"].ToString() + ".sharepoint.com"));
                request8.Headers.Add(new WebTestRequestHeader("Referer", (this.Context["BaseMsoUrl"].ToString() + this.Context["LoginMsoPost"].ToString())));
                FormPostHttpBody request8Body = new FormPostHttpBody();
                request8Body.FormPostParameters.Add("code", this.Context["$HIDDENLogin.code"].ToString());
                request8Body.FormPostParameters.Add("id_token", this.Context["$HIDDENLogin.id_token"].ToString());
                request8Body.FormPostParameters.Add("state", this.Context["$HIDDENLogin.state"].ToString());
                request8Body.FormPostParameters.Add("session_state", this.Context["$HIDDENLogin.session_state"].ToString());
                request8.Body = request8Body;
                ExtractHiddenFields extractionRule9 = new ExtractHiddenFields();
                extractionRule9.Required = true;
                extractionRule9.HtmlDecode = true;
                extractionRule9.ContextParameterName = "Ctx";
                request8.ExtractValues += new EventHandler<ExtractionEventArgs>(extractionRule9.Extract);
                yield return request8;
                request8 = null;
            }



            this.AddCommentToResult("Authentication completed");
            
        }
         

        private void AzureDomainLogin(string baseURl, string userName, string password)
        {
            driver.Navigate().GoToUrl(baseURl);
            driver.FindElementById("cred_userid_inputtext").SendKeys(userName);
            driver.FindElementById("cred_password_inputtext").SendKeys(password);
            driver.FindElementById("cred_sign_in_button").Click();
            var wait = new WebDriverWait(driver, timeout);
            wait.Until(driver =>
            {
                bool isCompleted = (bool)((IJavaScriptExecutor)driver)
                    .ExecuteScript("return $('#redirect_dots_animation').css('visibility') == 'hidden'");

                return isCompleted;
            });

            driver.FindElementById("cred_sign_in_button").Click();
            wait.Until(driver =>
            {
                bool isCompleted = (bool)((IJavaScriptExecutor)driver)
                  .ExecuteScript("return $('#s4-ribbonrow').length != 0");

                bool isLoaded = (bool)((IJavaScriptExecutor)driver)
                  .ExecuteScript("return $('#O365_MeFlexPane_ButtonID').length != 0");

                return isCompleted && isLoaded;
            });
        }

        private string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        private void FederatedLogin(string loginURl, string userName, string password)
        {
            var listNameValue = new Dictionary<string, string>();

            driver.Navigate().GoToUrl(loginURl);
            var content = driver.PageSource;
            var document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(content);

            var postUrl = WebUtility.HtmlDecode(document.DocumentNode.Descendants("form")
                            .Where(u => u.GetAttributeValue("name", "") == "f1")
                            .SingleOrDefault().GetAttributeValue("action", ""));

            foreach (var hidden in document.DocumentNode.Descendants("input")
                            .Where(u => u.GetAttributeValue("type", "") == "hidden"))
            {
                var attributeName = hidden.GetAttributeValue("name", "");
                var attributeValue = hidden.GetAttributeValue("value", "");

                if (!String.IsNullOrEmpty(attributeValue))
                {
                    var hiddenContextName = $"$HIDDEN1.{attributeName}";
                    listNameValue[hiddenContextName] = attributeValue;
                }
            }


            foreach (var hidden in document.DocumentNode.Descendants("input")
                            .Where(u => u.GetAttributeValue("type", "") != "hidden"))
            {
                var attributeName = hidden.GetAttributeValue("name", "");
                var attributeValue = hidden.GetAttributeValue("value", "");

                if (!String.IsNullOrEmpty(attributeValue))
                {
                    var hiddenContextName = $"$INPUT1.{attributeName}";
                    listNameValue[hiddenContextName] = attributeValue;
                }
            }

            driver.FindElementByName("passwd").SendKeys(password);
            driver.FindElementById("idSIButton9").Click();
        }
    }
}
