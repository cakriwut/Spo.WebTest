namespace Spo.WebTest
{
    using System;
    using System.Collections.Generic;
    using Microsoft.VisualStudio.TestTools.WebTesting;
    using Microsoft.VisualStudio.TestTools.WebTesting.Rules;

    [IncludeCodedWebTest("Spo.WebTest.SpoLogin", @"SpoLogin.cs")]
    public class ExampleWebTest : WebTest
    {
        

        public ExampleWebTest()
        {
            this.Context.Add("Tenant", "");  
            this.Context.Add("UserName", "");
            this.Context.Add("UserPassword", "");
            this.PreAuthenticate = true;
            this.Proxy = "default";
        }

        public override IEnumerator<WebTestRequest> GetRequestEnumerator()
        {
            // Initialize validation rules that apply to all requests in the WebTest
            if ((this.Context.ValidationLevel >= Microsoft.VisualStudio.TestTools.WebTesting.ValidationLevel.Low))
            {
                ValidateResponseUrl validationRule1 = new ValidateResponseUrl();
                this.ValidateResponse += new EventHandler<ValidationEventArgs>(validationRule1.Validate);
            }
            if ((this.Context.ValidationLevel >= Microsoft.VisualStudio.TestTools.WebTesting.ValidationLevel.Low))
            {
                ValidationRuleResponseTimeGoal validationRule2 = new ValidationRuleResponseTimeGoal();
                validationRule2.Tolerance = 0D;
                this.ValidateResponseOnPageComplete += new EventHandler<ValidationEventArgs>(validationRule2.Validate);
            }

            
            //
            foreach (WebTestRequest r in IncludeWebTest("SpoLogin", false)) { yield return r; };

            WebTestRequest request1 = new WebTestRequest($"{this.Context["Tenant"].ToString()}/_vti_bin/client.svc/ProcessQuery");
            request1.Method = "POST";
            request1.Headers.Add(new WebTestRequestHeader("X-Requested-With", "XMLHttpRequest"));
            request1.Headers.Add(new WebTestRequestHeader("X-RequestDigest", this.Context["$HIDDENCtx.__REQUESTDIGEST"].ToString()));
            StringHttpBody request1Body = new StringHttpBody();
            request1Body.ContentType = "text/xml";
            request1Body.InsertByteOrderMark = false;
            request1Body.BodyString = "<Request xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\" SchemaV" +
                "ersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\"Javascript Library\"" +
                "><Actions><ObjectPath Id=\"1\" ObjectPathId=\"0\" /><ObjectPath Id=\"3\" ObjectPathId=" +
                "\"2\" /><ObjectPath Id=\"5\" ObjectPathId=\"4\" /><ObjectPath Id=\"7\" ObjectPathId=\"6\" " +
                "/><ObjectPath Id=\"9\" ObjectPathId=\"8\" /><ExceptionHandlingScopeSimple Id=\"10\"><M" +
                "ethod Name=\"ValidateUpdateListItem\" Id=\"12\" ObjectPathId=\"8\"><Parameters><Parame" +
                "ter Type=\"Array\"><Object TypeId=\"{03745a5a-2400-440e-92d6-dad4afee30a6}\"><Proper" +
                "ty Name=\"ErrorMessage\" Type=\"Null\" /><Property Name=\"FieldName\" Type=\"String\">Ti" +
                "tle</Property><Property Name=\"FieldValue\" Type=\"String\">Test Hello</Property><Pr" +
                "operty Name=\"HasException\" Type=\"Boolean\">false</Property></Object><Object TypeI" +
                "d=\"{03745a5a-2400-440e-92d6-dad4afee30a6}\"><Property Name=\"ErrorMessage\" Type=\"N" +
                "ull\" /><Property Name=\"FieldName\" Type=\"String\">StartDate</Property><Property Na" +
                "me=\"FieldValue\" Type=\"String\">6/14/2017</Property><Property Name=\"HasException\" " +
                "Type=\"Boolean\">false</Property></Object><Object TypeId=\"{03745a5a-2400-440e-92d6" +
                "-dad4afee30a6}\"><Property Name=\"ErrorMessage\" Type=\"Null\" /><Property Name=\"Fiel" +
                "dName\" Type=\"String\">DueDate</Property><Property Name=\"FieldValue\" Type=\"String\"" +
                "></Property><Property Name=\"HasException\" Type=\"Boolean\">false</Property></Objec" +
                "t><Object TypeId=\"{03745a5a-2400-440e-92d6-dad4afee30a6}\"><Property Name=\"ErrorM" +
                "essage\" Type=\"Null\" /><Property Name=\"FieldName\" Type=\"String\">PercentComplete</" +
                "Property><Property Name=\"FieldValue\" Type=\"String\"></Property><Property Name=\"Ha" +
                "sException\" Type=\"Boolean\">false</Property></Object><Object TypeId=\"{03745a5a-24" +
                "00-440e-92d6-dad4afee30a6}\"><Property Name=\"ErrorMessage\" Type=\"Null\" /><Propert" +
                "y Name=\"FieldName\" Type=\"String\">Body</Property><Property Name=\"FieldValue\" Type" +
                "=\"String\">&lt;div&gt;&lt;/div&gt;</Property><Property Name=\"HasException\" Type=\"" +
                "Boolean\">false</Property></Object><Object TypeId=\"{03745a5a-2400-440e-92d6-dad4a" +
                "fee30a6}\"><Property Name=\"ErrorMessage\" Type=\"Null\" /><Property Name=\"FieldName\"" +
                " Type=\"String\">Predecessors</Property><Property Name=\"FieldValue\" Type=\"String\">" +
                "</Property><Property Name=\"HasException\" Type=\"Boolean\">false</Property></Object" +
                "><Object TypeId=\"{03745a5a-2400-440e-92d6-dad4afee30a6}\"><Property Name=\"ErrorMe" +
                "ssage\" Type=\"Null\" /><Property Name=\"FieldName\" Type=\"String\">Priority</Property" +
                "><Property Name=\"FieldValue\" Type=\"String\">(2) Normal</Property><Property Name=\"" +
                "HasException\" Type=\"Boolean\">false</Property></Object><Object TypeId=\"{03745a5a-" +
                "2400-440e-92d6-dad4afee30a6}\"><Property Name=\"ErrorMessage\" Type=\"Null\" /><Prope" +
                "rty Name=\"FieldName\" Type=\"String\">Status</Property><Property Name=\"FieldValue\" " +
                "Type=\"String\">Not Started</Property><Property Name=\"HasException\" Type=\"Boolean\"" +
                ">false</Property></Object><Object TypeId=\"{03745a5a-2400-440e-92d6-dad4afee30a6}" +
                "\"><Property Name=\"ErrorMessage\" Type=\"Null\" /><Property Name=\"FieldName\" Type=\"S" +
                "tring\">AssignedTo</Property><Property Name=\"FieldValue\" Type=\"String\"></Property" +
                "><Property Name=\"HasException\" Type=\"Boolean\">false</Property></Object><Object T" +
                "ypeId=\"{03745a5a-2400-440e-92d6-dad4afee30a6}\"><Property Name=\"ErrorMessage\" Typ" +
                "e=\"Null\" /><Property Name=\"FieldName\" Type=\"String\">ContentType</Property><Prope" +
                "rty Name=\"FieldValue\" Type=\"String\">Task</Property><Property Name=\"HasException\"" +
                " Type=\"Boolean\">false</Property></Object><Object TypeId=\"{03745a5a-2400-440e-92d" +
                "6-dad4afee30a6}\"><Property Name=\"ErrorMessage\" Type=\"Null\" /><Property Name=\"Fie" +
                "ldName\" Type=\"String\">ContentTypeId</Property><Property Name=\"FieldValue\" Type=\"" +
                "String\">0x010800DB18DF543E2EED45AD5C1065CD17430A</Property><Property Name=\"HasEx" +
                "ception\" Type=\"Boolean\">false</Property></Object></Parameter><Parameter Type=\"Bo" +
                "olean\">false</Parameter><Parameter Type=\"Null\" /></Parameters></Method></Excepti" +
                "onHandlingScopeSimple></Actions><ObjectPaths><StaticProperty Id=\"0\" TypeId=\"{374" +
                "7adcd-a3c3-41b9-bfab-4a64dd2f1e0a}\" Name=\"Current\" /><Property Id=\"2\" ParentId=\"" +
                "0\" Name=\"Web\" /><Property Id=\"4\" ParentId=\"2\" Name=\"Lists\" /><Method Id=\"6\" Pare" +
                "ntId=\"4\" Name=\"GetById\"><Parameters><Parameter Type=\"String\">40ec8848-453b-47eb-" +
                "95ab-7bc0cc4ca9b0</Parameter></Parameters></Method><Method Id=\"8\" ParentId=\"6\" N" +
                "ame=\"AddItem\"><Parameters><Parameter TypeId=\"{54cdbee5-0897-44ac-829f-411557fa11" +
                "be}\"><Property Name=\"FolderUrl\" Type=\"String\">/Lists/Tasks</Property><Property N" +
                "ame=\"LeafName\" Type=\"Null\" /><Property Name=\"UnderlyingObjectType\" Type=\"Number\"" +
                ">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>";
            request1.Body = request1Body;
            yield return request1;
            request1 = null;
        }

        
    }
}
