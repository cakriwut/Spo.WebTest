## SharePoint Online WebTest Starter Kit

### Background
When I start to use Visual Studio Web Performance Test to do automate test to my SharePoint tenant,
I face a problem. The VsWebTest capture Guids and unable to detect dynamic content properly. 
The main issue especially happens during Authentication process. 
This problem makes my test project is not reusable for other tenant, and I have to records the
steps and recreate the steps from begining everytime I need to create a test.

### Solution
The solution is to decode Authentication process and store in common SpoLogin.webtest. 
Then I can include the SpoLogin.webtest in the begining of every future test for SharePoint.
The SpoLogin.webtest will handle authentication process, and provide authentication cookie 
for the rest of the test.

### Known Issue
1. Login using Microsoft Live is now supported through Coded Web Performance Test

### Usage
1. Declarative SpoLogin.WebTest (Only support AAD Login) <br />
Create new webtest using Web Performane test recording, after finishing the recording, you can replace the 
authentication part using SpoLogin.WebTest. Right click on the new web test, and select "Add Call to a Web Test".

2. Coded Web Test (Support AAD Login and Federated Login such as Windows Live) <br />
Create a new webtest class that inherits from <code>Microsoft.VisualStudio.TestTools.WebTesting.WebTest</code>. 
Decorate the new class with <code>[IncludeCodedWebTest("Spo.WebTest.SpoLogin", @"SpoLogin.cs")]</code> 
Implement abstract method <code>GetRequestEnumerator()</code> and call <code>SpoLogin</code> from the method.

### Example
1. ExampleWebTest - shows how to use SpoLogin in coded web test.

### Context Parameters
1. Tenant - SharePoint online tenant name for example "libinuko" for libinuko.sharepoint.com 
2. UserName - Full email address of the username. For example joe@hotmail.com 
3. UserPassword - You know this.

### Issues
1. Log the issue or contact me directly.
