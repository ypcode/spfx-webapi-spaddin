# Business Documents Web API


## Enable CORS

In Package Manager console, type the following command
(Make sure the spaddin-webapiWeb project is targeted)
> Install-Package Microsoft.AspNet.WebApi.Cors

in App_Start/WebApiConfig.cs, add the following line
```
  config.EnableCors();
```

in BusinessDocumentsController.cs, decorate the class with the EnableCors attribute
Set the origins argument for your SharePoint site domain address
(e.g. https://mytenant.sharepoint.com or https://sharepoint.mycompany.local)

It is also possible to enable it globally or per controller action

Check this out for the complete documentation:
https://docs.microsoft.com/en-us/aspnet/web-api/overview/security/enabling-cross-origin-requests-in-web-api


## The Web API must be registered
The Web API registration ensures the user is properly authenticated through SharePoint.
The Web Application will then keep the user context in cache, and it will be used when the Web API is called.

It means the user has to have reached once the endpoint that will register the Web API.
In the case of a full external SPA, this endpoint will typically be the page itself.

In the case of a SPFx solution, we need to reach it through another mean that will be transparent to the user.
We will then use a hidden Iframe pointing to this URL.
When the Iframe is loaded, a flag 'authenticated' is set.

The API calls can be done only after this flag is set, the technique consists in implementing an asynchronous loop
to poll this flag value before actually performing the API call
