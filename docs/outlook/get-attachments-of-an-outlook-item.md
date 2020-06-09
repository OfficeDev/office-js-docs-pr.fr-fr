---
title: Obtenir des pièces jointes dans un complément Outlook
description: Votre complément peut utiliser les API de pièces jointes pour envoyer des informations sur les pièces jointes à un service distant.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: bcb8226ab0755351b9e3a365e40623d258887d3f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612079"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a><span data-ttu-id="cb3ab-103">Obtenir des pièces jointes d’un élément Outlook à partir du serveur</span><span class="sxs-lookup"><span data-stu-id="cb3ab-103">Get attachments of an Outlook item from the server</span></span>

<span data-ttu-id="cb3ab-p101">Un complément Outlook ne peut pas transmettre les pièces jointes d’un élément sélectionné directement au service distant exécuté sur votre serveur. À la place, le complément peut utiliser l’API de pièces jointes pour envoyer des informations relatives aux pièces jointes au service distant. Le service peut alors contacter directement le serveur Exchange pour récupérer les pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-p101">An Outlook add-in cannot pass the attachments of a selected item directly to the remote service that runs on your server. Instead, the add-in can use the attachments API to send information about the attachments to the remote service. The service can then contact the Exchange server directly to retrieve the attachments.</span></span>

<span data-ttu-id="cb3ab-107">Pour envoyer des informations sur les pièces jointes au service distant, vous utilisez les propriétés et la fonction suivantes :</span><span class="sxs-lookup"><span data-stu-id="cb3ab-107">To send attachment information to the remote service, you use the following properties and function:</span></span>

- <span data-ttu-id="cb3ab-p102">Propriété [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) &ndash; fournit l’URL des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres. Votre service utilise cette URL pour appeler la méthode [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) ou l’opération EWS [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation).</span><span class="sxs-lookup"><span data-stu-id="cb3ab-p102">[Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) property &ndash; Provides the URL of Exchange Web Services (EWS) on the Exchange server that hosts the mailbox. Your service uses this URL to call the [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation.</span></span>

- <span data-ttu-id="cb3ab-110">Propriété [Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) &ndash; obtient un tableau d’objets [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails), un pour chaque pièce jointe de l’élément.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-110">[Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property &ndash; Gets an array of [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) objects, one for each attachment to the item.</span></span>

- <span data-ttu-id="cb3ab-111">Fonction [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) &ndash; réalise un appel asynchrone vers le serveur Exchange hébergeant la boîte aux lettres pour obtenir un jeton de rappel que le serveur renvoie au serveur Exchange afin d’authentifier une demande de pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-111">[Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) function &ndash; Makes an asynchronous call to the Exchange server that hosts the mailbox to get a callback token that the server sends back to the Exchange server to authenticate a request for an attachment.</span></span>

## <a name="using-the-attachments-api"></a><span data-ttu-id="cb3ab-112">Utilisation de l’API de pièces jointes</span><span class="sxs-lookup"><span data-stu-id="cb3ab-112">Using the attachments API</span></span>

<span data-ttu-id="cb3ab-113">Pour utiliser l’API de pièces jointes afin d'obtenir des pièces jointes à partir d’une boîte aux lettres Exchange, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="cb3ab-113">To use the attachments API to get attachments from an Exchange mailbox, perform the following steps:</span></span>

1. <span data-ttu-id="cb3ab-114">Affichez le complément lorsque l’utilisateur visualise un message ou un rendez-vous qui contient une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-114">Show the add-in when the user is viewing a message or appointment that contains an attachment.</span></span>

1. <span data-ttu-id="cb3ab-115">Obtenez le jeton de rappel à partir du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-115">Get the callback token from the Exchange server.</span></span>

1. <span data-ttu-id="cb3ab-116">Envoyez le jeton de rappel et les informations de pièce jointe au service distant.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-116">Send the callback token and attachment information to the remote service.</span></span>

1. <span data-ttu-id="cb3ab-117">Obtenez des pièces jointes à partir du serveur Exchange à l’aide de la méthode  `ExchangeService.GetAttachments` ou de l’opération `GetAttachment`.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-117">Get the attachments from the Exchange server by using the `ExchangeService.GetAttachments` method or the `GetAttachment` operation.</span></span>

<span data-ttu-id="cb3ab-118">Chacune de ces étapes est décrite en détail dans les sections suivantes à l’aide du code de l’exemple [Compléments de messagerie pour Office : obtenir des pièces jointes d’un serveur Exchange](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).</span><span class="sxs-lookup"><span data-stu-id="cb3ab-118">Each of these steps is covered in detail in the following sections using code from the [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) sample.</span></span>

> [!NOTE]
> <span data-ttu-id="cb3ab-p103">§LTA Le code de ces exemples a été raccourci pour se concentrer sur les informations liées aux pièces jointes. L’exemple contient du code supplémentaire pour l’authentification du complément auprès du serveur distant et la gestion de l’état de la demande.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-p103">The code in these examples has been shortened to emphasize the attachment information. The sample contains additional code for authenticating the add-in with the remote server and managing the state of the request.</span></span>

## <a name="get-a-callback-token"></a><span data-ttu-id="cb3ab-121">Obtenir un jeton de rappel</span><span class="sxs-lookup"><span data-stu-id="cb3ab-121">Get a callback token</span></span>

<span data-ttu-id="cb3ab-122">L’objet [Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) fournit la fonction `getCallbackTokenAsync` afin d’obtenir un jeton dont peut se servir le serveur distant pour vous authentifier avec le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-122">The [Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) object provides the `getCallbackTokenAsync` function to get a token that the remote server can use to authenticate with the Exchange server.</span></span> <span data-ttu-id="cb3ab-123">Le code suivant montre une fonction dans un complément qui démarre la demande asynchrone pour obtenir le jeton de rappel, et la fonction de rappel qui récupère la réponse.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-123">The following code shows a function in an add-in that starts the asynchronous request to get the callback token, and the callback function that gets the response.</span></span> <span data-ttu-id="cb3ab-124">Le jeton de rappel est stocké dans l’objet de demande de service défini dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-124">The callback token is stored in the service request object that is defined in the next section.</span></span>

```js
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}
```

## <a name="send-attachment-information-to-the-remote-service"></a><span data-ttu-id="cb3ab-125">Envoyer des informations de pièce jointe au service distant</span><span class="sxs-lookup"><span data-stu-id="cb3ab-125">Send attachment information to the remote service</span></span>

<span data-ttu-id="cb3ab-p105">Le service distant appelé par votre complément définit les informations spécifiques relatives à l’envoi des informations de la pièce jointe au service. Dans cet exemple, le service distant est une API d’application web créée avec Visual Studio 2013. Le service distant attend les informations de la pièce jointe dans un objet JSON. Le code suivant initialise un objet qui contient les informations de la pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-p105">The remote service that your add-in calls defines the specifics of how you should send the attachment information to the service. In this example, the remote service is a Web API application created by using Visual Studio 2013. The remote service expects the attachment information in a JSON object. The following code initializes an object that contains the attachment information.</span></span>

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 var serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

<br/>

<span data-ttu-id="cb3ab-130">La propriété `Office.context.mailbox.item.attachments` inclut une collection d’objets `AttachmentDetails`, un par pièce jointe pour l’élément.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-130">The `Office.context.mailbox.item.attachments` property contains a collection of `AttachmentDetails` objects, one for each attachment to the item.</span></span> <span data-ttu-id="cb3ab-131">Dans la plupart des cas, le complément peut passer uniquement l’ID de propriété de pièce jointe d’un objet `AttachmentDetails` au service distant. </span><span class="sxs-lookup"><span data-stu-id="cb3ab-131">In most cases, the add-in can pass just the attachment ID property of an `AttachmentDetails` object to the remote service.</span></span> <span data-ttu-id="cb3ab-132">Si le service distant a besoin de détails supplémentaires sur la pièce jointe, vous pouvez passer l’intégralité ou des parties de l’objet `AttachmentDetails`.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-132">If the remote service needs more details about the attachment, you can pass all or part of the `AttachmentDetails` object.</span></span> <span data-ttu-id="cb3ab-133">Le code suivant définit une méthode qui place l’intégralité du tableau `AttachmentDetails` dans l’objet `serviceRequest` et envoie une demande au service distant.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-133">The following code defines a method that puts the entire `AttachmentDetails` array in the `serviceRequest` object and sends a request to the remote service.</span></span>

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (var i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      var names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (i = 0; i < response.attachmentNames.length; i++) {
        names += response.attachmentNames[i] + "<br />";
      }
      document.getElementById("names").innerHTML = names;
    } else {
      app.showNotification("Runtime error", response.message);
    }
  }).fail(function (status) {

  }).always(function () {
    $('.disable-while-sending').prop('disabled', false);
  })
}
```

## <a name="get-the-attachments-from-the-exchange-server"></a><span data-ttu-id="cb3ab-134">Obtenir des pièces jointes à partir du serveur Exchange</span><span class="sxs-lookup"><span data-stu-id="cb3ab-134">Get the attachments from the Exchange server</span></span>

<span data-ttu-id="cb3ab-p107">Votre service distant peut utiliser la méthode [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) de l’API managée EWS ou l’opération EWS [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) pour récupérer des pièces jointes à partir du serveur. L’application de service a besoin de deux objets pour désérialiser la chaîne JSON en objets .NET Framework pouvant être utilisés sur le serveur. Le code suivant indique les définitions des objets de désérialisation.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-p107">Your remote service can use either the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS Managed API method or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation to retrieve attachments from the server. The service application needs two objects to deserialize the JSON string into .NET Framework objects that can be used on the server. The following code shows the definitions of the deserialization objects.</span></span>

```cs
namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a><span data-ttu-id="cb3ab-138">Utiliser l’API managée EWS pour obtenir des pièces jointes</span><span class="sxs-lookup"><span data-stu-id="cb3ab-138">Use the EWS Managed API to get the attachments</span></span>

<span data-ttu-id="cb3ab-p108">Si vous utilisez l’[API managée EWS](https://go.microsoft.com/fwlink/?LinkID=255472) dans votre service distant, vous pouvez utiliser la méthode [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) qui va construire, recevoir et envoyer une demande SOAP EWS pour obtenir les pièces jointes. Nous vous recommandons d’utiliser l’API managée EWS car elle requiert moins de lignes de code et fournit une interface plus intuitive pour les appels vers EWS. Le code suivant effectue une demande pour récupérer toutes les pièces jointes et renvoie le nombre, ainsi que les noms des pièces jointes traitées.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-p108">If you use the [EWS Managed API](https://go.microsoft.com/fwlink/?LinkID=255472) in your remote service, you can use the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, which will construct, send, and receive an EWS SOAP request to get the attachments. We recommend that you use the EWS Managed API because it requires fewer lines of code and provides a more intuitive interface for making calls to EWS. The following code makes one request to retrieve all the attachments, and returns the count and names of the attachments processed.</span></span>

```cs
private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  // Create an ExchangeService object, set the credentials and the EWS URL.
  ExchangeService service = new ExchangeService();
  service.Credentials = new OAuthCredentials(request.attachmentToken);
  service.Url = new Uri(request.ewsUrl);

  var attachmentIds = new List<string>();

  foreach (AttachmentDetails attachment in request.attachments)
  {
    attachmentIds.Add(attachment.id);
  }

  // Call the GetAttachments method to retrieve the attachments on the message.
  // This method results in a GetAttachments EWS SOAP request and response
  // from the Exchange server.
  var getAttachmentsResponse =
    service.GetAttachments(attachmentIds.ToArray(),
                            null,
                            new PropertySet(BasePropertySet.FirstClassProperties,
                                            ItemSchema.MimeContent));

  if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
  {
    foreach (var attachmentResponse in getAttachmentsResponse)
    {
      attachmentNames.Add(attachmentResponse.Attachment.Name);

      // Write the content of each attachment to a stream.
      if (attachmentResponse.Attachment is FileAttachment)
      {
        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
        Stream s = new MemoryStream(fileAttachment.Content);
        // Process the contents of the attachment here.
      }

      if (attachmentResponse.Attachment is ItemAttachment)
      {
        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
        // Process the contents of the attachment here.
      }

      attachmentsProcessedCount++;
    }
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

### <a name="use-ews-to-get-the-attachments"></a><span data-ttu-id="cb3ab-142">Utiliser EWS pour obtenir les pièces jointes</span><span class="sxs-lookup"><span data-stu-id="cb3ab-142">Use EWS to get the attachments</span></span>

<span data-ttu-id="cb3ab-143">Si vous utilisez EWS dans votre service distant, vous devez construire une demande SOAP [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) pour obtenir les pièces jointes à partir du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-143">If you use EWS in your remote service, you need to construct a [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP request to get the attachments from the Exchange server.</span></span> <span data-ttu-id="cb3ab-144">Le code suivant renvoie une chaîne qui fournit la demande SOAP.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-144">The following code returns a string that provides the SOAP request.</span></span> <span data-ttu-id="cb3ab-145">Le service distant utilise la méthode `String.Format` pour insérer l’ID d’une pièce jointe dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-145">The remote service uses the `String.Format` method to insert the attachment ID for an attachment into the string.</span></span>


```cs
private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""https://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
```

<br/>

<span data-ttu-id="cb3ab-146">Enfin, la méthode suivante utilise une demande `GetAttachment` EWS pour obtenir les pièces jointes à partir du serveur Exchange. </span><span class="sxs-lookup"><span data-stu-id="cb3ab-146">Finally, the following method does the work of using an EWS `GetAttachment` request to get the attachments from the Exchange server.</span></span> <span data-ttu-id="cb3ab-147">Cette implémentation effectue une demande individuelle pour chaque pièce jointe et renvoie le nombre de pièces jointes traitées. </span><span class="sxs-lookup"><span data-stu-id="cb3ab-147">This implementation makes an individual request for each attachment, and returns the count of attachments processed.</span></span> <span data-ttu-id="cb3ab-148">Chaque réponse est traitée dans une méthode `ProcessXmlResponse` distincte, définie ci-après.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-148">Each response is processed in a separate `ProcessXmlResponse` method, defined next.</span></span>

```cs
private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  foreach (var attachment in request.attachments)
  {
    // Prepare a web request object.
    HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
    webRequest.Headers.Add("Authorization",
      string.Format("Bearer {0}", request.attachmentToken));
    webRequest.PreAuthenticate = true;
    webRequest.AllowAutoRedirect = false;
    webRequest.Method = "POST";
    webRequest.ContentType = "text/xml; charset=utf-8";

    // Construct the SOAP message for the GetAttachment operation.
    byte[] bodyBytes = Encoding.UTF8.GetBytes(
      string.Format(GetAttachmentSoapRequest, attachment.id));
    webRequest.ContentLength = bodyBytes.Length;

    Stream requestStream = webRequest.GetRequestStream();
    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
    requestStream.Close();

    // Make the request to the Exchange server and get the response.
    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

    // If the response is okay, create an XML document from the response
    // and process the request.
    if (webResponse.StatusCode == HttpStatusCode.OK)
    {
      var responseStream = webResponse.GetResponseStream();

      var responseEnvelope = XElement.Load(responseStream);

      // After creating a memory stream containing the contents of the
      // attachment, this method writes the XML document to the trace output.
      // Your service would perform it's processing here.
      if (responseEnvelope != null)
      {
        var processResult = ProcessXmlResponse(responseEnvelope);
        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

      }

      // Close the response stream.
      responseStream.Close();
      webResponse.Close();

    }
    // If the response is not OK, return an error message for the
    // attachment.
    else
    {
      var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
        "Error message: {1}.", attachment.name, webResponse.StatusDescription);
      attachmentNames.Add(errorString);
    }
    attachmentsProcessedCount++;
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

<br/>

<span data-ttu-id="cb3ab-149">Chaque réponse de l’opération `GetAttachment` est envoyée à la méthode `ProcessXmlResponse`.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-149">Each response from the `GetAttachment` operation is sent to the `ProcessXmlResponse` method.</span></span> <span data-ttu-id="cb3ab-150">Cette méthode consulte la réponse à la recherche d’erreurs.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-150">This method checks the response for errors.</span></span> <span data-ttu-id="cb3ab-151">Si elle ne trouve pas d’erreur, elle traite les fichiers joints et les éléments joints.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-151">If it doesn't find any errors, it processes file attachments and item attachments.</span></span> <span data-ttu-id="cb3ab-152">La méthode `ProcessXmlResponse` effectue l’essentiel du traitement de la pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="cb3ab-152">The `ProcessXmlResponse` method performs the bulk of the work to process the attachment.</span></span>

```cs
// This method processes the response from the Exchange server.
// In your application the bulk of the processing occurs here.
private string ProcessXmlResponse(XElement responseEnvelope)
{
  // First, check the response for web service errors.
  var errorCodes = from errorCode in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                    select errorCode;
  // Return the first error code found.
  foreach (var errorCode in errorCodes)
  {
    if (errorCode.Value != "NoError")
    {
      return string.Format("Could not process result. Error: {0}", errorCode.Value);
    }
  }

  // No errors found, proceed with processing the content.
  // First, get and process file attachments.
  var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                        select fileAttachment;
  foreach(var fileAttachment in fileAttachments)
  {
    var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
    var fileData = System.Convert.FromBase64String(fileContent.Value);
    var s = new MemoryStream(fileData);
    // Process the file attachment here.
  }

  // Second, get and process item attachments.
  var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                        select itemAttachment;
  foreach(var itemAttachment in itemAttachments)
  {
    var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
    if (message != null)
    {
      // Process a message here.
      break;
    }
    var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
    if (calendarItem != null)
    {
      // Process calendar item here.
      break;
    }
    var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
    if (contact != null)
    {
      // Process contact here.
      break;
    }
    var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
    if (task != null)
    {
      // Process task here.
      break;
    }
    var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
    if (meetingMessage != null)
    {
      // Process meeting message here.
      break;
    }
    var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
    if (meetingRequest != null)
    {
      // Process meeting request here.
      break;
    }
    var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
    if (meetingResponse != null)
    {
      // Process meeting response here.
      break;
    }
    var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
    if (meetingCancellation != null)
    {
      // Process meeting cancellation here.
      break;
    }
  }

  return string.Empty;
}
```

## <a name="see-also"></a><span data-ttu-id="cb3ab-153">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cb3ab-153">See also</span></span>

- [<span data-ttu-id="cb3ab-154">Créer des compléments Outlook pour des formulaires de lecture</span><span class="sxs-lookup"><span data-stu-id="cb3ab-154">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="cb3ab-155">Explorer l’API managée EWS, EWS et les services web dans Exchange</span><span class="sxs-lookup"><span data-stu-id="cb3ab-155">Explore the EWS Managed API, EWS, and web services in Exchange</span></span>](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [<span data-ttu-id="cb3ab-156">Prise en main des applications clientes d'API managée EWS</span><span class="sxs-lookup"><span data-stu-id="cb3ab-156">Get started with EWS Managed API client applications</span></span>](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [<span data-ttu-id="cb3ab-157">Exemple de complément Outlook AttachmentsDemo</span><span class="sxs-lookup"><span data-stu-id="cb3ab-157">AttachmentsDemo Sample Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-attachments-demo)
