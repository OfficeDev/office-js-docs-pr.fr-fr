---
title: Obtenir des pièces jointes dans un complément Outlook
description: Votre complément peut utiliser les API de pièces jointes pour envoyer des informations sur les pièces jointes à un service distant.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 637513a5ee94f4a3b9fa6b913f4c419dd5ec4d8e
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958824"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>Obtenir des pièces jointes d’un élément Outlook à partir du serveur

Vous pouvez obtenir les pièces jointes d’un élément Outlook de deux manières, mais l’option que vous utilisez dépend de votre scénario.

1. Envoyez les informations de pièce jointe à votre service distant.

    Votre complément peut utiliser l’API pièces jointes pour envoyer des informations sur les pièces jointes au service distant. Le service peut alors contacter directement le serveur Exchange pour récupérer les pièces jointes.

1. Utilisez l’API [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) , disponible à partir de l’ensemble de conditions requises 1.8. Formats pris en charge : [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat).

    Cette API peut être pratique si EWS/REST n’est pas disponible (par exemple, en raison de la configuration administrateur de votre serveur Exchange) ou si votre complément souhaite utiliser le contenu base64 directement en HTML ou JavaScript. En outre, l’API `getAttachmentContentAsync` est disponible dans les scénarios de composition où la pièce jointe n’a peut-être pas encore été synchronisée avec Exchange. Pour plus [d’informations, consultez Gérer les pièces jointes d’un élément dans un formulaire de composition dans Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md) .

Cet article explique en détail la première option. Pour envoyer des informations de pièce jointe au service distant, utilisez les propriétés et la méthode suivantes.

- Propriété [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) &ndash; fournit l’URL des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres. Votre service utilise cette URL pour appeler la méthode [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) ou l’opération EWS [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation).

- Propriété [Office.context.mailbox.item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) &ndash; obtient un tableau d’objets [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails), un pour chaque pièce jointe de l’élément.

- La méthode &ndash; [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) effectue un appel asynchrone au serveur Exchange qui héberge la boîte aux lettres pour obtenir un jeton de rappel que le serveur renvoie au serveur Exchange pour authentifier une demande de pièce jointe.

## <a name="using-the-attachments-api"></a>Utilisation de l’API de pièces jointes

Pour utiliser l’API pièces jointes pour obtenir des pièces jointes à partir d’une boîte aux lettres Exchange, effectuez les étapes suivantes.

1. Affichez le complément lorsque l’utilisateur visualise un message ou un rendez-vous qui contient une pièce jointe.

1. Obtenez le jeton de rappel à partir du serveur Exchange.

1. Envoyez le jeton de rappel et les informations de pièce jointe au service distant.

1. Obtenez des pièces jointes à partir du serveur Exchange à l’aide de la méthode  `ExchangeService.GetAttachments` ou de l’opération `GetAttachment`.

Chacune de ces étapes est décrite en détail dans les sections suivantes à l’aide du code de l’exemple [Compléments de messagerie pour Office : obtenir des pièces jointes d’un serveur Exchange](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).

> [!NOTE]
> §LTA Le code de ces exemples a été raccourci pour se concentrer sur les informations liées aux pièces jointes. L’exemple contient du code supplémentaire pour l’authentification du complément auprès du serveur distant et la gestion de l’état de la demande.

## <a name="get-a-callback-token"></a>Obtenir un jeton de rappel

L’objet [Office.context.mailbox](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox) fournit la `getCallbackTokenAsync` méthode permettant d’obtenir un jeton que le serveur distant peut utiliser pour s’authentifier auprès du serveur Exchange. Le code suivant montre une fonction dans un complément qui démarre la demande asynchrone pour obtenir le jeton de rappel, et la fonction de rappel qui récupère la réponse. Le jeton de rappel est stocké dans l’objet de demande de service défini dans la section suivante.

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

## <a name="send-attachment-information-to-the-remote-service"></a>Envoyer des informations de pièce jointe au service distant

Le service distant appelé par votre complément définit les informations spécifiques relatives à l’envoi des informations de la pièce jointe au service. Dans cet exemple, le service distant est une API d’application web créée avec Visual Studio 2013. Le service distant attend les informations de la pièce jointe dans un objet JSON. Le code suivant initialise un objet qui contient les informations de la pièce jointe.

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 const serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

La propriété `Office.context.mailbox.item.attachments` inclut une collection d’objets `AttachmentDetails`, un par pièce jointe pour l’élément. Dans la plupart des cas, le complément peut passer uniquement l’ID de propriété de pièce jointe d’un objet `AttachmentDetails` au service distant.  Si le service distant a besoin de détails supplémentaires sur la pièce jointe, vous pouvez passer l’intégralité ou des parties de l’objet `AttachmentDetails`. Le code suivant définit une méthode qui place l’intégralité du tableau `AttachmentDetails` dans l’objet `serviceRequest` et envoie une demande au service distant.

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (let i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      const names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (let i = 0; i < response.attachmentNames.length; i++) {
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

## <a name="get-the-attachments-from-the-exchange-server"></a>Obtenir des pièces jointes à partir du serveur Exchange

Votre service distant peut utiliser la méthode [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) de l’API managée EWS ou l’opération EWS [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) pour récupérer des pièces jointes à partir du serveur. L’application de service a besoin de deux objets pour désérialiser la chaîne JSON en objets .NET Framework pouvant être utilisés sur le serveur. Le code suivant indique les définitions des objets de désérialisation.

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

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>Utiliser l’API managée EWS pour obtenir des pièces jointes

Si vous utilisez l’[API managée EWS](/exchange/client-developer/web-service-reference/ews-managed-api-reference-for-exchange) dans votre service distant, vous pouvez utiliser la méthode [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) qui va construire, recevoir et envoyer une demande SOAP EWS pour obtenir les pièces jointes. Nous vous recommandons d’utiliser l’API managée EWS car elle requiert moins de lignes de code et fournit une interface plus intuitive pour les appels vers EWS. Le code suivant effectue une demande pour récupérer toutes les pièces jointes et renvoie le nombre, ainsi que les noms des pièces jointes traitées.

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

### <a name="use-ews-to-get-the-attachments"></a>Utiliser EWS pour obtenir les pièces jointes

Si vous utilisez EWS dans votre service distant, vous devez construire une demande SOAP [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) pour obtenir les pièces jointes à partir du serveur Exchange. Le code suivant renvoie une chaîne qui fournit la demande SOAP. Le service distant utilise la méthode `String.Format` pour insérer l’ID d’une pièce jointe dans la chaîne.

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

Enfin, la méthode suivante utilise une demande `GetAttachment` EWS pour obtenir les pièces jointes à partir du serveur Exchange.  Cette implémentation effectue une demande individuelle pour chaque pièce jointe et renvoie le nombre de pièces jointes traitées.  Chaque réponse est traitée dans une méthode `ProcessXmlResponse` distincte, définie ci-après.

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

Chaque réponse de l’opération `GetAttachment` est envoyée à la méthode `ProcessXmlResponse`. Cette méthode consulte la réponse à la recherche d’erreurs. Si elle ne trouve pas d’erreur, elle traite les fichiers joints et les éléments joints. La méthode `ProcessXmlResponse` effectue l’essentiel du traitement de la pièce jointe.

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

## <a name="see-also"></a>Voir aussi

- [Créer des compléments Outlook pour des formulaires de lecture](read-scenario.md)
- [Explorer l’API managée EWS, EWS et les services web dans Exchange](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [Prise en main des applications clientes d'API managée EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [Authentification unique du complément Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
