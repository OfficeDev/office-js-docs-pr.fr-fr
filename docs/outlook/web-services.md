---
title: Utiliser les services Web Exchange (EWS) à partir d’un complément Outlook.
description: Fournit un exemple qui illustre comment un complément Outlook peut demander des informations à partir des Services Web Exchange.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 94fff26fc7f9c16e2e385d6c44c128e4b03f968e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467012"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Appeler des services Web à partir d’un complément Outlook

Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.

The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.

**Tableau 1. Méthodes d’appel de services web à partir d’un complément Outlook**

|**Emplacement des services web**|**Méthode d’appel du service web**|
|:-----|:-----|
|Serveur Exchange qui héberge la boîte aux lettres cliente|Use the [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.|
|Serveur web qui fournit l’emplacement source de l’interface utilisateur du complément|Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.|
|Tous les autres emplacements|Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Utilisation de la méthode makeEwsRequestAsync pour accéder aux opérations EWS

Vous pouvez utiliser la méthode [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) pour effectuer une demande EWS auprès du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.

EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.

Pour utiliser la `makeEwsRequestAsync` méthode pour lancer une opération EWS, fournissez les éléments suivants :

- Code XML pour la demande SOAP pour cette opération EWS, en tant qu’argument du paramètre  _data_

- Fonction de rappel (en tant qu’argument  _de rappel_ )

- Toutes les données d’entrée facultatives pour cette fonction de rappel (en tant qu’argument  _userContext_ )

Une fois la requête SOAP EWS terminée, Outlook appelle la fonction de rappel avec un argument, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult) . La fonction de rappel peut accéder à deux propriétés de l’objet `AsyncResult` : la `value` propriété, qui contient la réponse SOAP XML de l’opération EWS, et éventuellement la `asyncContext` propriété, qui contient toutes les données passées en tant que `userContext` paramètre. En règle générale, la fonction de rappel analyse ensuite le code XML dans la réponse SOAP pour obtenir des informations pertinentes et traite ces informations en conséquence.

## <a name="tips-for-parsing-ews-responses"></a>Conseils pour l’analyse des réponses EWS

Lors de l’analyse d’une réponse SOAP à partir d’une opération EWS, notez les problèmes suivants dépendant du navigateur.

- Spécifiez le préfixe d’un nom de balise lors de l’utilisation de la méthode `getElementsByTagName`DOM pour inclure la prise en charge d’Internet Explorer.

  `getElementsByTagName` se comporte différemment en fonction du type de navigateur. Par exemple, une réponse EWS peut contenir le code XML suivant (mis en forme et abrégé à des fins d’affichage).

   ```XML
   <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
   PropertyName="MyProperty" 
   PropertyType="String"/>
   <t:Value>{
   ...
   }</t:Value></t:ExtendedProperty>
   ```

   Le code, comme dans ce qui suit, fonctionne sur un navigateur comme Chrome pour obtenir le code XML placé entre les `ExtendedProperty` balises.

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("ExtendedProperty")
   });
   ```

   Sur Internet Explorer, vous devez inclure le `t:` préfixe du nom de la balise, comme suit.

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("t:ExtendedProperty")
   });
   ```

- Utilisez la propriété `textContent` DOM pour obtenir le contenu d’une balise dans une réponse EWS, comme suit.

   ```js
   content = $.parseJSON(value.textContent);
   ```

   D’autres propriétés, par exemple `innerHTML` , peuvent ne pas fonctionner sur Internet Explorer pour certaines balises dans une réponse EWS.

## <a name="example"></a>Exemple

L’exemple suivant appelle `makeEwsRequestAsync` l’utilisation de l’opération [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour obtenir l’objet d’un élément. Cet exemple inclut les trois fonctions suivantes.

- `getSubjectRequest`&ndash; Prend un ID d’élément comme entrée et retourne le code XML de la demande SOAP à appeler `GetItem` pour l’élément spécifié.

- `sendRequest`&ndash; Appelle `getSubjectRequest` pour obtenir la demande SOAP pour l’élément sélectionné, puis transmet la requête SOAP et la fonction de rappel, `callback`pour `makeEwsRequestAsync` obtenir l’objet de l’élément spécifié.

- `callback` &ndash; Traite la réponse SOAP qui comprend l’objet et d’autres informations sur l’élément spécifié.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   const result = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return result;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   const mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   const result = asyncResult.value;
   const context = asyncResult.context;

   // Process the returned response here.
}
```

## <a name="ews-operations-that-add-ins-support"></a>Opérations EWS prises en charge par les compléments

Les compléments Outlook peuvent accéder à un sous-ensemble d’opérations disponibles dans EWS via la `makeEwsRequestAsync` méthode. Si vous n’êtes pas familiarisé avec les opérations EWS et la façon d’utiliser la `makeEwsRequestAsync` méthode pour accéder à une opération, commencez par un exemple de requête SOAP pour personnaliser votre argument _de données_ .

L’article suivant décrit comment utiliser la `makeEwsRequestAsync` méthode.

1. Dans le XML, remplacez les ID d’éléments et les attributs d’opération EWS par les valeurs appropriées.

1. Incluez la requête SOAP comme argument pour le paramètre de  _données_ de `makeEwsRequestAsync`.

1. Spécifiez une fonction de rappel et un appel `makeEwsRequestAsync`.

1. Dans la fonction de rappel, vérifiez les résultats de l’opération dans la réponse SOAP.

1. Utilisez les résultats de l’opération EWS en fonction de vos besoins.

The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Tableau 2. Opérations EWS prises en charge**

|**Opération EWS**|**Description**|
|:-----|:-----|
|[Opération CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)|Copie les éléments spécifiés et place les nouveaux éléments dans un dossier spécifique dans la banque d’informations Exchange.|
|[Opération CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)|Crée les dossiers dans l’emplacement spécifié dans la banque d’informations Exchange.|
|[Opération CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)|Crée les éléments spécifiés dans la banque d’informations Exchange.|
|[Opération ExpandDL](/exchange/client-developer/web-service-reference/expanddl-operation)|Affiche l’appartenance complète des listes de distribution.|
|[Opération FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)|Énumère une liste des conversations dans le dossier spécifié dans la banque d’informations Exchange.|
|[Opération FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)|Cherche les sous-dossiers d’un dossier donné et retourne un ensemble de propriétés qui décrit l’ensemble de sous-dossiers.|
|[Opération FindItem](/exchange/client-developer/web-service-reference/finditem-operation)|Identifie les éléments situés dans un dossier donné dans la banque d’informations Exchange.|
|[Opération GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)|Obtient un ou plusieurs ensembles d’éléments organisés en nœuds dans une conversation.|
|[Opération GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)|Obtient les propriétés spécifiées et le contenu des dossiers de la banque d’informations Exchange.|
|[Opération GetItem](/exchange/client-developer/web-service-reference/getitem-operation)|Obtient les propriétés spécifiées et le contenu des éléments de la banque d’informations Exchange.|
|[Opération GetUserAvailability](/exchange/client-developer/web-service-reference/getuseravailability-operation)|Fournit des informations détaillées sur la disponibilité d’un ensemble d’utilisateurs, salles et ressources sur une période spécifiée.|
|[Opération MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)|Déplace les messages électroniques vers le dossier Courrier indésirable, et ajoute ou supprime les expéditeurs des messages de la liste des expéditeurs bloqués.|
|[Opération MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)|Déplace les éléments dans un dossier de destination unique dans la banque d’informations Exchange.|
|[Opération ResolveNames](/exchange/client-developer/web-service-reference/resolvenames-operation)|Résout les adresses de messagerie et les noms d’affichage ambigus.|
|[Opération SendItem](/exchange/client-developer/web-service-reference/senditem-operation)|Envoie les messages électroniques situés dans la banque d’informations Exchange.|
|[Opération UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)|Modifie les propriétés des dossiers existants dans la banque d’informations Exchange.|
|[Opération UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)|Modifie les propriétés des éléments existants dans la banque d’informations Exchange.|

 > [!NOTE]
 > Les éléments FAI (Informations relatives au dossier) ne peuvent pas être mis à jour (ni créés) depuis un complément. Ces messages masqués sont stockés dans un dossier et utilisés pour stocker divers paramètres et données auxiliaires.  Si vous tentez d’utiliser l’opération UpdateItem, une erreur ErrorAccessDenied est générée : « L'annuaire d'entreprise n'est pas autorisé à mettre à jour ce type d'élément ». En guise d’alternative, vous pouvez utiliser l’[API managée EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) pour mettre à jour ces éléments à partir d’un client Windows ou d’une application serveur. Soyez vigilant car les structures de données internes de type service peuvent être modifiées et sont susceptibles d’endommager votre solution.

## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a>Authentification et autorisation pour la méthode makeEwsRequestAsync

Lorsque vous utilisez la `makeEwsRequestAsync` méthode, la demande est authentifiée à l’aide des informations d’identification du compte de messagerie de l’utilisateur actuel. La `makeEwsRequestAsync` méthode gère les informations d’identification pour vous afin que vous n’ayez pas à fournir d’informations d’identification d’authentification avec votre demande.

> [!NOTE]
> L’administrateur de serveur doit utiliser l’applet de commande [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) pour définir le paramètre `true` _OAuthAuthentication_ sur le répertoire EWS du serveur d’accès au client afin de permettre à la `makeEwsRequestAsync` méthode d’effectuer des requêtes EWS.

Pour utiliser la `makeEwsRequestAsync` méthode, votre complément doit demander l’autorisation de **boîte aux lettres en lecture/écriture** dans le manifeste. Le balisage varie en fonction du type de manifeste.

- **Manifeste XML** : définissez l’élément **\<Permissions\>** sur **ReadWriteMailbox**.
- **Manifeste Teams (préversion)** : définissez la propriété « name » d’un objet dans le tableau « authorization.permissions.resourceSpecific » sur « Mailbox.ReadWrite.User ».

Pour plus d’informations sur l’utilisation de l’autorisation de **boîte aux lettres en lecture/écriture** , consultez [l’autorisation de boîte aux lettres en lecture-écriture](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission).

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Résolutions des limites de stratégie d’origine identique dans les compléments Office](../develop/addressing-same-origin-policy-limitations.md)
- [Référence EWS pour Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Applications de messagerie pour Outlook et EWS dans Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

Consultez les informations suivantes pour créer des services principaux pour les compléments à l’aide de API Web ASP.NET.

- [Créer un service web pour un complément Office à l’aide de l’API Web ASP.NET](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [Principes fondamentaux de la création d’un service HTTP à l’aide de l’API Web ASP.NET](https://dotnet.microsoft.com/apps/aspnet/apis)
