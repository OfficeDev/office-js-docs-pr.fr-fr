---
title: Utiliser les services Web Exchange (EWS) à partir d’un complément Outlook.
description: Fournit un exemple qui illustre comment un complément Outlook peut demander des informations à partir des Services Web Exchange.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: d3f400b7f6c5cb65233dfaf66cd4c592bf975f65
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930308"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Appeler des services Web à partir d’un complément Outlook

Votre complément peut utiliser les services web Exchange (EWS) d’un ordinateur exécutant Exchange Server 2013, un service web disponible sur le serveur qui fournit l’emplacement source de l’interface utilisateur du complément ou un service web disponible sur Internet. Cette rubrique fournit des exemples expliquant comment un complément Outlook peut demander des informations à partir d’EWS.

La méthode d’appel d’un service Web dépend de l’emplacement de ce dernier. Le tableau 1 répertorie les méthodes d’appel d’un service Web en fonction de l’emplacement.


**Tableau 1. Méthodes d’appel de services web à partir d’un complément Outlook**

<br/>

|**Emplacement des services web**|**Méthode d’appel du service web**|
|:-----|:-----|
|Serveur Exchange qui héberge la boîte aux lettres cliente|Utilisez la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour appeler les opérations EWS qui permettent d'ajouter des compléments de prise en charge. Le serveur Exchange qui héberge la boîte aux lettres expose également EWS.|
|Serveur web qui fournit l’emplacement source de l’interface utilisateur du complément|Appelez le service web au moyen des techniques JavaScript standard. Le code JavaScript présent dans le cadre de l’interface utilisateur s’exécute dans le contexte du serveur web qui fournit l’interface utilisateur. Il est donc capable d’appeler les services web sur ce serveur sans provoquer d’erreur de script intersite.|
|Tous les autres emplacements|Créez un proxy pour le service web sur le serveur web qui fournit l’emplacement source de l’interface utilisateur. Si vous n’indiquez pas de proxy, votre complément ne s’exécutera pas en raison d’erreurs de script intersites. L’un des moyens de fournir un proxy consiste à utiliser JSON/P. Pour plus d’informations, voir [Confidentialité et sécurité pour les compléments Office](../develop/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Utilisation de la méthode makeEwsRequestAsync pour accéder aux opérations EWS

Vous pouvez utiliser la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour effectuer une demande EWS auprès du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.

EWS prend en charge en charge différentes opérations sur un serveur Exchange, par exemple, les opérations au niveau de l’élément pour copier, rechercher, mettre à jour ou envoyer un élément, et les opérations au niveau du dossier pour créer, obtenir ou mettre à jour un dossier. Pour exécuter une opération EWS, créez une demande SOAP XML pour cette opération. Une fois l’opération terminée, vous obtenez une réponse SOAP XML qui contient les données correspondant à l’opération. Les demandes et les réponses SOAP EWS suivent le schéma défini dans le fichier Messages.xsd. Comme d’autres fichiers de schéma EWS, le fichier Message.xsd se trouve dans le répertoire virtuel IIS qui héberge EWS.

Pour utiliser la `makeEwsRequestAsync` méthode pour lancer une opération EWS, fournissez les éléments suivants :

- Code XML pour la demande SOAP pour cette opération EWS, en tant qu’argument du paramètre  _data_

- Méthode de rappel (en tant qu’argument  _callback_)

- Données d’entrée facultatives pour cette méthode de rappel (en tant qu’argument  _userContext_)

Lorsque la demande de SOAP EWS est terminée, Outlook appelle la méthode de rappel avec un argument, qui est un objet [asyncResult](/javascript/api/office/office.asyncresult) . La méthode de rappel peut accéder à deux propriétés `AsyncResult` de l’objet `value` : la propriété, qui contient la réponse SOAP XML de l’opération EWS et, éventuellement, `asyncContext` la propriété, qui contient toutes les données transmises `userContext` en tant que paramètre. En règle générale, la méthode de rappel analyse ensuite le code XML dans la réponse SOAP pour obtenir des informations pertinentes et traite ces informations en conséquence.


## <a name="tips-for-parsing-ews-responses"></a>Conseils pour l’analyse des réponses EWS

Lors de l’analyse d’une réponse SOAP à partir d’une opération EWS, notez les problèmes dépendant du navigateur suivants :


- Spécifier le préfixe d’un nom de balise lorsque vous `getElementsByTagName`utilisez la méthode DOM, pour inclure la prise en charge d’Internet Explorer.

  `getElementsByTagName`se comporte différemment en fonction du type de navigateur. Par exemple, une réponse EWS peut contenir le code XML suivant (formaté et abrégé à des fins d’affichage) :

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   Le code, comme dans l’exemple suivant, fonctionnerait sur un navigateur comme chrome pour obtenir le code XML `ExtendedProperty` entouré par les balises :

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   Sur Internet Explorer, vous devez inclure le préfixe `t:` du nom de balise, comme indiqué ci-dessous :

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- Utilisez la propriété `textContent` DOM pour obtenir le contenu d’une balise dans une réponse EWS, comme indiqué ci-dessous :

   ```js
      content = $.parseJSON(value.textContent);
   ```

   D’autres propriétés telles `innerHTML` que peuvent ne pas fonctionner sur Internet Explorer pour certaines balises dans une réponse EWS.


## <a name="example"></a>Exemple

L’exemple suivant appelle `makeEwsRequestAsync` l’utilisation de l’opération [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour obtenir l’objet d’un élément. Cet exemple inclut les trois fonctions suivantes :

-  `getSubjectRequest`&ndash; Prend un ID d’élément comme entrée et retourne le XML pour la demande SOAP à appeler `GetItem` pour l’élément spécifié.

-  `sendRequest`&ndash; Appelle `getSubjectRequest` pour obtenir la demande SOAP pour l’élément sélectionné, puis transmet la demande SOAP et la méthode de rappel `callback`,, `makeEwsRequestAsync` à pour obtenir l’objet de l’élément spécifié.

-  `callback` &ndash; Traite la réponse SOAP qui comprend l’objet et d’autres informations sur l’élément spécifié.


```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
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
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}
```


## <a name="ews-operations-that-add-ins-support"></a>Opérations EWS prises en charge par les compléments

Les compléments Outlook peuvent accéder à un sous-ensemble d’opérations disponibles dans EWS via `makeEwsRequestAsync` la méthode. Si vous n’êtes pas familiarisé avec les opérations EWS et comment utiliser `makeEwsRequestAsync` la méthode pour accéder à une opération, commencez avec un exemple de demande SOAP pour personnaliser votre argument de _données_ .

Voici comment vous pouvez utiliser la `makeEwsRequestAsync` méthode :

1. Dans le XML, remplacez les ID d’éléments et les attributs d’opération EWS par les valeurs appropriées.

2. Incluez la demande SOAP en tant qu’argument _data_ pour le paramètre `makeEwsRequestAsync`Data de.

3. Spécifiez une méthode de rappel `makeEwsRequestAsync`et appelez.

4. Dans la méthode de rappel, vérifiez les résultats de l’opération dans la réponse SOAP.

5. Utilisez les résultats de l’opération EWS en fonction de vos besoins.

Le tableau suivant répertorie les opérations EWS prises en charge par les compléments. Pour afficher des exemples de demandes et réponses SOAP, choisissez le lien correspondant à chaque opération. Pour plus d’informations sur les opérations EWS, voir [Opérations EWS dans Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Tableau 2. Opérations EWS prises en charge**

<br/>

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

Lorsque vous utilisez la `makeEwsRequestAsync` méthode, la demande est authentifiée à l’aide des informations d’identification du compte de messagerie de l’utilisateur actuel. La `makeEwsRequestAsync` méthode gère les informations d’identification pour vous afin de ne pas avoir à fournir d’informations d’identification d’authentification avec votre demande.

> [!NOTE]
> L’administrateur de serveur doit utiliser la cmdlet [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) pour définir le paramètre _OAuthAuthentication_ sur **true** dans le répertoire EWS du serveur d’accès au client afin de `makeEwsRequestAsync` permettre à la méthode d’effectuer des demandes EWS.

Votre complément doit spécifier l' `ReadWriteMailbox` autorisation dans son manifeste de complément pour utiliser la `makeEwsRequestAsync` méthode. Pour plus d’informations sur `ReadWriteMailbox` l’utilisation de l’autorisation, voir la section [ReadWriteMailbox autorisation](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) dans [Understanding Outlook Add-in permissions](understanding-outlook-add-in-permissions.md).

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../develop/privacy-and-security.md)
- [Résolutions des limites de stratégie d’origine identique dans les compléments Office](../develop/addressing-same-origin-policy-limitations.md)
- [Référence EWS pour Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Applications de messagerie pour Outlook et EWS dans Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

Consultez la rubrique suivante pour créer des services principaux pour les compléments à l’aide de l’API Web ASP.NET :

- [Créer un service web pour un complément Office à l’aide de l’API Web ASP.NET](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)
- [Principes fondamentaux de la création d’un service HTTP à l’aide de l’API Web ASP.NET](https://www.asp.net/web-api)
