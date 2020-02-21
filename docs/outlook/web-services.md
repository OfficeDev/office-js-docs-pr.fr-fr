---
title: Utiliser les services Web Exchange (EWS) à partir d’un complément Outlook.
description: Fournit un exemple qui illustre comment un complément Outlook peut demander des informations à partir des Services Web Exchange.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4c0c97a9a796dc1f257b1a0b0ec880b3ca3d8e74
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166064"
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

Pour utiliser la méthode **makeEwsRequestAsync** pour initier une opération EWS, indiquez les éléments suivants :

- Code XML pour la demande SOAP pour cette opération EWS, en tant qu’argument du paramètre  _data_

- Méthode de rappel (en tant qu’argument  _callback_)

- Données d’entrée facultatives pour cette méthode de rappel (en tant qu’argument  _userContext_)

Une fois la demande SOAP EWS terminée, Outlook appelle la méthode de rappel avec un argument, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult). La méthode de rappel peut accéder à deux propriétés de l’objet  **AsyncResult** : la propriété  **value**, qui contient la réponse SOAP XML de l’opération EWS, et éventuellement la propriété  **asyncContext**, qui contient les données transmises en tant que paramètre  **userContext**. En règle générale, la méthode de rappel analyse ensuite le code XML dans la réponse SOAP pour obtenir les informations pertinentes et traite ces informations comme il se doit.


## <a name="tips-for-parsing-ews-responses"></a>Conseils pour l’analyse des réponses EWS

Lors de l’analyse d’une réponse SOAP à partir d’une opération EWS, notez les problèmes dépendant du navigateur suivants :


- Spécifiez le préfixe de nom de balise lorsque vous utilisez la méthode DOM **getElementsByTagName**, pour inclure la prise en charge d’Internet Explorer.

  **getElementsByTagName** se comporte différemment selon le type de navigateur. Par exemple, une réponse EWS peut contenir le code XML suivant (mis en forme et abrégé à des fins d’affichage) :

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   Un code tel que le suivant fonctionnera dans un navigateur tel que Chrome pour obtenir le code XML entouré par les balises **ExtendedProperty** :

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

- Utilisez la propriété DOM **textContent** pour obtenir le contenu d’une balise dans une réponse EWS, comme indiqué ci-dessous :
    
   ```js
      content = $.parseJSON(value.textContent);
   ```

   D’autres propriétés telles que **innerHTML** peuvent ne pas fonctionner sur Internet Explorer pour certaines balises dans une réponse EWS.
    

## <a name="example"></a>Exemple

L’exemple suivant appelle **makeEwsRequestAsync** pour utiliser l’opération [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) afin d’obtenir l’objet d’un élément. Cet exemple comprend les trois fonctions suivantes :

-  `getSubjectRequest` &ndash; Prend un ID d’élément comme entrée et retourne le XML pour la demande SOAP qui appelle **GetItem** pour l’élément spécifié.
    
-  `sendRequest` &ndash; Appellez `getSubjectRequest` pour obtenir la demande SOAP pour l’élément sélectionné, puis passez la demande SOAP et la méthode de rappel `callback` à **makeEwsRequestAsync** pour obtenir l’objet de l’élément spécifié.
    
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

Les compléments Outlook peuvent accéder à un sous-ensemble d’opérations disponibles dans EWS par le biais de la méthode **makeEwsRequestAsync**. Si vous ne connaissez pas les opérations EWS et ne savez pas comment utiliser la méthode **makeEwsRequestAsync** pour accéder à une opération, commencez avec un exemple de demande SOAP pour personnaliser votre argument _data_. 

Voici des explications sur la manière d’utiliser la méthode **makeEwsRequestAsync** :

1. Dans le XML, remplacez les ID d’éléments et les attributs d’opération EWS par les valeurs appropriées.
    
2. Intégrez la demande SOAP en tant qu’argument pour le paramètre  _data_ de **makeEwsRequestAsync**.
    
3. Spécifiez une méthode de rappel et appelez **makeEwsRequestAsync**.
    
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

Lorsque vous utilisez la méthode **makeEwsRequestAsync**, la demande est authentifiée à l’aide des informations d’identification du compte de messagerie de l’utilisateur actuel. La méthode **makeEwsRequestAsync** gère les informations d’identification pour vous de sorte que vous n’ayez pas à fournir d’informations d’identification d’authentification avec votre demande.

> [!NOTE]
> L’administrateur du serveur doit utiliser la cmdlet [New-WebServicesVirtualDirctory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) ou [Set-WebServicesVirtualDirecory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) pour définir le paramètre _OAuthAuthentication_ sur **true** dans le répertoire EWS du serveur Client Access afin d’activer la méthode **makeEwsRequestAsync** pour effectuer des demandes EWS.

Votre complément doit spécifier l’autorisation **ReadWriteMailbox** dans son manifeste de complément pour utiliser la méthode **makeEwsRequestAsync**. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox**, reportez-vous à la section [ReadWriteMailbox](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) dans [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).

> [!NOTE]
> L’administrateur du serveur doit utiliser la cmdlet [New-WebServicesVirtualDirctory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) ou [Set-WebServicesVirtualDirecory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) pour définir le paramètre _OAuthAuthentication_ sur **true** dans le répertoire EWS du serveur Client Access afin d’activer la méthode **makeEwsRequestAsync** pour effectuer des demandes EWS.



## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../develop/privacy-and-security.md)   
- [Résolutions des limites de stratégie d’origine identique dans les compléments Office](../develop/addressing-same-origin-policy-limitations.md)
- [Référence EWS pour Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)   
- [Applications de messagerie pour Outlook et EWS dans Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)
   
Consultez la rubrique suivante pour créer des services principaux pour les compléments à l’aide de l’API Web ASP.NET :

- [Créer un service web pour un complément Office à l’aide de l’API Web ASP.NET](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)    
- [Principes fondamentaux de la création d’un service HTTP à l’aide de l’API Web ASP.NET](https://www.asp.net/web-api)
    
