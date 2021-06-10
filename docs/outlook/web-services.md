---
title: Utiliser les services Web Exchange (EWS) à partir d’un complément Outlook.
description: Fournit un exemple qui illustre comment un complément Outlook peut demander des informations à partir des Services Web Exchange.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 16d20ca30f2860b8103257860a8619c1d51d8523
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/09/2021
ms.locfileid: "52853961"
---
# <a name="call-web-services-from-an-outlook-add-in"></a><span data-ttu-id="b390f-103">Appeler des services Web à partir d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="b390f-103">Call web services from an Outlook add-in</span></span>

<span data-ttu-id="b390f-p101">Votre complément peut utiliser les services web Exchange (EWS) d’un ordinateur exécutant Exchange Server 2013, un service web disponible sur le serveur qui fournit l’emplacement source de l’interface utilisateur du complément ou un service web disponible sur Internet. Cette rubrique fournit des exemples expliquant comment un complément Outlook peut demander des informations à partir d’EWS.</span><span class="sxs-lookup"><span data-stu-id="b390f-p101">Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.</span></span>

<span data-ttu-id="b390f-p102">La méthode d’appel d’un service Web dépend de l’emplacement de ce dernier. Le tableau 1 répertorie les méthodes d’appel d’un service Web en fonction de l’emplacement.</span><span class="sxs-lookup"><span data-stu-id="b390f-p102">The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.</span></span>


<span data-ttu-id="b390f-108">**Tableau 1. Méthodes d’appel de services web à partir d’un complément Outlook**</span><span class="sxs-lookup"><span data-stu-id="b390f-108">**Table 1. Ways to call web services from an Outlook add-in**</span></span>

<br/>

|<span data-ttu-id="b390f-109">**Emplacement des services web**</span><span class="sxs-lookup"><span data-stu-id="b390f-109">**Web service location**</span></span>|<span data-ttu-id="b390f-110">**Méthode d’appel du service web**</span><span class="sxs-lookup"><span data-stu-id="b390f-110">**Way to call the web service**</span></span>|
|:-----|:-----|
|<span data-ttu-id="b390f-111">Serveur Exchange qui héberge la boîte aux lettres cliente</span><span class="sxs-lookup"><span data-stu-id="b390f-111">The Exchange server that hosts the client mailbox</span></span>|<span data-ttu-id="b390f-p103">Utilisez la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour appeler les opérations EWS qui permettent d'ajouter des compléments de prise en charge. Le serveur Exchange qui héberge la boîte aux lettres expose également EWS.</span><span class="sxs-lookup"><span data-stu-id="b390f-p103">Use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.</span></span>|
|<span data-ttu-id="b390f-114">Serveur web qui fournit l’emplacement source de l’interface utilisateur du complément</span><span class="sxs-lookup"><span data-stu-id="b390f-114">The web server that provides the source location for the add-in UI</span></span>|<span data-ttu-id="b390f-p104">Appelez le service web au moyen des techniques JavaScript standard. Le code JavaScript présent dans le cadre de l’interface utilisateur s’exécute dans le contexte du serveur web qui fournit l’interface utilisateur. Il est donc capable d’appeler les services web sur ce serveur sans provoquer d’erreur de script intersite.</span><span class="sxs-lookup"><span data-stu-id="b390f-p104">Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.</span></span>|
|<span data-ttu-id="b390f-118">Tous les autres emplacements</span><span class="sxs-lookup"><span data-stu-id="b390f-118">All other locations</span></span>|<span data-ttu-id="b390f-p105">Créez un proxy pour le service web sur le serveur web qui fournit l’emplacement source de l’interface utilisateur. Si vous n’indiquez pas de proxy, votre complément ne s’exécutera pas en raison d’erreurs de script intersites. L’un des moyens de fournir un proxy consiste à utiliser JSON/P. Pour plus d’informations, voir [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md).</span><span class="sxs-lookup"><span data-stu-id="b390f-p105">Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).</span></span>|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a><span data-ttu-id="b390f-123">Utilisation de la méthode makeEwsRequestAsync pour accéder aux opérations EWS</span><span class="sxs-lookup"><span data-stu-id="b390f-123">Using the makeEwsRequestAsync method to access EWS operations</span></span>

<span data-ttu-id="b390f-124">Vous pouvez utiliser la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour effectuer une demande EWS auprès du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b390f-124">You can use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to make an EWS request to the Exchange server that hosts the user's mailbox.</span></span>

<span data-ttu-id="b390f-p106">EWS prend en charge en charge différentes opérations sur un serveur Exchange, par exemple, les opérations au niveau de l’élément pour copier, rechercher, mettre à jour ou envoyer un élément, et les opérations au niveau du dossier pour créer, obtenir ou mettre à jour un dossier. Pour exécuter une opération EWS, créez une demande SOAP XML pour cette opération. Une fois l’opération terminée, vous obtenez une réponse SOAP XML qui contient les données correspondant à l’opération. Les demandes et les réponses SOAP EWS suivent le schéma défini dans le fichier Messages.xsd. Comme d’autres fichiers de schéma EWS, le fichier Message.xsd se trouve dans le répertoire virtuel IIS qui héberge EWS.</span><span class="sxs-lookup"><span data-stu-id="b390f-p106">EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.</span></span>

<span data-ttu-id="b390f-130">Pour utiliser la `makeEwsRequestAsync` méthode pour lancer une opération EWS, fournissez les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b390f-130">To use the `makeEwsRequestAsync` method to initiate an EWS operation, provide the following:</span></span>

- <span data-ttu-id="b390f-131">Code XML pour la demande SOAP pour cette opération EWS, en tant qu’argument du paramètre  _data_</span><span class="sxs-lookup"><span data-stu-id="b390f-131">The XML for the SOAP request for that EWS operation, as an argument to the  _data_ parameter</span></span>

- <span data-ttu-id="b390f-132">Méthode de rappel (en tant qu’argument  _callback_)</span><span class="sxs-lookup"><span data-stu-id="b390f-132">A callback method (as the  _callback_ argument)</span></span>

- <span data-ttu-id="b390f-133">Données d’entrée facultatives pour cette méthode de rappel (en tant qu’argument  _userContext_)</span><span class="sxs-lookup"><span data-stu-id="b390f-133">Any optional input data for that callback method (as the  _userContext_ argument)</span></span>

<span data-ttu-id="b390f-134">Une fois la demande SOAP EWS terminée, Outlook appelle la méthode de rappel avec un argument, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b390f-134">When the EWS SOAP request is complete, Outlook calls the callback method with one argument, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="b390f-135">La méthode de rappel peut accéder à deux propriétés de l’objet : la propriété, qui contient la réponse SOAP XML de l’opération `AsyncResult` EWS, et éventuellement la propriété, qui contient toutes les données transmises en tant que `value` `asyncContext` `userContext` paramètre.</span><span class="sxs-lookup"><span data-stu-id="b390f-135">The callback method can access two properties of the `AsyncResult` object: the `value` property, which contains the XML SOAP response of the EWS operation, and optionally, the `asyncContext` property, which contains any data passed as the `userContext` parameter.</span></span> <span data-ttu-id="b390f-136">En règle générale, la méthode de rappel analyse ensuite le code XML dans la réponse SOAP pour obtenir les informations pertinentes et traite ces informations comme il se doit.</span><span class="sxs-lookup"><span data-stu-id="b390f-136">Typically, the callback method then parses the XML in the SOAP response to get any relevant information, and processes that information accordingly.</span></span>


## <a name="tips-for-parsing-ews-responses"></a><span data-ttu-id="b390f-137">Conseils pour l’analyse des réponses EWS</span><span class="sxs-lookup"><span data-stu-id="b390f-137">Tips for parsing EWS responses</span></span>

<span data-ttu-id="b390f-138">Lors de l’analyse d’une réponse SOAP à partir d’une opération EWS, notez les problèmes dépendant du navigateur suivants :</span><span class="sxs-lookup"><span data-stu-id="b390f-138">When parsing a SOAP response from an EWS operation, note the following browser-dependent issues:</span></span>


- <span data-ttu-id="b390f-139">Spécifiez le préfixe d’un nom de balise lors de l’utilisation de la méthode DOM, pour inclure la prise en `getElementsByTagName` charge d’Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="b390f-139">Specify the prefix for a tag name when using the DOM method `getElementsByTagName`, to include support for Internet Explorer.</span></span>

  <span data-ttu-id="b390f-140">`getElementsByTagName` se comporte différemment selon le type de navigateur.</span><span class="sxs-lookup"><span data-stu-id="b390f-140">`getElementsByTagName` behaves differently depending on browser type.</span></span> <span data-ttu-id="b390f-141">Par exemple, une réponse EWS peut contenir le XML suivant (formaté et abrégé à des fins d’affichage) :</span><span class="sxs-lookup"><span data-stu-id="b390f-141">For example, an EWS response can contain the following XML (formatted and abbreviated for display purposes):</span></span>

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   <span data-ttu-id="b390f-142">Le code, comme dans les exemples suivants, fonctionne sur un navigateur tel que Chrome pour obtenir le code XML entouré par les `ExtendedProperty` balises :</span><span class="sxs-lookup"><span data-stu-id="b390f-142">Code, as in the following, would work on a browser like Chrome to get the XML enclosed by the `ExtendedProperty` tags:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   <span data-ttu-id="b390f-143">Sur Internet Explorer, vous devez inclure le préfixe `t:` du nom de balise, comme indiqué ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="b390f-143">On Internet Explorer, you must include the `t:` prefix of the tag name, as shown below:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- <span data-ttu-id="b390f-144">Utilisez la propriété DOM pour obtenir le contenu d’une balise dans une réponse `textContent` EWS, comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="b390f-144">Use the DOM property `textContent` to get the contents of a tag in an EWS response, as shown below:</span></span>

   ```js
      content = $.parseJSON(value.textContent);
   ```

   <span data-ttu-id="b390f-145">D’autres propriétés telles `innerHTML` que peut ne pas fonctionner sur Internet Explorer pour certaines balises dans une réponse EWS.</span><span class="sxs-lookup"><span data-stu-id="b390f-145">Other properties such as `innerHTML` may not work on Internet Explorer for some tags in an EWS response.</span></span>


## <a name="example"></a><span data-ttu-id="b390f-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="b390f-146">Example</span></span>

<span data-ttu-id="b390f-147">L’exemple suivant appelle `makeEwsRequestAsync` l’utilisation de [l’opération GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b390f-147">The following example calls `makeEwsRequestAsync` to use the [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to get the subject of an item.</span></span> <span data-ttu-id="b390f-148">Cet exemple comprend les trois fonctions suivantes :</span><span class="sxs-lookup"><span data-stu-id="b390f-148">This example includes the following three functions:</span></span>

-  <span data-ttu-id="b390f-149">`getSubjectRequest`Prend un ID d’élément comme entrée et renvoie le XML de la demande SOAP à appeler &ndash; `GetItem` pour l’élément spécifié.</span><span class="sxs-lookup"><span data-stu-id="b390f-149">`getSubjectRequest` &ndash; Takes an item ID as input, and returns the XML for the SOAP request to call `GetItem` for the specified item.</span></span>

-  <span data-ttu-id="b390f-150">`sendRequest`Appels pour obtenir la demande SOAP pour l’élément sélectionné, puis passe la demande SOAP et la méthode de rappel, pour obtenir l’objet de &ndash;  `getSubjectRequest` `callback` `makeEwsRequestAsync` l’élément spécifié.</span><span class="sxs-lookup"><span data-stu-id="b390f-150">`sendRequest` &ndash; Calls  `getSubjectRequest` to get the SOAP request for the selected item, then passes the SOAP request and the callback method, `callback`, to `makeEwsRequestAsync` to get the subject of the specified item.</span></span>

-  <span data-ttu-id="b390f-151">`callback` &ndash; Traite la réponse SOAP qui comprend l’objet et d’autres informations sur l’élément spécifié.</span><span class="sxs-lookup"><span data-stu-id="b390f-151">`callback` &ndash; Processes the SOAP response which includes any subject and other information about the specified item.</span></span>


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


## <a name="ews-operations-that-add-ins-support"></a><span data-ttu-id="b390f-152">Opérations EWS prises en charge par les compléments</span><span class="sxs-lookup"><span data-stu-id="b390f-152">EWS operations that add-ins support</span></span>

<span data-ttu-id="b390f-153">Outlook peuvent accéder à un sous-ensemble d’opérations disponibles dans EWS via la `makeEwsRequestAsync` méthode.</span><span class="sxs-lookup"><span data-stu-id="b390f-153">Outlook add-ins can access a subset of operations that are available in EWS via the `makeEwsRequestAsync` method.</span></span> <span data-ttu-id="b390f-154">Si vous ne connaissez pas les opérations EWS et que vous ne savez pas comment utiliser la méthode pour accéder à une opération, commencez par un exemple de requête SOAP pour personnaliser votre `makeEwsRequestAsync` argument _de_ données.</span><span class="sxs-lookup"><span data-stu-id="b390f-154">If you are unfamiliar with EWS operations and how to use the `makeEwsRequestAsync` method to access an operation, start with a SOAP request example to customize your _data_ argument.</span></span>

<span data-ttu-id="b390f-155">La description suivante décrit comment utiliser la `makeEwsRequestAsync` méthode :</span><span class="sxs-lookup"><span data-stu-id="b390f-155">The following describes how you can use the `makeEwsRequestAsync` method:</span></span>

1. <span data-ttu-id="b390f-156">Dans le XML, remplacez les ID d’éléments et les attributs d’opération EWS par les valeurs appropriées.</span><span class="sxs-lookup"><span data-stu-id="b390f-156">In the XML, substitute any item IDs and relevant EWS operation attributes with appropriate values.</span></span>

2. <span data-ttu-id="b390f-157">Incluez la requête SOAP en tant qu’argument pour le  _paramètre de_ données de `makeEwsRequestAsync` .</span><span class="sxs-lookup"><span data-stu-id="b390f-157">Include the SOAP request as an argument for the  _data_ parameter of `makeEwsRequestAsync`.</span></span>

3. <span data-ttu-id="b390f-158">Spécifiez une méthode de rappel et un `makeEwsRequestAsync` appel.</span><span class="sxs-lookup"><span data-stu-id="b390f-158">Specify a callback method and call `makeEwsRequestAsync`.</span></span>

4. <span data-ttu-id="b390f-159">Dans la méthode de rappel, vérifiez les résultats de l’opération dans la réponse SOAP.</span><span class="sxs-lookup"><span data-stu-id="b390f-159">In the callback method, verify the results of the operation in the SOAP response.</span></span>

5. <span data-ttu-id="b390f-160">Utilisez les résultats de l’opération EWS en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="b390f-160">Use the results of the EWS operation according to your needs.</span></span>

<span data-ttu-id="b390f-p111">Le tableau suivant répertorie les opérations EWS prises en charge par les compléments. Pour afficher des exemples de demandes et réponses SOAP, choisissez le lien correspondant à chaque opération. Pour plus d’informations sur les opérations EWS, voir [Opérations EWS dans Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="b390f-p111">The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span></span>

<span data-ttu-id="b390f-164">**Tableau 2. Opérations EWS prises en charge**</span><span class="sxs-lookup"><span data-stu-id="b390f-164">**Table 2. Supported EWS operations**</span></span>

<br/>

|<span data-ttu-id="b390f-165">**Opération EWS**</span><span class="sxs-lookup"><span data-stu-id="b390f-165">**EWS operation**</span></span>|<span data-ttu-id="b390f-166">**Description**</span><span class="sxs-lookup"><span data-stu-id="b390f-166">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="b390f-167">Opération CopyItem</span><span class="sxs-lookup"><span data-stu-id="b390f-167">CopyItem operation</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)|<span data-ttu-id="b390f-168">Copie les éléments spécifiés et place les nouveaux éléments dans un dossier spécifique dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-168">Copies the specified items and puts the new items in a designated folder in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-169">Opération CreateFolder</span><span class="sxs-lookup"><span data-stu-id="b390f-169">CreateFolder operation</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)|<span data-ttu-id="b390f-170">Crée les dossiers dans l’emplacement spécifié dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-170">Creates folders in the specified location in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-171">Opération CreateItem</span><span class="sxs-lookup"><span data-stu-id="b390f-171">CreateItem operation</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)|<span data-ttu-id="b390f-172">Crée les éléments spécifiés dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-172">Creates the specified items in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-173">Opération ExpandDL</span><span class="sxs-lookup"><span data-stu-id="b390f-173">ExpandDL operation</span></span>](/exchange/client-developer/web-service-reference/expanddl-operation)|<span data-ttu-id="b390f-174">Affiche l’appartenance complète des listes de distribution.</span><span class="sxs-lookup"><span data-stu-id="b390f-174">Displays the full membership of distribution lists.</span></span>|
|[<span data-ttu-id="b390f-175">Opération FindConversation</span><span class="sxs-lookup"><span data-stu-id="b390f-175">FindConversation operation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)|<span data-ttu-id="b390f-176">Énumère une liste des conversations dans le dossier spécifié dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-176">Enumerates a list of conversations in the specified folder in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-177">Opération FindFolder</span><span class="sxs-lookup"><span data-stu-id="b390f-177">FindFolder operation</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)|<span data-ttu-id="b390f-178">Cherche les sous-dossiers d’un dossier donné et retourne un ensemble de propriétés qui décrit l’ensemble de sous-dossiers.</span><span class="sxs-lookup"><span data-stu-id="b390f-178">Finds subfolders of an identified folder and returns a set of properties that describe the set of subfolders.</span></span>|
|[<span data-ttu-id="b390f-179">Opération FindItem</span><span class="sxs-lookup"><span data-stu-id="b390f-179">FindItem operation</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)|<span data-ttu-id="b390f-180">Identifie les éléments situés dans un dossier donné dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-180">Identifies items that are located in a specified folder in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-181">Opération GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="b390f-181">GetConversationItems operation</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)|<span data-ttu-id="b390f-182">Obtient un ou plusieurs ensembles d’éléments organisés en nœuds dans une conversation.</span><span class="sxs-lookup"><span data-stu-id="b390f-182">Gets one or more sets of items that are organized in nodes in a conversation.</span></span>|
|[<span data-ttu-id="b390f-183">Opération GetFolder</span><span class="sxs-lookup"><span data-stu-id="b390f-183">GetFolder operation</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)|<span data-ttu-id="b390f-184">Obtient les propriétés spécifiées et le contenu des dossiers de la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-184">Gets the specified properties and contents of folders from the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-185">Opération GetItem</span><span class="sxs-lookup"><span data-stu-id="b390f-185">GetItem operation</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)|<span data-ttu-id="b390f-186">Obtient les propriétés spécifiées et le contenu des éléments de la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-186">Gets the specified properties and contents of items from the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-187">Opération GetUserAvailability</span><span class="sxs-lookup"><span data-stu-id="b390f-187">GetUserAvailability operation</span></span>](/exchange/client-developer/web-service-reference/getuseravailability-operation)|<span data-ttu-id="b390f-188">Fournit des informations détaillées sur la disponibilité d’un ensemble d’utilisateurs, salles et ressources sur une période spécifiée.</span><span class="sxs-lookup"><span data-stu-id="b390f-188">Provides detailed information about the availability of a set of users, rooms, and resources within a specified time period.</span></span>|
|[<span data-ttu-id="b390f-189">Opération MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="b390f-189">MarkAsJunk operation</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)|<span data-ttu-id="b390f-190">Déplace les messages électroniques vers le dossier Courrier indésirable, et ajoute ou supprime les expéditeurs des messages de la liste des expéditeurs bloqués.</span><span class="sxs-lookup"><span data-stu-id="b390f-190">Moves email messages to the Junk Email folder, and adds or removes senders of the messages from the blocked senders list accordingly.</span></span>|
|[<span data-ttu-id="b390f-191">Opération MoveItem</span><span class="sxs-lookup"><span data-stu-id="b390f-191">MoveItem operation</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)|<span data-ttu-id="b390f-192">Déplace les éléments dans un dossier de destination unique dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-192">Moves items to a single destination folder in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-193">Opération ResolveNames</span><span class="sxs-lookup"><span data-stu-id="b390f-193">ResolveNames operation</span></span>](/exchange/client-developer/web-service-reference/resolvenames-operation)|<span data-ttu-id="b390f-194">Résout les adresses de messagerie et les noms d’affichage ambigus.</span><span class="sxs-lookup"><span data-stu-id="b390f-194">Resolves ambiguous email addresses and display names.</span></span>|
|[<span data-ttu-id="b390f-195">Opération SendItem</span><span class="sxs-lookup"><span data-stu-id="b390f-195">SendItem operation</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)|<span data-ttu-id="b390f-196">Envoie les messages électroniques situés dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-196">Sends email messages that are located in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-197">Opération UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="b390f-197">UpdateFolder operation</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)|<span data-ttu-id="b390f-198">Modifie les propriétés des dossiers existants dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-198">Modifies the properties of existing folders in the Exchange store.</span></span>|
|[<span data-ttu-id="b390f-199">Opération UpdateItem</span><span class="sxs-lookup"><span data-stu-id="b390f-199">UpdateItem operation</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)|<span data-ttu-id="b390f-200">Modifie les propriétés des éléments existants dans la banque d’informations Exchange.</span><span class="sxs-lookup"><span data-stu-id="b390f-200">Modifies the properties of existing items in the Exchange store.</span></span>|

 > [!NOTE]
 > <span data-ttu-id="b390f-201">Les éléments FAI (Informations relatives au dossier) ne peuvent pas être mis à jour (ni créés) depuis un complément.</span><span class="sxs-lookup"><span data-stu-id="b390f-201">FAI (Folder Associated Information) items cannot be updated (or created) from an add-in.</span></span> <span data-ttu-id="b390f-202">Ces messages masqués sont stockés dans un dossier et utilisés pour stocker divers paramètres et données auxiliaires.</span><span class="sxs-lookup"><span data-stu-id="b390f-202">These hidden messages are stored in a folder and are used to store a variety of settings and auxiliary data.</span></span>  <span data-ttu-id="b390f-203">Si vous tentez d’utiliser l’opération UpdateItem, une erreur ErrorAccessDenied est générée : « L'annuaire d'entreprise n'est pas autorisé à mettre à jour ce type d'élément ».</span><span class="sxs-lookup"><span data-stu-id="b390f-203">Attempting to use the UpdateItem operation will throw an ErrorAccessDenied error: "Office extension is not allowed to update this type of item".</span></span> <span data-ttu-id="b390f-204">En guise d’alternative, vous pouvez utiliser l’[API managée EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) pour mettre à jour ces éléments à partir d’un client Windows ou d’une application serveur.</span><span class="sxs-lookup"><span data-stu-id="b390f-204">As an alternative, you may use the [EWS Managed API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) to update these items from a Windows client or a server application.</span></span> <span data-ttu-id="b390f-205">Soyez vigilant car les structures de données internes de type service peuvent être modifiées et sont susceptibles d’endommager votre solution.</span><span class="sxs-lookup"><span data-stu-id="b390f-205">Caution is recommended as internal, service-type data structures are subject to change and could break your solution.</span></span>


## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a><span data-ttu-id="b390f-206">Authentification et autorisation pour la méthode makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b390f-206">Authentication and permission considerations for makeEwsRequestAsync</span></span>

<span data-ttu-id="b390f-207">Lorsque vous utilisez la méthode, la demande est authentifiée à l’aide des informations d’identification du compte de messagerie `makeEwsRequestAsync` de l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="b390f-207">When you use the `makeEwsRequestAsync` method, the request is authenticated by using the email account credentials of the current user.</span></span> <span data-ttu-id="b390f-208">La méthode gère les informations d’identification pour vous afin de ne pas avoir à fournir d’informations d’identification `makeEwsRequestAsync` d’authentification avec votre demande.</span><span class="sxs-lookup"><span data-stu-id="b390f-208">The `makeEwsRequestAsync` method manages the credentials for you so that you do not have to provide authentication credentials with your request.</span></span>

> [!NOTE]
> <span data-ttu-id="b390f-209">L’administrateur de serveur doit utiliser la cmdlet [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) pour définir le paramètre _OAuthAuthentication_ sur **true** dans le répertoire EWS du serveur d’accès au client afin de permettre à la méthode d’effectuer des demandes `makeEwsRequestAsync` EWS.</span><span class="sxs-lookup"><span data-stu-id="b390f-209">The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) cmdlet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

<span data-ttu-id="b390f-210">Votre add-in doit spécifier l’autorisation dans son manifeste de `ReadWriteMailbox` add-in pour utiliser la `makeEwsRequestAsync` méthode.</span><span class="sxs-lookup"><span data-stu-id="b390f-210">Your add-in must specify the `ReadWriteMailbox` permission in its add-in manifest to use the `makeEwsRequestAsync` method.</span></span> <span data-ttu-id="b390f-211">Pour plus d’informations sur l’utilisation de l’autorisation, voir la `ReadWriteMailbox` section [Autorisation ReadWriteMailbox](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) dans [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="b390f-211">For information about using the `ReadWriteMailbox` permission, see the section [ReadWriteMailbox permission](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) in [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b390f-212">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b390f-212">See also</span></span>

- [<span data-ttu-id="b390f-213">Confidentialité et sécurité pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="b390f-213">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
- [<span data-ttu-id="b390f-214">Résolutions des limites de stratégie d’origine identique dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="b390f-214">Addressing same-origin policy limitations in Office Add-ins</span></span>](../develop/addressing-same-origin-policy-limitations.md)
- [<span data-ttu-id="b390f-215">Référence EWS pour Exchange</span><span class="sxs-lookup"><span data-stu-id="b390f-215">EWS reference for Exchange</span></span>](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [<span data-ttu-id="b390f-216">Applications de messagerie pour Outlook et EWS dans Exchange</span><span class="sxs-lookup"><span data-stu-id="b390f-216">Mail apps for Outlook and EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

<span data-ttu-id="b390f-217">Consultez la rubrique suivante pour créer des services principaux pour les compléments à l’aide de l’API Web ASP.NET :</span><span class="sxs-lookup"><span data-stu-id="b390f-217">See the following for creating backend services for add-ins using ASP.NET Web API:</span></span>

- [<span data-ttu-id="b390f-218">Créer un service web pour un complément Office à l’aide de l’API Web ASP.NET</span><span class="sxs-lookup"><span data-stu-id="b390f-218">Create a web service for an Office Add-in using the ASP.NET Web API</span></span>](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [<span data-ttu-id="b390f-219">Principes fondamentaux de la création d’un service HTTP à l’aide de l’API Web ASP.NET</span><span class="sxs-lookup"><span data-stu-id="b390f-219">The basics of building an HTTP service using ASP.NET Web API</span></span>](https://dotnet.microsoft.com/apps/aspnet/apis)