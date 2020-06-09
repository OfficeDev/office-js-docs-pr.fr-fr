---
title: Utilisation des API REST Outlook d’un complément Outlook
description: Découvrez comment utiliser des API REST Outlook à partir d’un complément Outlook pour obtenir un jeton d’accès.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 7cd26c26e277d7d5fe93664494eb84b4e94bcc47
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611616"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a><span data-ttu-id="bae8e-103">Utilisation des API REST Outlook d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="bae8e-103">Use the Outlook REST APIs from an Outlook add-in</span></span>

<span data-ttu-id="bae8e-p101">L’espace de noms [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) permet d’accéder à de nombreux champs communs pour les messages et les rendez-vous. Toutefois, dans certains scénarios, un complément peut avoir besoin d’accéder aux données qui ne sont pas exposées par l’espace de noms. Par exemple, le complément peut dépendre de propriétés personnalisées définies par une application extérieure ou avoir besoin rechercher dans la boîte aux lettres de l’utilisateur des messages provenant du même expéditeur. Dans ces scénarios, l’[API REST Outlook](/outlook/rest/index) est la méthode recommandée pour récupérer les informations.</span><span class="sxs-lookup"><span data-stu-id="bae8e-p101">The [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest/index) is the recommended method to retrieve the information.</span></span>

## <a name="get-an-access-token"></a><span data-ttu-id="bae8e-108">Obtenir un jeton d’accès</span><span class="sxs-lookup"><span data-stu-id="bae8e-108">Get an access token</span></span>

<span data-ttu-id="bae8e-p102">Les API REST Outlook nécessitent un jeton du porteur dans l’en-tête `Authorization`. En règle générale, les applications utilisent les flux OAuth2 pour extraire un jeton. Toutefois, les compléments peuvent récupérer un jeton sans mettre en œuvre OAuth2 à l’aide de la nouvelle méthode [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) introduite dans la version 1.5 de l’ensemble de conditions de boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="bae8e-p102">The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method introduced in the Mailbox requirement set 1.5.</span></span>

<span data-ttu-id="bae8e-112">En définissant l’option `isRest` sur `true`, vous pouvez demander un jeton compatible avec les API REST.</span><span class="sxs-lookup"><span data-stu-id="bae8e-112">By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.</span></span>

### <a name="add-in-permissions-and-token-scope"></a><span data-ttu-id="bae8e-113">Autorisations des compléments et étendue du jeton</span><span class="sxs-lookup"><span data-stu-id="bae8e-113">Add-in permissions and token scope</span></span>

<span data-ttu-id="bae8e-p103">Il est important de savoir de quel niveau d’accès votre complément aura besoin avec les API REST. Dans la plupart des cas, le jeton renvoyé par `getCallbackTokenAsync` fournit un accès en lecture seule à l’élément actif uniquement. Cela est vrai même si votre complément spécifie le niveau d’autorisation `ReadWriteItem` dans son manifeste.</span><span class="sxs-lookup"><span data-stu-id="bae8e-p103">It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.</span></span>

<span data-ttu-id="bae8e-p104">Si votre complément nécessitera un accès en écriture à l’élément actif ou à d’autres éléments de la boîte aux lettres de l’utilisateur, votre complément doit spécifier le niveau d’autorisation `ReadWriteMailbox` dans son manifeste. Dans ce cas, le jeton renvoyé contiendra l’accès en lecture/écriture aux messages, événements et contacts de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bae8e-p104">If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.</span></span>

### <a name="example"></a><span data-ttu-id="bae8e-119">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae8e-119">Example</span></span>

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a><span data-ttu-id="bae8e-120">Obtenir l’ID de l’élément</span><span class="sxs-lookup"><span data-stu-id="bae8e-120">Get the item ID</span></span>

<span data-ttu-id="bae8e-121">Pour extraire l’élément en cours via REST, votre complément aura besoin de l’ID de l’élément, correctement mis en forme pour REST.</span><span class="sxs-lookup"><span data-stu-id="bae8e-121">To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST.</span></span> <span data-ttu-id="bae8e-122">Cet ID peut être extrait de la propriété [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), mais certaines vérifications doivent être apportées pour vous assurer qu’il s’agit d’un ID au format REST.</span><span class="sxs-lookup"><span data-stu-id="bae8e-122">This is obtained from the [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.</span></span>

- <span data-ttu-id="bae8e-123">Dans Outlook Mobile, la valeur renvoyée par `Office.context.mailbox.item.itemId` est un ID au format REST et peut être utilisé comme tel.</span><span class="sxs-lookup"><span data-stu-id="bae8e-123">In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.</span></span>
- <span data-ttu-id="bae8e-124">Dans d’autres clients Outlook, la valeur renvoyée par `Office.context.mailbox.item.itemId` est un ID au format EWS et doit être convertie à l’aide de la méthode [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="bae8e-124">In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>
- <span data-ttu-id="bae8e-125">Vous devez également convertir l’ID de pièce jointe en ID au format REST afin de l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="bae8e-125">Note you must also convert Attachment ID to a REST-formatted ID in order to use it.</span></span> <span data-ttu-id="bae8e-126">La raison pour laquelle les ID doivent être convertis est que les ID EWS peuvent contenir des valeurs approuvées autres que des URL, ce qui entraîne des problèmes pour REST.</span><span class="sxs-lookup"><span data-stu-id="bae8e-126">The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.</span></span>

<span data-ttu-id="bae8e-127">Votre complément peut déterminer dans quel client Outlook il est chargé en consultant la propriété [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname).</span><span class="sxs-lookup"><span data-stu-id="bae8e-127">Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) property.</span></span>

### <a name="example"></a><span data-ttu-id="bae8e-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae8e-128">Example</span></span>

```js
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## <a name="get-the-rest-api-url"></a><span data-ttu-id="bae8e-129">Obtenir l’URL de l’API REST</span><span class="sxs-lookup"><span data-stu-id="bae8e-129">Get the REST API URL</span></span>

<span data-ttu-id="bae8e-p107">La dernière information dont votre complément a besoin pour appeler l’API REST est le nom d’hôte qu'il doit utiliser pour envoyer des demandes d’API. Cette information figure dans la propriété [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties).</span><span class="sxs-lookup"><span data-stu-id="bae8e-p107">The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property.</span></span>

### <a name="example"></a><span data-ttu-id="bae8e-132">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae8e-132">Example</span></span>

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a><span data-ttu-id="bae8e-133">Appel de l’API</span><span class="sxs-lookup"><span data-stu-id="bae8e-133">Call the API</span></span>

<span data-ttu-id="bae8e-134">Une fois que votre complément a le jeton d’accès, l’ID de l’élément et l’URL de l’API REST, il peut transmettre ces informations à un service principal qui appelle l’API REST, ou l’appeler directement à l’aide d’AJAX.</span><span class="sxs-lookup"><span data-stu-id="bae8e-134">After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX.</span></span> <span data-ttu-id="bae8e-135">L’exemple suivant appelle l’API REST de courrier Outlook pour obtenir le message actuel.</span><span class="sxs-lookup"><span data-stu-id="bae8e-135">The following example calls the Outlook Mail REST API to get the current message.</span></span>

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    var subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a><span data-ttu-id="bae8e-136">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bae8e-136">See also</span></span>

- <span data-ttu-id="bae8e-137">Pour obtenir un exemple qui appelle les API REST à partir d’un complément Outlook, reportez-vous à [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) sur GitHub.</span><span class="sxs-lookup"><span data-stu-id="bae8e-137">For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
- <span data-ttu-id="bae8e-138">Les API REST Outlook sont également disponibles via le point de terminaison Microsoft Graph, mais il existe quelques différences clés, notamment sur la façon dont votre complément obtient un jeton d’accès.</span><span class="sxs-lookup"><span data-stu-id="bae8e-138">Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token.</span></span> <span data-ttu-id="bae8e-139">Pour plus d’informations, reportez-vous à [API REST Outlook via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span><span class="sxs-lookup"><span data-stu-id="bae8e-139">For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span></span>