---
title: Utilisation des API REST Outlook d’un complément Outlook
description: Découvrez comment utiliser des API REST Outlook à partir d’un complément Outlook pour obtenir un jeton d’accès.
ms.date: 09/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8630cc6c075d80546e019ba41f57d46d97eb0f13
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616001"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Utilisation des API REST Outlook d’un complément Outlook

L’espace de noms [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) permet d’accéder à de nombreux champs communs pour les messages et les rendez-vous. Toutefois, dans certains scénarios, un complément peut avoir besoin d’accéder aux données qui ne sont pas exposées par l’espace de noms. Par exemple, le complément peut dépendre de propriétés personnalisées définies par une application extérieure ou avoir besoin rechercher dans la boîte aux lettres de l’utilisateur des messages provenant du même expéditeur. Dans ces scénarios, l’[API REST Outlook](/outlook/rest) est la méthode recommandée pour récupérer les informations.

> [!IMPORTANT]
> **Les API REST Outlook sont déconseillées**
>
> Les points de terminaison REST Outlook seront entièrement désactivés le 30 novembre 2022 (pour plus d’informations, voir [l’annonce de novembre 2020](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)). Vous devez migrer des compléments existants pour utiliser [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph). Pour obtenir des conseils, consultez Comparer les points de terminaison [de l’API REST Microsoft Graph et Outlook](/outlook/rest/compare-graph).
>
> Pour vous aider à effectuer la migration, les compléments actifs qui utilisent le service REST peuvent bénéficier d’une exemption pour continuer à utiliser le service jusqu’à la [fin du support étendu pour Outlook 2019 le 14 octobre 2025](/lifecycle/end-of-support/end-of-support-2025). Cela inclut les nouveaux compléments développés après le 30 novembre 2022. L’exemption est basée sur l’ID de manifeste du complément et s’applique aux compléments hébergés par AppSource et publiés en privé.
>
> L’identification automatique du trafic des compléments Outlook qui utilisent le service REST est actuellement testée pour la validation de l’exemption. Si vous souhaitez participer à cette phase de test, veuillez remplir le formulaire de [vérification du complément DE L’API REST](https://aka.ms/RESTCheck) avant novembre 2022. Pour plus d’informations, consultez le [billet de blog d’appel de la communauté des compléments Office d’août 2022](https://pnp.github.io/blog/office-add-ins-community-call/2022-08-10/).

## <a name="get-an-access-token"></a>Obtenir un jeton d’accès

Les API REST Outlook nécessitent un jeton du porteur dans l’en-tête `Authorization`. En règle générale, les applications utilisent les flux OAuth2 pour extraire un jeton. Toutefois, les compléments peuvent récupérer un jeton sans mettre en œuvre OAuth2 à l’aide de la nouvelle méthode [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) introduite dans la version 1.5 de l’ensemble de conditions de boîte aux lettres.

En définissant l’option `isRest` sur `true`, vous pouvez demander un jeton compatible avec les API REST.

### <a name="add-in-permissions-and-token-scope"></a>Autorisations des compléments et étendue du jeton

Il est important de savoir de quel niveau d’accès votre complément aura besoin avec les API REST. Dans la plupart des cas, le jeton renvoyé par `getCallbackTokenAsync` fournit un accès en lecture seule à l’élément actif uniquement. Cela est vrai même si votre complément spécifie le niveau d’autorisation `ReadWriteItem` dans son manifeste.

Si votre complément nécessitera un accès en écriture à l’élément actif ou à d’autres éléments de la boîte aux lettres de l’utilisateur, votre complément doit spécifier le niveau d’autorisation `ReadWriteMailbox` dans son manifeste. Dans ce cas, le jeton renvoyé contiendra l’accès en lecture/écriture aux messages, événements et contacts de l’utilisateur.

### <a name="example"></a>Exemple

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    const accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>Obtenir l’ID de l’élément

Pour extraire l’élément en cours via REST, votre complément aura besoin de l’ID de l’élément, correctement mis en forme pour REST. Cet ID peut être extrait de la propriété [Office.context.mailbox.item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), mais certaines vérifications doivent être apportées pour vous assurer qu’il s’agit d’un ID au format REST.

- Dans Outlook Mobile, la valeur renvoyée par `Office.context.mailbox.item.itemId` est un ID au format REST et peut être utilisé comme tel.
- Dans d’autres clients Outlook, la valeur renvoyée par `Office.context.mailbox.item.itemId` est un ID au format EWS et doit être convertie à l’aide de la méthode [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).
- Vous devez également convertir l’ID de pièce jointe en ID au format REST afin de l’utiliser. La raison pour laquelle les ID doivent être convertis est que les ID EWS peuvent contenir des valeurs approuvées autres que des URL, ce qui entraîne des problèmes pour REST.

Votre complément peut déterminer dans quel client Outlook il est chargé en consultant la propriété [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostname-member).

### <a name="example"></a>Exemple

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

## <a name="get-the-rest-api-url"></a>Obtenir l’URL de l’API REST

La dernière information dont votre complément a besoin pour appeler l’API REST est le nom d’hôte qu'il doit utiliser pour envoyer des demandes d’API. Cette information figure dans la propriété [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties).

### <a name="example"></a>Exemple

```js
// Example: https://outlook.office.com
const restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>Appel de l’API

Une fois que votre complément a le jeton d’accès, l’ID de l’élément et l’URL de l’API REST, il peut transmettre ces informations à un service principal qui appelle l’API REST, ou l’appeler directement à l’aide d’AJAX. L’exemple suivant appelle l’API REST de courrier Outlook pour obtenir le message actuel.

> [!IMPORTANT]
> Pour les déploiements Exchange locaux, les demandes côté client utilisant AJAX ou des bibliothèques similaires échouent, car CORS n’est pas pris en charge dans cette configuration de serveur.

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  const itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  const getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    const subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>Voir aussi

- Pour obtenir un exemple qui appelle les API REST à partir d’un complément Outlook, reportez-vous à [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) sur GitHub.
- Les API REST Outlook sont également disponibles via le point de terminaison Microsoft Graph, mais il existe quelques différences clés, notamment sur la façon dont votre complément obtient un jeton d’accès. Pour plus d’informations, reportez-vous à [API REST Outlook via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).
