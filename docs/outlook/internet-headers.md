---
title: Obtenir et définir des en-têtes Internet
description: Comment obtenir et définir des en-têtes Internet sur un message dans un Outlook de recherche.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 39e328f26ca849a95cf359b31480db5a1ca1830c80f4c414e34bb07657fe9b75
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089481"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Obtenir et définir des en-têtes Internet sur un message dans un Outlook de recherche

## <a name="background"></a>Contexte

Une exigence courante dans Outlook développement de add-ins consiste à stocker les propriétés personnalisées associées à un add-in à différents niveaux. Actuellement, les propriétés personnalisées sont stockées au niveau de l’élément ou de la boîte aux lettres.

- Niveau d’élément : pour les propriétés qui s’appliquent à un élément spécifique, utilisez [l’objet CustomProperties.](/javascript/api/outlook/office.customproperties) Par exemple, stockez un code client associé à la personne qui a envoyé le courrier électronique.
- Niveau de boîte aux lettres : pour les propriétés qui s’appliquent à tous les éléments de messagerie de la boîte aux lettres de l’utilisateur, utilisez [l’objet RoamingSettings.](/javascript/api/outlook/office.roamingsettings) Par exemple, stockez les préférences d’un utilisateur pour afficher la température dans une échelle particulière.

Les deux types de propriétés ne sont pas conservés après que l’élément a quitté le serveur Exchange afin que les destinataires de courrier ne peuvent pas obtenir de propriétés définies sur l’élément. Par conséquent, les développeurs ne peuvent pas accéder à ces paramètres ou à d’autres propriétés MIME pour permettre de meilleurs scénarios de lecture.

Bien qu’il soit possible de définir les en-têtes Internet par le biais de demandes EWS, dans certains scénarios, l’utilisation d’une demande EWS ne fonctionne pas. Par exemple, en mode composition Outlook bureau, l’ID d’élément n’est pas synchronisé en  `saveAsync`   mode mis en cache.

> [!TIP]
> Voir [Obtenir et définir des métadonnées](metadata-for-an-outlook-add-in.md) de Outlook pour en savoir plus sur l’utilisation de ces options.

## <a name="purpose-of-the-internet-headers-api"></a>Objectif de l’API d’en-têtes Internet

Introduites dans [l’ensemble de conditions requises 1.8,](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)les API d’en-têtes Internet permettent aux développeurs de :

- Marquez les informations d’un e-mail qui persistent après son Exchange tous les clients.
- Lire les informations d’un e-mail qui ont persisté après qu’il a été laissé Exchange tous les clients dans les scénarios de lecture de courrier électronique.
- Accéder à l’intégralité de l’en-tête MIME de l’e-mail.

![Diagramme des en-têtes Internet. Texte : l’utilisateur 1 envoie un e-mail. Le add-in gère les en-têtes Internet personnalisés pendant que l’utilisateur compose des messages électroniques. L’utilisateur 2 reçoit le message électronique. Le add-in obtient les en-têtes Internet provenant du courrier électronique reçu, puis il parcourt et utilise des en-têtes personnalisés.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Définir des en-têtes Internet lors de la composition d’un message

Essayez d’utiliser [la propriété item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetHeaders) pour gérer les en-têtes Internet personnalisés que vous placez sur le message actuel en mode Composition.

### <a name="set-get-and-remove-custom-headers-example"></a>Exemple de définir, d’obtenir et de supprimer des en-têtes personnalisés

L’exemple suivant montre comment définir, obtenir et supprimer des en-têtes personnalisés.

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-preferred-fruit": "orange", "x-preferred-vegetable": "broccoli", "x-best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["x-preferred-fruit", "x-preferred-vegetable", "x-best-vegetable", "x-nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["x-best-vegetable", "x-nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"x-best-vegetable":"spinach","x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
*/
```

## <a name="get-internet-headers-while-reading-a-message"></a>Obtenir des en-têtes Internet lors de la lecture d’un message

Essayez [d’appeler item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getAllInternetHeadersAsync_options__callback_) pour obtenir des en-têtes Internet sur le message actuel en mode lecture.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Obtenir les préférences de l’expéditeur à partir de l’exemple d’en-têtes MIME actuels

En s’axant sur l’exemple de la section précédente, le code suivant montre comment obtenir les préférences de l’expéditeur à partir des en-têtes MIME de l’e-mail actuel.

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/x-preferred-fruit:.*/gim)[0].slice(19));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/x-preferred-vegetable:.*/gim)[0].slice(23));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> Cet exemple fonctionne pour des cas simples. Pour une récupération d’informations plus complexe (par exemple, des en-têtes à instances multiples ou des valeurs pliées comme décrit dans [la RFC 2822),](https://tools.ietf.org/html/rfc2822)essayez d’utiliser une bibliothèque d’assinage MIME appropriée.

## <a name="recommended-practices"></a>Pratiques recommandées

Actuellement, les en-têtes Internet sont une ressource finie sur la boîte aux lettres d’un utilisateur. Lorsque le quota est épuisé, vous ne pouvez plus créer d’en-têtes Internet sur cette boîte aux lettres, ce qui peut entraîner un comportement inattendu de la part des clients qui s’en appuient pour fonctionner.

Appliquez les instructions suivantes lorsque vous créez des en-têtes Internet dans votre application.

- Créez le nombre minimal d’en-têtes requis.
- Nommez les en-têtes afin de pouvoir réutiliser et mettre à jour leurs valeurs ultérieurement. En tant que tel, évitez d’nommer les en-têtes de manière variable (par exemple, en fonction de l’entrée utilisateur, de l’timestamp, etc.).

## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des métadonnées de complément pour un complément Outlook](metadata-for-an-outlook-add-in.md)
