---
title: Obtenir et définir des en-têtes Internet
description: Comment obtenir et définir des en-têtes Internet sur un message dans un complément Outlook.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 1b6bdbbe77998ce92ea1b1b43874a32a30aa160a
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930287"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Obtenir et définir des en-têtes Internet sur un message dans un complément Outlook

## <a name="background"></a>Arrière-plan

Une exigence courante dans le développement des compléments Outlook est le stockage des propriétés personnalisées associées à un complément à différents niveaux. À l’actuelle, les propriétés personnalisées sont stockées au niveau de l’élément ou de la boîte aux lettres.

- Niveau de l’élément : pour les propriétés qui s’appliquent à un élément spécifique, utilisez l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) . Par exemple, stockez un code client associé à la personne qui a envoyé le message électronique.
- Niveau de la boîte aux lettres : pour les propriétés qui s’appliquent à tous les éléments de courrier dans la boîte aux lettres de l’utilisateur, utilisez l’objet [RoamingSettings](/javascript/api/outlook/office.roamingsettings) . Par exemple, stockez la préférence d’un utilisateur pour afficher la température dans une mise à l’horizontale particulière.

Les deux types de propriétés ne sont pas conservés après que l’élément a quitté le serveur Exchange, de sorte que les destinataires du courrier électronique ne peuvent pas obtenir les propriétés définies sur l’élément. Par conséquent, les développeurs ne peuvent pas accéder à ces paramètres ou à d’autres propriétés MIME pour permettre de meilleurs scénarios de lecture.

Bien qu’il existe un moyen de définir les en-têtes Internet par le biais de demandes EWS, dans certains scénarios, la demande EWS ne fonctionnera pas. Par exemple, en mode composition sur le bureau Outlook, l’ID d’élément n’est pas `saveAsync` synchronisé en mode mis en cache.

> [!TIP]
> Pour en savoir plus sur l’utilisation de ces options, consultez la rubrique [obtenir et définir des métadonnées de complément pour un complément Outlook](metadata-for-an-outlook-add-in.md) .

## <a name="purpose-of-the-internet-headers-api"></a>Objectif de l’API des en-têtes Internet

Introduit dans l' [ensemble de conditions requises 1,8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), les API d’en-têtes Internet permettent aux développeurs d’effectuer les opérations suivantes :

- Informations de marquage sur un courrier électronique qui persistent une fois qu’il a quitté Exchange sur tous les clients.
- Lire les informations d’un e-mail qui persistent après que le courrier électronique a quitté Exchange sur tous les clients dans les scénarios de lecture de messagerie.
- Accéder à l’intégralité de l’en-tête MIME du courrier électronique.

![Diagramme des en-têtes Internet. Text : l’utilisateur 1 envoie des courriers électroniques. Le complément gère les en-têtes Internet personnalisés pendant que l’utilisateur compose le courrier électronique. L’utilisateur 2 reçoit le courrier électronique. Le complément obtient les en-têtes Internet du courrier électronique reçu, puis analyse et utilise des en-têtes personnalisés.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Définir des en-têtes Internet lors de la composition d’un message

Essayez d’utiliser la propriété [Item. internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) pour gérer les en-têtes Internet personnalisés que vous placez sur le message en cours en mode composition.

### <a name="set-get-and-remove-custom-headers-example"></a>Exemple de définition, d’obtention et de suppression d’en-têtes personnalisés

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

Essayez d’appeler [Item. getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) pour obtenir les en-têtes Internet sur le message actif en mode lecture.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Obtenir les préférences de l’expéditeur à partir des en-têtes MIME actuels-exemple

En vous appuyant sur l’exemple de la section précédente, le code suivant montre comment obtenir les préférences de l’expéditeur à partir des en-têtes MIME de l’e-mail actuel.

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
> Cet exemple fonctionne pour des cas simples. Pour une extraction plus complexe des informations (par exemple, des en-têtes à plusieurs instances ou des valeurs pliées, comme décrit dans la [norme RFC 2822](https://tools.ietf.org/html/rfc2822)), essayez d’utiliser une bibliothèque d’analyse MIME appropriée.

## <a name="recommended-practices"></a>Pratiques recommandées

Actuellement, les en-têtes Internet sont une ressource finie sur la boîte aux lettres d’un utilisateur. Lorsque le quota est épuisé, vous ne pouvez plus créer d’en-têtes Internet supplémentaires sur cette boîte aux lettres, ce qui peut entraîner un comportement inattendu de la part des clients qui dépendent de cette fonctionnalité.

Appliquez les instructions suivantes lorsque vous créez des en-têtes Internet dans votre complément.

- Créez le nombre minimal d’en-têtes requis.
- Les en-têtes de nom afin que vous puissiez réutiliser et mettre à jour leurs valeurs ultérieurement. En tant que telle, évitez les en-têtes de nom de manière variable (par exemple, en fonction de l’entrée utilisateur, de l’horodatage, etc.).

## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des métadonnées de complément pour un complément Outlook](metadata-for-an-outlook-add-in.md)
