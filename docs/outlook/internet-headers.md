---
title: Obtenir et définir des en-têtes Internet
description: Comment obtenir et définir des en-têtes Internet sur un message dans un complément Outlook.
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8e4af70b24a96b8d00acc7ea4101acf53e2b71
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616027"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Obtenir et définir des en-têtes Internet sur un message dans un complément Outlook

## <a name="background"></a>Contexte

Une exigence courante dans le développement de compléments Outlook consiste à stocker les propriétés personnalisées associées à un complément à différents niveaux. À l’heure actuelle, les propriétés personnalisées sont stockées au niveau de l’élément ou de la boîte aux lettres.

- Niveau d’élément : pour les propriétés qui s’appliquent à un élément spécifique, utilisez l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) . Par exemple, stockez un code client associé à la personne qui a envoyé l’e-mail.
- Niveau boîte aux lettres : pour les propriétés qui s’appliquent à tous les éléments de messagerie de la boîte aux lettres de l’utilisateur, utilisez l’objet [RoamingSettings](/javascript/api/outlook/office.roamingsettings) . Par exemple, stockez la préférence d’un utilisateur pour afficher la température dans une échelle particulière.

Les deux types de propriétés ne sont pas conservés une fois que l’élément a quitté le serveur Exchange, de sorte que les destinataires de l’e-mail ne peuvent pas obtenir de propriétés définies sur l’élément. Par conséquent, les développeurs ne peuvent pas accéder à ces paramètres ou à d’autres propriétés MIME (Multipurpose Internet Mail Extensions) pour permettre de meilleurs scénarios de lecture.

Bien qu’il existe un moyen de définir les en-têtes Internet par le biais de requêtes EWS (Exchange Web Services), dans certains scénarios, l’établissement d’une requête EWS ne fonctionnera pas. Par exemple, en mode Composition sur le bureau Outlook, l’ID d’élément n’est pas synchronisé `saveAsync` en mode mis en cache.

> [!TIP]
> Pour en savoir plus sur l’utilisation de ces options, consultez [Obtenir et définir des métadonnées de complément pour un complément Outlook](metadata-for-an-outlook-add-in.md).

## <a name="purpose-of-the-internet-headers-api"></a>Objectif de l’API d’en-têtes Internet

Introduites dans [l’ensemble de conditions requises 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8), les API d’en-têtes Internet permettent aux développeurs de :

- Horodatage des informations sur un e-mail qui persiste une fois qu’il a quitté Exchange sur tous les clients.
- Lisez les informations sur un e-mail qui a persisté après que l’e-mail a quitté Exchange sur tous les clients dans des scénarios de lecture de courrier.
- Accédez à l’en-tête MIME entier de l’e-mail.

![Diagramme des en-têtes Internet. Texte : l’utilisateur 1 envoie un e-mail. Le complément gère les en-têtes Internet personnalisés pendant que l’utilisateur compose des e-mails. L’utilisateur 2 reçoit l’e-mail. Le complément obtient les en-têtes Internet de l’e-mail reçu, puis analyse et utilise des en-têtes personnalisés.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Définir des en-têtes Internet lors de la composition d’un message

Utilisez la propriété [item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) pour gérer les en-têtes Internet personnalisés que vous placez sur le message actuel en mode Compose.

### <a name="set-get-and-remove-custom-internet-headers-example"></a>Exemple de définition, d’obtention et de suppression d’en-têtes Internet personnalisés

L’exemple suivant montre comment définir, obtenir et supprimer des en-têtes Internet personnalisés.

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

Appelez [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) pour obtenir des en-têtes Internet sur le message actuel en mode lecture.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Obtenir les préférences de l’expéditeur à partir de l’exemple d’en-têtes MIME actuel

S’appuyant sur l’exemple de la section précédente, le code suivant montre comment obtenir les préférences de l’expéditeur à partir des en-têtes MIME de l’e-mail actuel.

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
> Cet exemple fonctionne pour les cas simples. Pour une récupération d’informations plus complexe (par exemple, des en-têtes multi-instances ou des valeurs pliées comme décrit dans [RFC 2822](https://tools.ietf.org/html/rfc2822)), essayez d’utiliser une bibliothèque d’analyse MIME appropriée.

## <a name="recommended-practices"></a>Pratiques recommandées

Actuellement, les en-têtes Internet sont une ressource finie sur la boîte aux lettres d’un utilisateur. Lorsque le quota est épuisé, vous ne pouvez pas créer d’en-têtes Internet supplémentaires sur cette boîte aux lettres, ce qui peut entraîner un comportement inattendu des clients qui s’appuient sur ce paramètre pour fonctionner.

Appliquez les instructions suivantes lorsque vous créez des en-têtes Internet dans votre complément.

- Créez le nombre minimal d’en-têtes requis. Le quota d’en-tête est basé sur la taille totale des en-têtes appliqués à un message. Dans Exchange Online, la limite d’en-tête est limitée à 256 Ko, tandis que dans un environnement Exchange local, la limite est déterminée par l’administrateur de votre organisation. Pour plus d’informations sur les limites d’en-tête, consultez [Exchange Online limites des messages](/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits) et [Exchange Server limites des messages](/exchange/mail-flow/message-size-limits).
- Nommez les en-têtes afin que vous puissiez réutiliser et mettre à jour leurs valeurs ultérieurement. Par conséquent, évitez d’nommer les en-têtes de manière variable (par exemple, en fonction de l’entrée utilisateur, de l’horodatage, etc.).

## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des métadonnées de complément pour un complément Outlook](metadata-for-an-outlook-add-in.md)
