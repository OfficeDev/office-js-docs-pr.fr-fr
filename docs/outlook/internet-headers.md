---
title: Obtenir et définir des en-têtes Internet
description: Comment obtenir et définir des en-têtes Internet sur un message dans un complément Outlook.
ms.date: 11/04/2019
localization_priority: Normal
ms.openlocfilehash: d7f38b7564683ce51ed0bd840480b4a8b2040bf6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166187"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="8b00b-103">Obtenir et définir des en-têtes Internet sur un message dans un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="8b00b-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="8b00b-104">Arrière-plan</span><span class="sxs-lookup"><span data-stu-id="8b00b-104">Background</span></span>

<span data-ttu-id="8b00b-105">Une exigence courante dans le développement des compléments Outlook est le stockage des propriétés personnalisées associées à un complément à différents niveaux.</span><span class="sxs-lookup"><span data-stu-id="8b00b-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="8b00b-106">À l’actuelle, les propriétés personnalisées sont stockées au niveau de l’élément ou de la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="8b00b-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="8b00b-107">Niveau de l’élément : pour les propriétés qui s’appliquent à un élément spécifique, utilisez l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) .</span><span class="sxs-lookup"><span data-stu-id="8b00b-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="8b00b-108">Par exemple, stockez un code client associé à la personne qui a envoyé le message électronique.</span><span class="sxs-lookup"><span data-stu-id="8b00b-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="8b00b-109">Niveau de la boîte aux lettres : pour les propriétés qui s’appliquent à tous les éléments de courrier dans la boîte aux lettres de l’utilisateur, utilisez l’objet [RoamingSettings](/javascript/api/outlook/office.roamingsettings) .</span><span class="sxs-lookup"><span data-stu-id="8b00b-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="8b00b-110">Par exemple, stockez la préférence d’un utilisateur pour afficher la température dans une mise à l’horizontale particulière.</span><span class="sxs-lookup"><span data-stu-id="8b00b-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="8b00b-111">Les deux types de propriétés ne sont pas conservés après que l’élément a quitté le serveur Exchange, de sorte que les destinataires du courrier électronique ne peuvent pas obtenir les propriétés définies sur l’élément.</span><span class="sxs-lookup"><span data-stu-id="8b00b-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="8b00b-112">Par conséquent, les développeurs ne peuvent pas accéder à ces paramètres ou à d’autres propriétés MIME pour permettre de meilleurs scénarios de lecture.</span><span class="sxs-lookup"><span data-stu-id="8b00b-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="8b00b-113">Bien qu’il existe un moyen de définir les en-têtes Internet par le biais de demandes EWS, dans certains scénarios, la demande EWS ne fonctionnera pas.</span><span class="sxs-lookup"><span data-stu-id="8b00b-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="8b00b-114">Par exemple, en mode composition sur le bureau Outlook, l’ID d’élément n’est pas `saveAsync` synchronisé en mode mis en cache.</span><span class="sxs-lookup"><span data-stu-id="8b00b-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="8b00b-115">Pour en savoir plus sur l’utilisation de ces options, consultez la rubrique [obtenir et définir des métadonnées de complément pour un complément Outlook](metadata-for-an-outlook-add-in.md) .</span><span class="sxs-lookup"><span data-stu-id="8b00b-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="8b00b-116">Objectif de l’API des en-têtes Internet</span><span class="sxs-lookup"><span data-stu-id="8b00b-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="8b00b-117">Introduit dans l’ensemble de conditions requises 1,8, les API d’en-têtes Internet permettent aux développeurs d’effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="8b00b-117">Introduced in requirement set 1.8, the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="8b00b-118">Informations de marquage sur un courrier électronique qui persistent une fois qu’il a quitté Exchange sur tous les clients.</span><span class="sxs-lookup"><span data-stu-id="8b00b-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="8b00b-119">Lire les informations d’un e-mail qui persistent après que le courrier électronique a quitté Exchange sur tous les clients dans les scénarios de lecture de messagerie.</span><span class="sxs-lookup"><span data-stu-id="8b00b-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="8b00b-120">Accéder à l’intégralité de l’en-tête MIME du courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="8b00b-120">Access the entire MIME header of the email.</span></span>

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="8b00b-121">Définir des en-têtes Internet lors de la composition d’un message</span><span class="sxs-lookup"><span data-stu-id="8b00b-121">Set internet headers while composing a message</span></span>

<span data-ttu-id="8b00b-122">Essayez d’utiliser la propriété [Item. internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) pour gérer les en-têtes Internet personnalisés que vous placez sur le message en cours en mode composition.</span><span class="sxs-lookup"><span data-stu-id="8b00b-122">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="8b00b-123">Exemple de définition, d’obtention et de suppression d’en-têtes personnalisés</span><span class="sxs-lookup"><span data-stu-id="8b00b-123">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="8b00b-124">L’exemple suivant montre comment définir, obtenir et supprimer des en-têtes personnalisés.</span><span class="sxs-lookup"><span data-stu-id="8b00b-124">The following example shows how to set, get, and remove custom headers.</span></span>

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

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="8b00b-125">Obtenir des en-têtes Internet lors de la lecture d’un message</span><span class="sxs-lookup"><span data-stu-id="8b00b-125">Get internet headers while reading a message</span></span>

<span data-ttu-id="8b00b-126">Essayez d’appeler [Item. getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) pour obtenir les en-têtes Internet sur le message actif en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="8b00b-126">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="8b00b-127">Obtenir les préférences de l’expéditeur à partir des en-têtes MIME actuels-exemple</span><span class="sxs-lookup"><span data-stu-id="8b00b-127">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="8b00b-128">En vous appuyant sur l’exemple de la section précédente, le code suivant montre comment obtenir les préférences de l’expéditeur à partir des en-têtes MIME de l’e-mail actuel.</span><span class="sxs-lookup"><span data-stu-id="8b00b-128">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

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
> <span data-ttu-id="8b00b-129">Cet exemple fonctionne pour des cas simples.</span><span class="sxs-lookup"><span data-stu-id="8b00b-129">This sample works for simple cases.</span></span> <span data-ttu-id="8b00b-130">Pour une extraction plus complexe des informations (par exemple, des en-têtes à plusieurs instances ou des valeurs pliées, comme décrit dans la [norme RFC 2822](https://tools.ietf.org/html/rfc2822)), essayez d’utiliser une bibliothèque d’analyse MIME appropriée.</span><span class="sxs-lookup"><span data-stu-id="8b00b-130">For more complex information retrieval (e.g., multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="see-also"></a><span data-ttu-id="8b00b-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8b00b-131">See also</span></span>

- [<span data-ttu-id="8b00b-132">Obtenir et définir des métadonnées de complément pour un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="8b00b-132">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
