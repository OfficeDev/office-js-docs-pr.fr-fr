---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: 4365dab3d8dd1ddb876536b3030926d68a89ac49
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605672"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="ca507-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="ca507-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="ca507-103">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="ca507-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ca507-104">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="ca507-104">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="ca507-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="ca507-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="ca507-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="ca507-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="ca507-107">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="ca507-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="ca507-108">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="ca507-108">Features in preview</span></span>

<span data-ttu-id="ca507-109">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="ca507-109">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="ca507-110">Ajouter à l’envoi</span><span class="sxs-lookup"><span data-stu-id="ca507-110">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="ca507-111">Office. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="ca507-111">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="ca507-112">Ajout d’une nouvelle fonction à `Body` l’objet qui ajoute des données à la fin du corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="ca507-112">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="ca507-113">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ca507-113">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="ca507-114">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="ca507-114">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="ca507-115">Ajout d’un nouvel élément au manifeste dans lequel `AppendOnSend` l’autorisation étendue doit être incluse dans la collection des autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="ca507-115">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="ca507-116">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ca507-116">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="ca507-117">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="ca507-117">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="ca507-118">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="ca507-118">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="ca507-119">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="ca507-119">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="ca507-120">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="ca507-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="ca507-121">Thème Office</span><span class="sxs-lookup"><span data-stu-id="ca507-121">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="ca507-122">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="ca507-122">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="ca507-123">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="ca507-123">Added ability to get Office theme.</span></span>

<span data-ttu-id="ca507-124">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ca507-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="ca507-125">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="ca507-125">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="ca507-126">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="ca507-126">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="ca507-127">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ca507-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="ca507-128">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="ca507-128">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="ca507-129">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="ca507-129">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="ca507-130">Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ca507-130">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="ca507-131">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="ca507-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="ca507-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ca507-132">See also</span></span>

- [<span data-ttu-id="ca507-133">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="ca507-133">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="ca507-134">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="ca507-134">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="ca507-135">Prise en main</span><span class="sxs-lookup"><span data-stu-id="ca507-135">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="ca507-136">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="ca507-136">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
