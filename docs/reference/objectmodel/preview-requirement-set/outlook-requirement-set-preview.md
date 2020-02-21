---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: c297904ff8343fd4c958c80b41170c5f2e93c739
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165502"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="f8fd4-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="f8fd4-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="f8fd4-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f8fd4-104">Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="f8fd4-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="f8fd4-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="f8fd4-107">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="f8fd4-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="f8fd4-108">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="f8fd4-108">Features in preview</span></span>

<span data-ttu-id="f8fd4-109">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="f8fd4-110">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="f8fd4-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="f8fd4-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f8fd4-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f8fd4-112">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="f8fd4-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="f8fd4-113">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="f8fd4-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="f8fd4-114">Thème Office</span><span class="sxs-lookup"><span data-stu-id="f8fd4-114">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="f8fd4-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="f8fd4-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="f8fd4-116">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="f8fd4-117">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f8fd4-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="f8fd4-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="f8fd4-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f8fd4-119">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="f8fd4-120">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f8fd4-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="f8fd4-121">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="f8fd4-121">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="f8fd4-122">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="f8fd4-122">OfficeRuntime.auth.getAccessToken</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="f8fd4-123">Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="f8fd4-123">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="f8fd4-124">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="f8fd4-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="f8fd4-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f8fd4-125">See also</span></span>

- [<span data-ttu-id="f8fd4-126">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="f8fd4-126">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="f8fd4-127">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="f8fd4-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f8fd4-128">Prise en main</span><span class="sxs-lookup"><span data-stu-id="f8fd4-128">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="f8fd4-129">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="f8fd4-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
