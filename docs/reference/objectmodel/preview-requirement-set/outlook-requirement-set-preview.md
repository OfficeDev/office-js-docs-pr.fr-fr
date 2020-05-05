---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 05/04/2020
localization_priority: Normal
ms.openlocfilehash: 433abf5ea15cfe9ee7cefa579b1ad965fdeb93cf
ms.sourcegitcommit: 800dacf0399465318489c9d949e259b5cf0f81ca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/05/2020
ms.locfileid: "44022182"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="1fbbc-103">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="1fbbc-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1fbbc-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="1fbbc-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="1fbbc-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="1fbbc-108">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="1fbbc-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="1fbbc-109">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="1fbbc-109">Features in preview</span></span>

<span data-ttu-id="1fbbc-110">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-110">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="1fbbc-111">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="1fbbc-111">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="1fbbc-112">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="1fbbc-112">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="1fbbc-113">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-113">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="1fbbc-114">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="1fbbc-115">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="1fbbc-115">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="1fbbc-116">Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-116">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="1fbbc-117">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="1fbbc-118">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="1fbbc-118">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="1fbbc-119">Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-119">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="1fbbc-120">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="1fbbc-121">Office. Context. Mailbox. Item. Sensitivity</span><span class="sxs-lookup"><span data-stu-id="1fbbc-121">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="1fbbc-122">Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-122">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="1fbbc-123">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-123">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="1fbbc-124">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="1fbbc-124">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="1fbbc-125">Ajout d’une nouvelle `AppointmentSensitivityType` énumération qui représente les options de critère de diffusion disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-125">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="1fbbc-126">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-126">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="1fbbc-127">Ajouter à l’envoi</span><span class="sxs-lookup"><span data-stu-id="1fbbc-127">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="1fbbc-128">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="1fbbc-128">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="1fbbc-129">Ajout d’une nouvelle fonction à `Body` l’objet qui ajoute des données à la fin du corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-129">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="1fbbc-130">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="1fbbc-131">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="1fbbc-131">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="1fbbc-132">Ajout d’un nouvel élément au manifeste dans lequel `AppendOnSend` l’autorisation étendue doit être incluse dans la collection des autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-132">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="1fbbc-133">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-133">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="1fbbc-134">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="1fbbc-134">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="1fbbc-135">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="1fbbc-135">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="1fbbc-136">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="1fbbc-136">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="1fbbc-137">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="1fbbc-138">Signature de courrier électronique</span><span class="sxs-lookup"><span data-stu-id="1fbbc-138">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="1fbbc-139">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="1fbbc-139">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="1fbbc-140">Ajout d’une nouvelle fonction à `Body` l’objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-140">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="1fbbc-141">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-141">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="1fbbc-142">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="1fbbc-142">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="1fbbc-143">Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-143">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="1fbbc-144">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-144">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="1fbbc-145">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="1fbbc-145">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="1fbbc-146">Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-146">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="1fbbc-147">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-147">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="1fbbc-148">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="1fbbc-148">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="1fbbc-149">Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-149">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="1fbbc-150">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-150">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="1fbbc-151">Office. MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="1fbbc-151">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="1fbbc-152">Ajout d’une nouvelle `ComposeType` énumération disponible en mode composition.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-152">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="1fbbc-153">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-153">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="1fbbc-154">Thème Office</span><span class="sxs-lookup"><span data-stu-id="1fbbc-154">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="1fbbc-155">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="1fbbc-155">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="1fbbc-156">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-156">Added ability to get Office theme.</span></span>

<span data-ttu-id="1fbbc-157">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-157">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="1fbbc-158">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="1fbbc-158">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="1fbbc-159">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-159">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="1fbbc-160">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-160">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="online-meeting-provider-integration"></a><span data-ttu-id="1fbbc-161">Intégration des fournisseurs de réunions en ligne</span><span class="sxs-lookup"><span data-stu-id="1fbbc-161">Online meeting provider integration</span></span>

<span data-ttu-id="1fbbc-162">Prise en charge supplémentaire de l’intégration des réunions en ligne dans les compléments Outlook Mobile. Pour en savoir plus, reportez-vous à la rubrique [créer un complément Outlook Mobile pour un fournisseur de réunion en ligne](../../../outlook/online-meeting.md) .</span><span class="sxs-lookup"><span data-stu-id="1fbbc-162">Added support for online-meeting integration in Outlook mobile add-ins. See [Create an Outlook mobile add-in for an online-meeting provider](../../../outlook/online-meeting.md) to learn more.</span></span>

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[<span data-ttu-id="1fbbc-163">Point d’extension MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1fbbc-163">MobileOnlineMeetingCommandSurface extension point</span></span>](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

<span data-ttu-id="1fbbc-164">Ajout `MobileOnlineMeetingCommandSurface` du point d’extension au manifeste.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-164">Added `MobileOnlineMeetingCommandSurface` extension point to manifest.</span></span> <span data-ttu-id="1fbbc-165">Il définit l’intégration de la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-165">It defines the online meeting integration.</span></span>

<span data-ttu-id="1fbbc-166">**Disponible dans**: Outlook sur Android (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-166">**Available in**: Outlook on Android (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="1fbbc-167">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="1fbbc-167">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="1fbbc-168">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="1fbbc-168">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="1fbbc-169">Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1fbbc-169">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="1fbbc-170">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="1fbbc-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="1fbbc-171">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1fbbc-171">See also</span></span>

- [<span data-ttu-id="1fbbc-172">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="1fbbc-172">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="1fbbc-173">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="1fbbc-173">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="1fbbc-174">Prise en main</span><span class="sxs-lookup"><span data-stu-id="1fbbc-174">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="1fbbc-175">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="1fbbc-175">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
