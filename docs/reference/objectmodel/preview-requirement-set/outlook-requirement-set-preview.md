---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 24cad394f0f3ffb95a05a81ccb38ee4aa72a3797
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431065"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="cbecb-103">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="cbecb-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="cbecb-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="cbecb-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cbecb-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="cbecb-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="cbecb-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="cbecb-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="cbecb-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="cbecb-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="cbecb-108">Vous pouvez afficher un aperçu des fonctionnalités dans Outlook sur le Web en [configurant la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="cbecb-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="cbecb-109">« Configurer l’accès en aperçu » est indiqué sur cette page pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="cbecb-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="cbecb-110">Pour les autres fonctionnalités, vous pouvez demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="cbecb-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="cbecb-111">« Demander un accès en aperçu » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="cbecb-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="cbecb-112">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="cbecb-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="cbecb-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="cbecb-113">Features in preview</span></span>

<span data-ttu-id="cbecb-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="cbecb-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="cbecb-115">Activation des compléments sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)</span><span class="sxs-lookup"><span data-stu-id="cbecb-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="cbecb-116">Les compléments peuvent désormais être activés sur les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="cbecb-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="cbecb-117">Pour activer cette fonctionnalité, un administrateur client doit activer le droit d' `OBJMODEL` utilisation en définissant l’option autoriser la stratégie personnalisée d' **accès par programme** dans Office.</span><span class="sxs-lookup"><span data-stu-id="cbecb-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="cbecb-118">Pour plus d’informations [, voir droits et descriptions d’utilisation](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) .</span><span class="sxs-lookup"><span data-stu-id="cbecb-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="cbecb-119">**Disponible dans**: Outlook sur Windows, en commençant par Build 13229,10000 (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="cbecb-120">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="cbecb-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="cbecb-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="cbecb-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="cbecb-122">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="cbecb-123">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="cbecb-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="cbecb-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="cbecb-125">Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="cbecb-126">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="cbecb-127">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="cbecb-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="cbecb-128">Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.</span><span class="sxs-lookup"><span data-stu-id="cbecb-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="cbecb-129">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="cbecb-130">Office. Context. Mailbox. Item. Sensitivity</span><span class="sxs-lookup"><span data-stu-id="cbecb-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="cbecb-131">Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="cbecb-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="cbecb-132">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="cbecb-133">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="cbecb-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="cbecb-134">Ajout d’une nouvelle énumération `AppointmentSensitivityType` qui représente les options de critère de diffusion disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="cbecb-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="cbecb-135">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="cbecb-136">Ajouter à l'envoi</span><span class="sxs-lookup"><span data-stu-id="cbecb-136">Append on send</span></span>

<span data-ttu-id="cbecb-137">Pour en savoir plus sur l’utilisation de la fonctionnalité Ajout à l’envoi, consultez la rubrique [implémenter Append lors de l’envoi dans votre complément Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="cbecb-137">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="cbecb-138">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-138">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-)

<span data-ttu-id="cbecb-139">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute des données à la fin du corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-139">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="cbecb-140">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="cbecb-141">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="cbecb-141">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="cbecb-142">Ajout d’un nouvel élément au manifeste dans lequel l' `AppendOnSend` autorisation étendue doit être incluse dans la collection des autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="cbecb-142">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="cbecb-143">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="cbecb-144">Versions Async des `display` API</span><span class="sxs-lookup"><span data-stu-id="cbecb-144">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="cbecb-145">Office. Context. Mailbox. displayAppointmentFormAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-145">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="cbecb-146">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un rendez-vous existant.</span><span class="sxs-lookup"><span data-stu-id="cbecb-146">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="cbecb-147">Il s’agit de la version asynchrone de la `displayAppointmentForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="cbecb-147">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="cbecb-148">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="cbecb-149">Office. Context. Mailbox. displayMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-149">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="cbecb-150">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="cbecb-150">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="cbecb-151">Il s’agit de la version asynchrone de la `displayMessageForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="cbecb-151">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="cbecb-152">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-152">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="cbecb-153">Office. Context. Mailbox. displayNewAppointmentFormAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-153">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="cbecb-154">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="cbecb-154">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="cbecb-155">Il s’agit de la version asynchrone de la `displayNewAppointmentForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="cbecb-155">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="cbecb-156">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-156">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="cbecb-157">Office. Context. Mailbox. displayNewMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-157">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="cbecb-158">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="cbecb-158">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="cbecb-159">Il s’agit de la version asynchrone de la `displayNewMessageForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="cbecb-159">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="cbecb-160">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="cbecb-161">Office. Context. Mailbox. Item. displayReplyAllFormAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-161">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="cbecb-162">Ajout d’une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre à tous » en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="cbecb-162">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="cbecb-163">Il s’agit de la version asynchrone de la `displayReplyAllForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="cbecb-163">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="cbecb-164">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-164">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="cbecb-165">Office. Context. Mailbox. Item. displayReplyFormAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-165">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="cbecb-166">Ajout d’une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre » en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="cbecb-166">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="cbecb-167">Il s’agit de la version asynchrone de la `displayReplyForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="cbecb-167">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="cbecb-168">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-168">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="cbecb-169">Activation basée sur un événement</span><span class="sxs-lookup"><span data-stu-id="cbecb-169">Event-based activation</span></span>

<span data-ttu-id="cbecb-170">Prise en charge supplémentaire de la fonctionnalité d’activation basée sur un événement dans les compléments Outlook. Pour en savoir plus, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="cbecb-170">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="cbecb-171">Point d’extension LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="cbecb-171">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="cbecb-172">Ajout `LaunchEvent` de la prise en charge du point d’extension au manifeste.</span><span class="sxs-lookup"><span data-stu-id="cbecb-172">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="cbecb-173">Il configure les fonctionnalités d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="cbecb-173">It configures event-based activation functionality.</span></span>

<span data-ttu-id="cbecb-174">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="cbecb-174">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="cbecb-175">Élément de manifeste LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="cbecb-175">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="cbecb-176">Ajout `LaunchEvents` de l’élément à manifest.</span><span class="sxs-lookup"><span data-stu-id="cbecb-176">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="cbecb-177">Il prend en charge la configuration de la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="cbecb-177">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="cbecb-178">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="cbecb-178">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="cbecb-179">Élément de manifeste runtimes</span><span class="sxs-lookup"><span data-stu-id="cbecb-179">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="cbecb-180">Ajout de la prise en charge d’Outlook à l' `Runtimes` élément de manifeste.</span><span class="sxs-lookup"><span data-stu-id="cbecb-180">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="cbecb-181">Il fait référence aux fichiers HTML et JavaScript nécessaires à la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="cbecb-181">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="cbecb-182">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="cbecb-182">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="cbecb-183">Obtenir toutes les propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="cbecb-183">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="cbecb-184">CustomProperties. getAll</span><span class="sxs-lookup"><span data-stu-id="cbecb-184">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true#getall--)

<span data-ttu-id="cbecb-185">Ajout d’une nouvelle fonction à l' `CustomProperties` objet qui obtient toutes les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="cbecb-185">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="cbecb-186">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne), Outlook sur Mac (connecté à un abonnement Microsoft 365), Outlook sur Android, Outlook sur iOS</span><span class="sxs-lookup"><span data-stu-id="cbecb-186">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="cbecb-187">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="cbecb-187">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="cbecb-188">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-188">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="cbecb-189">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="cbecb-189">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="cbecb-190">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (classique)</span><span class="sxs-lookup"><span data-stu-id="cbecb-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="cbecb-191">Signature de courrier électronique</span><span class="sxs-lookup"><span data-stu-id="cbecb-191">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="cbecb-192">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-192">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="cbecb-193">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-193">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="cbecb-194">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="cbecb-195">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-195">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="cbecb-196">Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-196">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="cbecb-197">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="cbecb-198">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-198">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="cbecb-199">Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-199">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="cbecb-200">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-200">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="cbecb-201">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="cbecb-201">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="cbecb-202">Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-202">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="cbecb-203">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-203">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="cbecb-204">Office. MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="cbecb-204">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="cbecb-205">Ajout d’une nouvelle énumération `ComposeType` disponible en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-205">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="cbecb-206">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="cbecb-206">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="cbecb-207">Messages de notification avec actions</span><span class="sxs-lookup"><span data-stu-id="cbecb-207">Notification messages with actions</span></span>

<span data-ttu-id="cbecb-208">Cette fonctionnalité permet à votre complément d’inclure un message de notification avec une action personnalisée en plus de l’action **Ignorer** par défaut.</span><span class="sxs-lookup"><span data-stu-id="cbecb-208">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="cbecb-209">Office. NotificationMessageDetails. actions</span><span class="sxs-lookup"><span data-stu-id="cbecb-209">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="cbecb-210">Ajout d’une nouvelle propriété qui vous permet d’ajouter une `InsightMessage` notification avec une action personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cbecb-210">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="cbecb-211">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-211">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="cbecb-212">Office. NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="cbecb-212">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="cbecb-213">Ajout d’un nouvel objet dans lequel vous définissez une action personnalisée pour votre `InsightMessage` notification.</span><span class="sxs-lookup"><span data-stu-id="cbecb-213">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="cbecb-214">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-214">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="cbecb-215">Office. MailboxEnums. ActionType</span><span class="sxs-lookup"><span data-stu-id="cbecb-215">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="cbecb-216">Ajout d’une nouvelle énumération `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="cbecb-216">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="cbecb-217">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-217">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="cbecb-218">Office. MailboxEnums. ItemNotificationMessageType. InsightMessage</span><span class="sxs-lookup"><span data-stu-id="cbecb-218">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="cbecb-219">Ajout d’un nouveau type `InsightMessage` à l' `ItemNotificationMessageType` énumération.</span><span class="sxs-lookup"><span data-stu-id="cbecb-219">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="cbecb-220">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="cbecb-220">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="cbecb-221">Thème Office</span><span class="sxs-lookup"><span data-stu-id="cbecb-221">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="cbecb-222">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="cbecb-222">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="cbecb-223">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="cbecb-223">Added ability to get Office theme.</span></span>

<span data-ttu-id="cbecb-224">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-224">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="cbecb-225">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="cbecb-225">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="cbecb-226">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="cbecb-226">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="cbecb-227">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-227">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="cbecb-228">Données de session</span><span class="sxs-lookup"><span data-stu-id="cbecb-228">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="cbecb-229">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="cbecb-229">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="cbecb-230">Ajout d’un nouvel objet qui représente les données de session d’un élément.</span><span class="sxs-lookup"><span data-stu-id="cbecb-230">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="cbecb-231">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-231">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="cbecb-232">Office. Context. Mailbox. Item. sessionData</span><span class="sxs-lookup"><span data-stu-id="cbecb-232">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="cbecb-233">Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="cbecb-233">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="cbecb-234">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="cbecb-234">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="cbecb-235">Authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="cbecb-235">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="cbecb-236">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="cbecb-236">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="cbecb-237">Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="cbecb-237">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="cbecb-238">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur Mac (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne), Outlook sur le Web (classique)</span><span class="sxs-lookup"><span data-stu-id="cbecb-238">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="cbecb-239">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cbecb-239">See also</span></span>

- [<span data-ttu-id="cbecb-240">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="cbecb-240">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="cbecb-241">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="cbecb-241">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="cbecb-242">Prise en main</span><span class="sxs-lookup"><span data-stu-id="cbecb-242">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="cbecb-243">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="cbecb-243">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
