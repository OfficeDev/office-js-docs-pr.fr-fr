---
title: Outlook conditions requises de l’API du module de prévisualisation du add-in
description: Fonctionnalités et API actuellement en prévisualisation pour Outlook de recherche.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: c7ca92e6a30f3109baff5721ae4e9930ef23dc56
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/09/2021
ms.locfileid: "52854010"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="f6dce-103">Outlook conditions requises de l’API du module de prévisualisation du add-in</span><span class="sxs-lookup"><span data-stu-id="f6dce-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="f6dce-104">Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.</span><span class="sxs-lookup"><span data-stu-id="f6dce-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f6dce-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="f6dce-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="f6dce-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="f6dce-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="f6dce-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="f6dce-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="f6dce-108">Vous pourrez peut-être afficher un aperçu des fonctionnalités Outlook sur le web en configurant la version ciblée [sur votre Microsoft 365 client.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="f6dce-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="f6dce-109">« Configurer l’accès à l’aperçu » est indiqué sur cette page pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="f6dce-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="f6dce-110">Pour d’autres fonctionnalités, vous pouvez demander l’accès aux bits d’aperçu pour Outlook sur le web à l’aide de votre compte Microsoft 365 en complétant et en envoyant ce [formulaire.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="f6dce-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="f6dce-111">« Demander l’accès en prévisualisation » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="f6dce-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="f6dce-112">L’ensemble de conditions requises de prévisualisation inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="f6dce-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="f6dce-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="f6dce-113">Features in preview</span></span>

<span data-ttu-id="f6dce-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="f6dce-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="f6dce-115">Activation de compléments sur des éléments protégés par la Gestion des droits de l’information (IRM)</span><span class="sxs-lookup"><span data-stu-id="f6dce-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="f6dce-116">Les add-ins peuvent désormais être activés sur les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="f6dce-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="f6dce-117">Pour activer cette fonctionnalité, un administrateur client doit activer le droit d’utilisation en paramètres de stratégie personnalisée Autoriser l’accès par programme `OBJMODEL` dans Office. </span><span class="sxs-lookup"><span data-stu-id="f6dce-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="f6dce-118">Pour plus [d’informations, voir droits d’utilisation et descriptions.](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions)</span><span class="sxs-lookup"><span data-stu-id="f6dce-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="f6dce-119">**Disponible dans**: Outlook sur Windows, à partir de la build 13229.10000 (connectée à Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="f6dce-120">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="f6dce-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="f6dce-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="f6dce-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="f6dce-122">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée d’un rendez-vous en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="f6dce-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="f6dce-123">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="f6dce-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="f6dce-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="f6dce-125">Ajout d’un nouvel objet qui représente la sensibilité d’un rendez-vous en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="f6dce-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="f6dce-126">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="f6dce-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="f6dce-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="f6dce-128">Ajout d’une nouvelle propriété qui représente si un rendez-vous est un événement d’une journée.</span><span class="sxs-lookup"><span data-stu-id="f6dce-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="f6dce-129">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="f6dce-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="f6dce-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="f6dce-131">Ajout d’une nouvelle propriété qui représente la sensibilité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f6dce-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="f6dce-132">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="f6dce-133">Office. MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="f6dce-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="f6dce-134">Ajout d’une nouvelle enum `AppointmentSensitivityType` qui représente les options de sensibilité disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f6dce-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="f6dce-135">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="f6dce-136">Activation basée sur un événement</span><span class="sxs-lookup"><span data-stu-id="f6dce-136">Event-based activation</span></span>

<span data-ttu-id="f6dce-137">Cette fonctionnalité a été publiée dans [l’ensemble de conditions requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="f6dce-137">This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="f6dce-138">Toutefois, des événements supplémentaires sont désormais disponibles en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="f6dce-138">However, additional events are now available in preview.</span></span> <span data-ttu-id="f6dce-139">Pour plus d’informations, voir [Événements pris en charge.](../../../outlook/autolaunch.md#supported-events)</span><span class="sxs-lookup"><span data-stu-id="f6dce-139">To learn more, see [Supported events](../../../outlook/autolaunch.md#supported-events).</span></span>

<span data-ttu-id="f6dce-140">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="f6dce-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="f6dce-141">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="f6dce-141">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="f6dce-142">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f6dce-142">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f6dce-143">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="f6dce-143">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="f6dce-144">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="f6dce-144">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="f6dce-145">Thème Office</span><span class="sxs-lookup"><span data-stu-id="f6dce-145">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="f6dce-146">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="f6dce-146">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="f6dce-147">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="f6dce-147">Added ability to get Office theme.</span></span>

<span data-ttu-id="f6dce-148">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="f6dce-149">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="f6dce-149">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f6dce-150">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="f6dce-150">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="f6dce-151">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="f6dce-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="f6dce-152">Données de session</span><span class="sxs-lookup"><span data-stu-id="f6dce-152">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="f6dce-153">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="f6dce-153">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="f6dce-154">Ajout d’un nouvel objet qui représente les données de session d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f6dce-154">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="f6dce-155">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="f6dce-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="f6dce-156">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="f6dce-156">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="f6dce-157">Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="f6dce-157">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="f6dce-158">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="f6dce-158">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="f6dce-159">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f6dce-159">See also</span></span>

- [<span data-ttu-id="f6dce-160">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="f6dce-160">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="f6dce-161">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="f6dce-161">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f6dce-162">Prise en main</span><span class="sxs-lookup"><span data-stu-id="f6dce-162">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="f6dce-163">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="f6dce-163">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
