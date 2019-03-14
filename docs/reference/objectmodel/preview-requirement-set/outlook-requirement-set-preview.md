---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: b1a3f5c675b2bcb43003ad15b3358e3febd80260
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512859"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="08f44-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="08f44-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="08f44-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="08f44-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="08f44-104">Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="08f44-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="08f44-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="08f44-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="08f44-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="08f44-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="08f44-107">La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.</span><span class="sxs-lookup"><span data-stu-id="08f44-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="08f44-108">Vous devrez également participer au [programme Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="08f44-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="08f44-109">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="08f44-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="08f44-110">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="08f44-110">Features in preview</span></span>

<span data-ttu-id="08f44-111">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="08f44-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="08f44-112">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="08f44-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="08f44-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="08f44-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="08f44-114">Ajout d’un nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="08f44-114">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="08f44-115">Cette valeur est utilisée pour annuler l’exécution d’un événement.</span><span class="sxs-lookup"><span data-stu-id="08f44-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="08f44-116">**Disponible dans** : Outlook sur le web (Classique)</span><span class="sxs-lookup"><span data-stu-id="08f44-116">**Available in**: Outlook on the web (Classic)</span></span>

### <a name="attachments"></a><span data-ttu-id="08f44-117">Pièces jointes</span><span class="sxs-lookup"><span data-stu-id="08f44-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="08f44-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="08f44-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="08f44-119">Ajout d’un nouvel objet représentant le contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="08f44-119">AttachmentContent - Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="08f44-120">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-120">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="08f44-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="08f44-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="08f44-122">Ajout d’une nouvelle méthode qui vous permet de joindre un fichier représenté par une chaîne encodée en base 64 à un message ou à un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="08f44-122">Office.context.mailbox.item.addFileAttachmentFromBase64Async - Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="08f44-123">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-123">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="08f44-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="08f44-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent)

<span data-ttu-id="08f44-125">Ajout d’une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.</span><span class="sxs-lookup"><span data-stu-id="08f44-125">Office.context.mailbox.item.getAttachmentContentAsync - Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="08f44-126">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-126">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>[<span data-ttu-id="08f44-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="08f44-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails)

<span data-ttu-id="08f44-128">Ajout d’une nouvelle méthode qui obtient les pièces jointes d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="08f44-128">Office.context.mailbox.item.getAttachmentsAsync - Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="08f44-129">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-129">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="08f44-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="08f44-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="08f44-131">Ajout d’une nouvelle énumération qui spécifie la mise en forme qui s’applique au contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="08f44-131">Office.MailboxEnums.AttachmentContentFormat - Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="08f44-132">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-132">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="08f44-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="08f44-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="08f44-134">Ajout d’une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.</span><span class="sxs-lookup"><span data-stu-id="08f44-134">Office.MailboxEnums.AttachmentStatus - Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="08f44-135">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-135">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="08f44-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="08f44-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="08f44-137">Ajout de l’événement `AttachmentsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="08f44-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="08f44-138">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-138">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="delegate-access"></a><span data-ttu-id="08f44-139">Accès délégué</span><span class="sxs-lookup"><span data-stu-id="08f44-139">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="08f44-140">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="08f44-140">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="08f44-141">Ajout d’un nouvel objet qui représente les propriétés d’un élément rendez-vous ou message dans un dossier, un calendrier ou une boîte aux lettres partagés.</span><span class="sxs-lookup"><span data-stu-id="08f44-141">SharedProperties - Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="08f44-142">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-142">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="08f44-143">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="08f44-143">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="08f44-144">Ajout d’une nouvelle méthode qui obtient un objet qui représente les sharedProperties d’un élément rendez-vous ou message.</span><span class="sxs-lookup"><span data-stu-id="08f44-144">Office.context.mailbox.item.getSharedPropertiesAsync - Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="08f44-145">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-145">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="08f44-146">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="08f44-146">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="08f44-147">Ajout d’une nouvelle énumération d’indicateur binaire qui spécifie les autorisations accordées aux délégués.</span><span class="sxs-lookup"><span data-stu-id="08f44-147">Office.MailboxEnums.DelegatePermissions - Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="08f44-148">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-148">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="08f44-149">Élément de manifeste SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="08f44-149">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="08f44-150">Ajout d’un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="08f44-150">SupportsSharedFolders manifest element - Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="08f44-151">Définit si le complément est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="08f44-151">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="08f44-152">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-152">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="enhanced-location"></a><span data-ttu-id="08f44-153">Emplacement amélioré</span><span class="sxs-lookup"><span data-stu-id="08f44-153">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="08f44-154">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="08f44-154">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="08f44-155">Ajout d’un nouvel objet représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="08f44-155">EnhancedLocation - Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="08f44-156">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-156">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="08f44-157">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="08f44-157">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="08f44-158">Ajout d’un nouvel objet représentant un emplacement.</span><span class="sxs-lookup"><span data-stu-id="08f44-158">LocationDetails - Added a new object that represents a location.</span></span> <span data-ttu-id="08f44-159">En lecture seule.</span><span class="sxs-lookup"><span data-stu-id="08f44-159">Read only.</span></span>

<span data-ttu-id="08f44-160">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-160">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="08f44-161">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="08f44-161">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="08f44-162">Ajout d’un nouvel objet représentant l’ID d’un emplacement.</span><span class="sxs-lookup"><span data-stu-id="08f44-162">LocationIdentifier - Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="08f44-163">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-163">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="08f44-164">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="08f44-164">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation)

<span data-ttu-id="08f44-165">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="08f44-165">Office.context.mailbox.item.enhancedLocation - Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="08f44-166">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-166">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="08f44-167">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="08f44-167">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="08f44-168">Ajout d’une nouvelle énumération qui spécifie le type d’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="08f44-168">Office.MailboxEnums.LocationType - Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="08f44-169">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-169">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="08f44-170">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="08f44-170">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="08f44-171">Ajout de l’événement `EnhancedLocationsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="08f44-171">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="08f44-172">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-172">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="08f44-173">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="08f44-173">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="08f44-174">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="08f44-174">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="08f44-175">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="08f44-175">Office.context.mailbox.item.getInitializationContextAsync - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="08f44-176">**Disponible dans** : Office 2019 pour Windows (abonnement Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="08f44-176">**Available in**: Office 2019 for Windows (Office 365 subscription), Outlook on the web (Classic)</span></span>

### <a name="internet-headers"></a><span data-ttu-id="08f44-177">En-têtes Internet</span><span class="sxs-lookup"><span data-stu-id="08f44-177">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="08f44-178">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="08f44-178">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="08f44-179">Ajout d’un nouvel objet représentant les en-têtes Internet d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="08f44-179">InternetHeaders - Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="08f44-180">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-180">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="08f44-181">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="08f44-181">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders)

<span data-ttu-id="08f44-182">Ajout d’une nouvelle propriété représentant les en-têtes Internet d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="08f44-182">Office.context.mailbox.item.internetHeaders - Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="08f44-183">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-183">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="office-theme"></a><span data-ttu-id="08f44-184">Thème Office</span><span class="sxs-lookup"><span data-stu-id="08f44-184">Office Theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="08f44-185">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="08f44-185">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="08f44-186">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="08f44-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="08f44-187">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-187">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="08f44-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="08f44-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="08f44-189">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="08f44-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="08f44-190">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="08f44-190">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="sso"></a><span data-ttu-id="08f44-191">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="08f44-191">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasynchttpsdocsmicrosoftcomofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="08f44-192">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="08f44-192">Office.context.auth.getAccessTokenAsync</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="08f44-193">Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="08f44-193">Office.context.auth.getAccessTokenAsync - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="08f44-194">**Disponible dans** : Outlook 2019 pour Windows (abonnement Office 365), Outlook 2019 pour Mac, Outlook sur le web (Office 365 et Outlook.com), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="08f44-194">**Available in**: Outlook 2019 for Windows (Office 365 subscription), Outlook 2019 for Mac, Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="08f44-195">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="08f44-195">See also</span></span>

- [<span data-ttu-id="08f44-196">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="08f44-196">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="08f44-197">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="08f44-197">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="08f44-198">Prise en main</span><span class="sxs-lookup"><span data-stu-id="08f44-198">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
