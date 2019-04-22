---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 04/17/2019
localization_priority: Priority
ms.openlocfilehash: 9a3d09a78a7644b3b26c345ba2588a1fae59c1eb
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914269"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="e5f52-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="e5f52-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="e5f52-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5f52-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e5f52-104">Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="e5f52-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="e5f52-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="e5f52-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="e5f52-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e5f52-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="e5f52-107">La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.</span><span class="sxs-lookup"><span data-stu-id="e5f52-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="e5f52-108">Vous devrez également participer au [programme Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="e5f52-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="e5f52-109">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="e5f52-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="e5f52-110">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="e5f52-110">Features in preview</span></span>

<span data-ttu-id="e5f52-111">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="e5f52-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="e5f52-112">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="e5f52-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="e5f52-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="e5f52-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="e5f52-114">Ajout d’un nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="e5f52-114">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="e5f52-115">Cette valeur est utilisée pour annuler l’exécution d’un événement.</span><span class="sxs-lookup"><span data-stu-id="e5f52-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="e5f52-116">**Disponible dans** : Outlook sur le web (Classique)</span><span class="sxs-lookup"><span data-stu-id="e5f52-116">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="attachments"></a><span data-ttu-id="e5f52-117">Pièces jointes</span><span class="sxs-lookup"><span data-stu-id="e5f52-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="e5f52-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="e5f52-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="e5f52-119">Ajout d’un nouvel objet représentant le contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="e5f52-119">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="e5f52-120">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-120">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="e5f52-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e5f52-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="e5f52-122">Ajout d’une nouvelle méthode qui vous permet de joindre un fichier représenté par une chaîne encodée en base 64 à un message ou à un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e5f52-122">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="e5f52-123">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-123">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="e5f52-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="e5f52-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="e5f52-125">Ajout d’une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.</span><span class="sxs-lookup"><span data-stu-id="e5f52-125">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="e5f52-126">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-126">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="e5f52-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="e5f52-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="e5f52-128">Ajout d’une nouvelle méthode qui obtient les pièces jointes d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="e5f52-128">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="e5f52-129">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-129">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="e5f52-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="e5f52-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="e5f52-131">Ajout d’une nouvelle énumération qui spécifie la mise en forme qui s’applique au contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="e5f52-131">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="e5f52-132">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-132">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="e5f52-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="e5f52-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="e5f52-134">Ajout d’une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.</span><span class="sxs-lookup"><span data-stu-id="e5f52-134">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="e5f52-135">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-135">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e5f52-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="e5f52-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e5f52-137">Ajout de l’événement `AttachmentsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="e5f52-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="e5f52-138">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-138">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="categories"></a><span data-ttu-id="e5f52-139">Catégories</span><span class="sxs-lookup"><span data-stu-id="e5f52-139">Categories</span></span>

<span data-ttu-id="e5f52-140">Dans Outlook, un utilisateur peut regrouper des messages et des rendez-vous à l’aide d’une catégorie pour leur appliquer un code de couleur.</span><span class="sxs-lookup"><span data-stu-id="e5f52-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="e5f52-141">L’utilisateur définit les catégories dans une liste sur sa boîte aux lettres principale.</span><span class="sxs-lookup"><span data-stu-id="e5f52-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="e5f52-142">Ils peuvent ensuite appliquer une ou plusieurs catégories à un élément.</span><span class="sxs-lookup"><span data-stu-id="e5f52-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="e5f52-143">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="e5f52-143">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="e5f52-144">Catégories</span><span class="sxs-lookup"><span data-stu-id="e5f52-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="e5f52-145">Ajout d’un nouvel objet représentant des catégories d’un élément.</span><span class="sxs-lookup"><span data-stu-id="e5f52-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="e5f52-146">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-146">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="e5f52-147">Détailscatégorie</span><span class="sxs-lookup"><span data-stu-id="e5f52-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="e5f52-148">Ajout un nouvel objet qui représente les détails d’une catégorie (son nom et la couleur associée).</span><span class="sxs-lookup"><span data-stu-id="e5f52-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="e5f52-149">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-149">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="e5f52-150">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="e5f52-150">masterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="e5f52-151">Ajout d’ un nouvel objet qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e5f52-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="e5f52-152">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-152">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="e5f52-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="e5f52-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="e5f52-154">Ajout d’ un nouveau propriétaire qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e5f52-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="e5f52-155">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-155">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="e5f52-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="e5f52-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="e5f52-157">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un élément.</span><span class="sxs-lookup"><span data-stu-id="e5f52-157">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e5f52-158">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-158">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="e5f52-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="e5f52-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="e5f52-160">Ajouté un nouvel enum qui spécifie les couleurs disponibles à associer à des catégories.</span><span class="sxs-lookup"><span data-stu-id="e5f52-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="e5f52-161">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-161">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="e5f52-162">Accès délégué</span><span class="sxs-lookup"><span data-stu-id="e5f52-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="e5f52-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="e5f52-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="e5f52-164">Ajout d’un nouvel objet qui représente les propriétés d’un élément rendez-vous ou message dans un dossier, un calendrier ou une boîte aux lettres partagés.</span><span class="sxs-lookup"><span data-stu-id="e5f52-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="e5f52-165">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-165">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="e5f52-166">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e5f52-166">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="e5f52-167">Ajout d’une nouvelle méthode qui obtient un objet qui représente les sharedProperties d’un élément rendez-vous ou message.</span><span class="sxs-lookup"><span data-stu-id="e5f52-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="e5f52-168">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-168">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="e5f52-169">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="e5f52-169">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="e5f52-170">Ajout d’une nouvelle énumération d’indicateur binaire qui spécifie les autorisations accordées aux délégués.</span><span class="sxs-lookup"><span data-stu-id="e5f52-170">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="e5f52-171">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-171">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="e5f52-172">Élément de manifeste SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e5f52-172">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="e5f52-173">Ajout d’un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e5f52-173">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="e5f52-174">Définit si le complément est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="e5f52-174">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="e5f52-175">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-175">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="e5f52-176">Emplacement amélioré</span><span class="sxs-lookup"><span data-stu-id="e5f52-176">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="e5f52-177">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e5f52-177">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="e5f52-178">Ajout d’un nouvel objet représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e5f52-178">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e5f52-179">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-179">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="e5f52-180">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="e5f52-180">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="e5f52-181">Ajout d’un nouvel objet représentant un emplacement.</span><span class="sxs-lookup"><span data-stu-id="e5f52-181">Added a new object that represents a location.</span></span> <span data-ttu-id="e5f52-182">En lecture seule.</span><span class="sxs-lookup"><span data-stu-id="e5f52-182">Read only.</span></span>

<span data-ttu-id="e5f52-183">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-183">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="e5f52-184">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="e5f52-184">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="e5f52-185">Ajout d’un nouvel objet représentant l’ID d’un emplacement.</span><span class="sxs-lookup"><span data-stu-id="e5f52-185">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="e5f52-186">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-186">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="e5f52-187">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e5f52-187">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="e5f52-188">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e5f52-188">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e5f52-189">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-189">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="e5f52-190">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="e5f52-190">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="e5f52-191">Ajout d’une nouvelle énumération qui spécifie le type d’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e5f52-191">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="e5f52-192">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-192">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e5f52-193">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="e5f52-193">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e5f52-194">Ajout de l’événement `EnhancedLocationsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="e5f52-194">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="e5f52-195">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-195">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="e5f52-196">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="e5f52-196">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="e5f52-197">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e5f52-197">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="e5f52-198">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="e5f52-198">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="e5f52-199">**Disponible dans** : Outlook pour Windows (Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="e5f52-199">**Available in**: Outlook for Windows (Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="e5f52-200">En-têtes Internet</span><span class="sxs-lookup"><span data-stu-id="e5f52-200">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="e5f52-201">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e5f52-201">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="e5f52-202">Ajout d’un nouvel objet représentant les en-têtes Internet d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="e5f52-202">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="e5f52-203">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-203">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="e5f52-204">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="e5f52-204">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="e5f52-205">Ajout d’une nouvelle propriété représentant les en-têtes Internet d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="e5f52-205">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="e5f52-206">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-206">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="e5f52-207">Thème Office</span><span class="sxs-lookup"><span data-stu-id="e5f52-207">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="e5f52-208">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="e5f52-208">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="e5f52-209">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="e5f52-209">Added ability to get Office theme.</span></span>

<span data-ttu-id="e5f52-210">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-210">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e5f52-211">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="e5f52-211">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e5f52-212">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="e5f52-212">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="e5f52-213">**Disponible dans** : Outlook pour Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e5f52-213">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="e5f52-214">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="e5f52-214">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="e5f52-215">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e5f52-215">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="e5f52-216">Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e5f52-216">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="e5f52-217">**Disponible dans** : Outlook pour Windows (Office 365), Outlook pour Mac (Office 365), Outlook sur le web (Office 365 et Outlook.com), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="e5f52-217">**Available in**: Outlook for Windows (Office 365), Outlook for Mac (Office 365), Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="e5f52-218">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e5f52-218">See also</span></span>

- [<span data-ttu-id="e5f52-219">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="e5f52-219">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="e5f52-220">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="e5f52-220">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="e5f52-221">Prise en main</span><span class="sxs-lookup"><span data-stu-id="e5f52-221">Get started</span></span>](/outlook/add-ins/quick-start)
