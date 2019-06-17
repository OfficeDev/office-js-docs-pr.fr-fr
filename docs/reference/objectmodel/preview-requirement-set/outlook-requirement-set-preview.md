---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 06/14/2019
localization_priority: Priority
ms.openlocfilehash: 346750557e68508f2a5707433dea122052bc2016
ms.sourcegitcommit: e112a9b29376b1f574ee13b01c818131b2c7889d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2019
ms.locfileid: "34997371"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="77572-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="77572-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="77572-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="77572-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="77572-104">Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="77572-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="77572-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="77572-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="77572-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="77572-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="77572-107">La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.</span><span class="sxs-lookup"><span data-stu-id="77572-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="77572-108">Vous devrez également participer au [programme Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="77572-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="77572-109">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="77572-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="77572-110">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="77572-110">Features in preview</span></span>

<span data-ttu-id="77572-111">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="77572-111">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="77572-112">Pièces jointes</span><span class="sxs-lookup"><span data-stu-id="77572-112">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="77572-113">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="77572-113">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="77572-114">Ajout d’un nouvel objet représentant le contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="77572-114">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="77572-115">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-115">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="77572-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="77572-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="77572-117">Ajout d’une nouvelle méthode qui vous permet de joindre un fichier représenté par une chaîne encodée en base 64 à un message ou à un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="77572-117">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="77572-118">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-118">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="77572-119">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="77572-119">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="77572-120">Ajout d’une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.</span><span class="sxs-lookup"><span data-stu-id="77572-120">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="77572-121">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-121">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="77572-122">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="77572-122">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="77572-123">Ajout d’une nouvelle méthode qui obtient les pièces jointes d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="77572-123">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="77572-124">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-124">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="77572-125">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="77572-125">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="77572-126">Ajout d’une nouvelle énumération qui spécifie la mise en forme qui s’applique au contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="77572-126">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="77572-127">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-127">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="77572-128">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="77572-128">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="77572-129">Ajout d’une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.</span><span class="sxs-lookup"><span data-stu-id="77572-129">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="77572-130">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-130">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="77572-131">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="77572-131">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="77572-132">Ajout de l’événement `AttachmentsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="77572-132">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="77572-133">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-133">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="77572-134">Blocage lors de l’envoi</span><span class="sxs-lookup"><span data-stu-id="77572-134">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="77572-135">Event.completed</span><span class="sxs-lookup"><span data-stu-id="77572-135">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="77572-136">Ajout d’un nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="77572-136">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="77572-137">Cette valeur est utilisée pour annuler l’exécution d’un événement.</span><span class="sxs-lookup"><span data-stu-id="77572-137">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="77572-138">**Disponible dans** : Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="77572-138">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="categories"></a><span data-ttu-id="77572-139">Catégories</span><span class="sxs-lookup"><span data-stu-id="77572-139">Categories</span></span>

<span data-ttu-id="77572-140">Dans Outlook, un utilisateur peut regrouper des messages et des rendez-vous à l’aide d’une catégorie pour leur appliquer un code de couleur.</span><span class="sxs-lookup"><span data-stu-id="77572-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="77572-141">L’utilisateur définit les catégories dans une liste sur sa boîte aux lettres principale.</span><span class="sxs-lookup"><span data-stu-id="77572-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="77572-142">Ils peuvent ensuite appliquer une ou plusieurs catégories à un élément.</span><span class="sxs-lookup"><span data-stu-id="77572-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="77572-143">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="77572-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="77572-144">Catégories</span><span class="sxs-lookup"><span data-stu-id="77572-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="77572-145">Ajout d’un nouvel objet représentant des catégories d’un élément.</span><span class="sxs-lookup"><span data-stu-id="77572-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="77572-146">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-146">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="77572-147">Détailscatégorie</span><span class="sxs-lookup"><span data-stu-id="77572-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="77572-148">Ajout un nouvel objet qui représente les détails d’une catégorie (son nom et la couleur associée).</span><span class="sxs-lookup"><span data-stu-id="77572-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="77572-149">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-149">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="77572-150">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="77572-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="77572-151">Ajout d’ un nouvel objet qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="77572-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="77572-152">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-152">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="77572-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="77572-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="77572-154">Ajout d’ un nouveau propriétaire qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="77572-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="77572-155">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-155">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="77572-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="77572-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="77572-157">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un élément.</span><span class="sxs-lookup"><span data-stu-id="77572-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="77572-158">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-158">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="77572-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="77572-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="77572-160">Ajouté un nouvel enum qui spécifie les couleurs disponibles à associer à des catégories.</span><span class="sxs-lookup"><span data-stu-id="77572-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="77572-161">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-161">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="77572-162">Accès délégué</span><span class="sxs-lookup"><span data-stu-id="77572-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="77572-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="77572-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="77572-164">Ajout d’un nouvel objet qui représente les propriétés d’un élément rendez-vous ou message dans un dossier, un calendrier ou une boîte aux lettres partagés.</span><span class="sxs-lookup"><span data-stu-id="77572-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="77572-165">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-165">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="77572-166">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="77572-166">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="77572-167">Ajout d’une nouvelle méthode qui obtient l’ID d’un rendez-vous ou d’un élément de message enregistré.</span><span class="sxs-lookup"><span data-stu-id="77572-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="77572-168">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-168">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="77572-169">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="77572-169">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="77572-170">Ajout d’une nouvelle méthode qui obtient un objet qui représente les sharedProperties d’un élément rendez-vous ou message.</span><span class="sxs-lookup"><span data-stu-id="77572-170">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="77572-171">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-171">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="77572-172">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="77572-172">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="77572-173">Ajout d’une nouvelle énumération d’indicateur binaire qui spécifie les autorisations accordées aux délégués.</span><span class="sxs-lookup"><span data-stu-id="77572-173">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="77572-174">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-174">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="77572-175">Élément de manifeste SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="77572-175">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="77572-176">Ajout d’un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="77572-176">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="77572-177">Définit si le complément est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="77572-177">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="77572-178">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-178">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="77572-179">Emplacement amélioré</span><span class="sxs-lookup"><span data-stu-id="77572-179">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="77572-180">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="77572-180">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="77572-181">Ajout d’un nouvel objet représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="77572-181">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="77572-182">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-182">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="77572-183">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="77572-183">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="77572-184">Ajout d’un nouvel objet représentant un emplacement.</span><span class="sxs-lookup"><span data-stu-id="77572-184">Added a new object that represents a location.</span></span> <span data-ttu-id="77572-185">En lecture seule.</span><span class="sxs-lookup"><span data-stu-id="77572-185">Read only.</span></span>

<span data-ttu-id="77572-186">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-186">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="77572-187">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="77572-187">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="77572-188">Ajout d’un nouvel objet représentant l’ID d’un emplacement.</span><span class="sxs-lookup"><span data-stu-id="77572-188">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="77572-189">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-189">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="77572-190">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="77572-190">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="77572-191">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="77572-191">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="77572-192">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-192">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="77572-193">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="77572-193">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="77572-194">Ajout d’une nouvelle énumération qui spécifie le type d’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="77572-194">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="77572-195">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-195">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="77572-196">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="77572-196">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="77572-197">Ajout de l’événement `EnhancedLocationsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="77572-197">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="77572-198">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-198">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="77572-199">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="77572-199">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="77572-200">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="77572-200">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="77572-201">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="77572-201">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="77572-202">**Disponible dans** : Outlook pour Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="77572-202">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="77572-203">En-têtes Internet</span><span class="sxs-lookup"><span data-stu-id="77572-203">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="77572-204">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="77572-204">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="77572-205">Ajout d’un nouvel objet représentant les en-têtes Internet d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="77572-205">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="77572-206">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-206">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="77572-207">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="77572-207">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="77572-208">Ajout d’une nouvelle propriété représentant les en-têtes Internet d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="77572-208">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="77572-209">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-209">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="77572-210">Thème Office</span><span class="sxs-lookup"><span data-stu-id="77572-210">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="77572-211">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="77572-211">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="77572-212">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="77572-212">Added ability to get Office theme.</span></span>

<span data-ttu-id="77572-213">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="77572-214">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="77572-214">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="77572-215">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="77572-215">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="77572-216">**Disponible dans** : Outlook pour Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="77572-216">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="77572-217">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="77572-217">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="77572-218">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="77572-218">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="77572-219">Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="77572-219">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="77572-220">**Disponible dans** : Outlook pour Windows (connecté à Office 365), Outlook pour Mac (connecté à Office 365), Outlook sur le web (nouveau), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="77572-220">**Available in**: Outlook on Windows (connected to Office 365), Outlook for Mac (connected to Office 365), Outlook on the web (Outlook.com and connected to Office 365), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="77572-221">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="77572-221">See also</span></span>

- [<span data-ttu-id="77572-222">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="77572-222">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="77572-223">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="77572-223">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="77572-224">Prise en main</span><span class="sxs-lookup"><span data-stu-id="77572-224">Get started</span></span>](/outlook/add-ins/quick-start)
