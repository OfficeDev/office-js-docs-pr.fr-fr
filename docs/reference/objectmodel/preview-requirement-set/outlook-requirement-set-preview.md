---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: aa3f46c505e8c87508699f6e84194272ee4d13bb
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696455"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="f4da5-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="f4da5-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="f4da5-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="f4da5-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f4da5-104">Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="f4da5-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="f4da5-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="f4da5-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="f4da5-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="f4da5-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="f4da5-107">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="f4da5-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="f4da5-108">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="f4da5-108">Features in preview</span></span>

<span data-ttu-id="f4da5-109">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="f4da5-109">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="f4da5-110">Pièces jointes</span><span class="sxs-lookup"><span data-stu-id="f4da5-110">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="f4da5-111">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="f4da5-111">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="f4da5-112">Ajout d’un nouvel objet représentant le contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f4da5-112">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="f4da5-113">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="f4da5-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="f4da5-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="f4da5-115">Ajout d’une nouvelle méthode qui vous permet de joindre un fichier représenté par une chaîne encodée en base 64 à un message ou à un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f4da5-115">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="f4da5-116">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-116">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="f4da5-117">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="f4da5-117">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="f4da5-118">Ajout d’une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.</span><span class="sxs-lookup"><span data-stu-id="f4da5-118">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="f4da5-119">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-119">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="f4da5-120">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="f4da5-120">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="f4da5-121">Ajout d’une nouvelle méthode qui obtient les pièces jointes d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f4da5-121">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="f4da5-122">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-122">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="f4da5-123">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="f4da5-123">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="f4da5-124">Ajout d’une nouvelle énumération qui spécifie la mise en forme qui s’applique au contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f4da5-124">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="f4da5-125">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="f4da5-126">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="f4da5-126">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="f4da5-127">Ajout d’une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f4da5-127">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="f4da5-128">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f4da5-129">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="f4da5-129">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f4da5-130">Ajout de l’événement `AttachmentsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="f4da5-130">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="f4da5-131">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="block-on-send"></a><span data-ttu-id="f4da5-132">Blocage lors de l’envoi</span><span class="sxs-lookup"><span data-stu-id="f4da5-132">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="f4da5-133">Event.completed</span><span class="sxs-lookup"><span data-stu-id="f4da5-133">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="f4da5-134">Ajout d’un nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="f4da5-134">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="f4da5-135">Cette valeur est utilisée pour annuler l’exécution d’un événement.</span><span class="sxs-lookup"><span data-stu-id="f4da5-135">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="f4da5-136">**Disponible dans** : Outlook sur le web (classique), Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-136">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="categories"></a><span data-ttu-id="f4da5-137">Catégories</span><span class="sxs-lookup"><span data-stu-id="f4da5-137">Categories</span></span>

<span data-ttu-id="f4da5-138">Dans Outlook, un utilisateur peut regrouper des messages et des rendez-vous à l’aide d’une catégorie pour leur appliquer un code de couleur.</span><span class="sxs-lookup"><span data-stu-id="f4da5-138">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="f4da5-139">L’utilisateur définit les catégories dans une liste sur sa boîte aux lettres principale.</span><span class="sxs-lookup"><span data-stu-id="f4da5-139">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="f4da5-140">Ils peuvent ensuite appliquer une ou plusieurs catégories à un élément.</span><span class="sxs-lookup"><span data-stu-id="f4da5-140">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="f4da5-141">Cette fonctionnalité n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f4da5-141">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="f4da5-142">Categories</span><span class="sxs-lookup"><span data-stu-id="f4da5-142">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="f4da5-143">Ajout d’un nouvel objet représentant des catégories d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f4da5-143">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="f4da5-144">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-144">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="f4da5-145">Détailscatégorie</span><span class="sxs-lookup"><span data-stu-id="f4da5-145">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="f4da5-146">Ajout un nouvel objet qui représente les détails d’une catégorie (son nom et la couleur associée).</span><span class="sxs-lookup"><span data-stu-id="f4da5-146">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="f4da5-147">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-147">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="f4da5-148">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="f4da5-148">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="f4da5-149">Ajout d’ un nouvel objet qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="f4da5-149">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="f4da5-150">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-150">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="f4da5-151">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="f4da5-151">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="f4da5-152">Ajout d’ un nouveau propriétaire qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="f4da5-152">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="f4da5-153">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-153">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="f4da5-154">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="f4da5-154">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="f4da5-155">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un élément.</span><span class="sxs-lookup"><span data-stu-id="f4da5-155">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="f4da5-156">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="f4da5-157">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="f4da5-157">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="f4da5-158">Ajouté un nouvel enum qui spécifie les couleurs disponibles à associer à des catégories.</span><span class="sxs-lookup"><span data-stu-id="f4da5-158">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="f4da5-159">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="delegate-access"></a><span data-ttu-id="f4da5-160">Accès délégué</span><span class="sxs-lookup"><span data-stu-id="f4da5-160">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="f4da5-161">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="f4da5-161">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="f4da5-162">Ajout d’un nouvel objet qui représente les propriétés d’un élément rendez-vous ou message dans un dossier, un calendrier ou une boîte aux lettres partagés.</span><span class="sxs-lookup"><span data-stu-id="f4da5-162">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="f4da5-163">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="f4da5-164">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="f4da5-164">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="f4da5-165">Ajout d’une nouvelle méthode qui obtient l’ID d’un rendez-vous ou d’un élément de message enregistré.</span><span class="sxs-lookup"><span data-stu-id="f4da5-165">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="f4da5-166">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-166">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="f4da5-167">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f4da5-167">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="f4da5-168">Ajout d’une nouvelle méthode qui obtient un objet qui représente les sharedProperties d’un élément rendez-vous ou message.</span><span class="sxs-lookup"><span data-stu-id="f4da5-168">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="f4da5-169">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-169">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="f4da5-170">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="f4da5-170">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="f4da5-171">Ajout d’une nouvelle énumération d’indicateur binaire qui spécifie les autorisations accordées aux délégués.</span><span class="sxs-lookup"><span data-stu-id="f4da5-171">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="f4da5-172">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-172">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="f4da5-173">Élément de manifeste SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="f4da5-173">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="f4da5-174">Ajout d’un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="f4da5-174">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="f4da5-175">Définit si le complément est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="f4da5-175">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="f4da5-176">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="enhanced-location"></a><span data-ttu-id="f4da5-177">Emplacement amélioré</span><span class="sxs-lookup"><span data-stu-id="f4da5-177">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="f4da5-178">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f4da5-178">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="f4da5-179">Ajout d’un nouvel objet représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f4da5-179">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="f4da5-180">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="f4da5-181">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="f4da5-181">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="f4da5-182">Ajout d’un nouvel objet représentant un emplacement.</span><span class="sxs-lookup"><span data-stu-id="f4da5-182">Added a new object that represents a location.</span></span> <span data-ttu-id="f4da5-183">En lecture seule.</span><span class="sxs-lookup"><span data-stu-id="f4da5-183">Read only.</span></span>

<span data-ttu-id="f4da5-184">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="f4da5-185">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="f4da5-185">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="f4da5-186">Ajout d’un nouvel objet représentant l’ID d’un emplacement.</span><span class="sxs-lookup"><span data-stu-id="f4da5-186">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="f4da5-187">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="f4da5-188">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f4da5-188">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="f4da5-189">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f4da5-189">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="f4da5-190">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-190">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="f4da5-191">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="f4da5-191">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="f4da5-192">Ajout d’une nouvelle énumération qui spécifie le type d’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f4da5-192">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="f4da5-193">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-193">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f4da5-194">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="f4da5-194">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f4da5-195">Ajout de l’événement `EnhancedLocationsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="f4da5-195">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="f4da5-196">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-196">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="f4da5-197">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="f4da5-197">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="f4da5-198">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f4da5-198">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="f4da5-199">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="f4da5-199">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="f4da5-200">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="f4da5-200">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

<br>

---

### <a name="internet-headers"></a><span data-ttu-id="f4da5-201">En-têtes Internet</span><span class="sxs-lookup"><span data-stu-id="f4da5-201">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="f4da5-202">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="f4da5-202">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="f4da5-203">Ajout d’un nouvel objet représentant les en-têtes Internet personnalisés d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="f4da5-203">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="f4da5-204">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-204">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="f4da5-205">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="f4da5-205">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="f4da5-206">Ajout d’une nouvelle propriété représentant les en-têtes Internet personnalisés d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="f4da5-206">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="f4da5-207">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-207">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="office-theme"></a><span data-ttu-id="f4da5-208">Thème Office</span><span class="sxs-lookup"><span data-stu-id="f4da5-208">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="f4da5-209">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="f4da5-209">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="f4da5-210">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="f4da5-210">Added ability to get Office theme.</span></span>

<span data-ttu-id="f4da5-211">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-211">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f4da5-212">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="f4da5-212">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f4da5-213">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="f4da5-213">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="f4da5-214">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f4da5-214">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="f4da5-215">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="f4da5-215">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="f4da5-216">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f4da5-216">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="f4da5-217">Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="f4da5-217">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="f4da5-218">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="f4da5-218">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="f4da5-219">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f4da5-219">See also</span></span>

- [<span data-ttu-id="f4da5-220">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="f4da5-220">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="f4da5-221">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="f4da5-221">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f4da5-222">Prise en main</span><span class="sxs-lookup"><span data-stu-id="f4da5-222">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="f4da5-223">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="f4da5-223">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
