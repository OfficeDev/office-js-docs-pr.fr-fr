---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: b563d6cfc279a18a6a61f39c33a5ab42e1bd6984
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395707"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9bf2e-102">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="9bf2e-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9bf2e-104">Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="9bf2e-105">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9bf2e-106">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="9bf2e-107">La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span>
>
> <span data-ttu-id="9bf2e-108">Pour utiliser les API disponibles en préversion :</span><span class="sxs-lookup"><span data-stu-id="9bf2e-108">To use preview APIs:</span></span>
>
> - <span data-ttu-id="9bf2e-109">Vous devez référencer la bibliothèque **bêta** sur le CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="9bf2e-109">You must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span>
> - <span data-ttu-id="9bf2e-110">Vous devrez peut-être aussi rejoindre le [programme Office Insider](https://products.office.com/office-insider) pour accéder aux versions plus récentes d’Office.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-110">You may also need to join the [Office Insider program](https://products.office.com/office-insider) for access to more recent Office builds.</span></span>

<span data-ttu-id="9bf2e-111">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="9bf2e-111">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9bf2e-112">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="9bf2e-112">Features in preview</span></span>

<span data-ttu-id="9bf2e-113">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-113">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="9bf2e-114">Pièces jointes</span><span class="sxs-lookup"><span data-stu-id="9bf2e-114">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="9bf2e-115">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="9bf2e-115">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="9bf2e-116">Ajout d’un nouvel objet représentant le contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-116">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="9bf2e-117">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-117">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="9bf2e-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="9bf2e-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="9bf2e-119">Ajout d’une nouvelle méthode qui vous permet de joindre un fichier représenté par une chaîne encodée en base 64 à un message ou à un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-119">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="9bf2e-120">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="9bf2e-121">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="9bf2e-121">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="9bf2e-122">Ajout d’une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-122">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="9bf2e-123">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-123">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="9bf2e-124">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="9bf2e-124">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="9bf2e-125">Ajout d’une nouvelle méthode qui obtient les pièces jointes d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-125">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="9bf2e-126">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-126">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="9bf2e-127">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="9bf2e-127">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="9bf2e-128">Ajout d’une nouvelle énumération qui spécifie la mise en forme qui s’applique au contenu d’une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-128">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="9bf2e-129">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-129">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="9bf2e-130">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="9bf2e-130">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="9bf2e-131">Ajout d’une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-131">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="9bf2e-132">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-132">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="9bf2e-133">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="9bf2e-133">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9bf2e-134">Ajout de l’événement `AttachmentsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-134">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="9bf2e-135">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="9bf2e-136">Blocage lors de l’envoi</span><span class="sxs-lookup"><span data-stu-id="9bf2e-136">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="9bf2e-137">Event.completed</span><span class="sxs-lookup"><span data-stu-id="9bf2e-137">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="9bf2e-138">Ajout d’un nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-138">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="9bf2e-139">Cette valeur est utilisée pour annuler l’exécution d’un événement.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-139">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="9bf2e-140">**Disponible dans** : Outlook sur le web (classique), Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-140">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="categories"></a><span data-ttu-id="9bf2e-141">Catégories</span><span class="sxs-lookup"><span data-stu-id="9bf2e-141">Categories</span></span>

<span data-ttu-id="9bf2e-142">Dans Outlook, un utilisateur peut regrouper des messages et des rendez-vous à l’aide d’une catégorie pour leur appliquer un code de couleur.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-142">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="9bf2e-143">L’utilisateur définit les catégories dans une liste sur sa boîte aux lettres principale.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-143">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="9bf2e-144">Ils peuvent ensuite appliquer une ou plusieurs catégories à un élément.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-144">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="9bf2e-145">Cette fonctionnalité n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-145">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="9bf2e-146">Categories</span><span class="sxs-lookup"><span data-stu-id="9bf2e-146">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="9bf2e-147">Ajout d’un nouvel objet représentant des catégories d’un élément.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-147">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="9bf2e-148">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="9bf2e-149">Détailscatégorie</span><span class="sxs-lookup"><span data-stu-id="9bf2e-149">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="9bf2e-150">Ajout un nouvel objet qui représente les détails d’une catégorie (son nom et la couleur associée).</span><span class="sxs-lookup"><span data-stu-id="9bf2e-150">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="9bf2e-151">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-151">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="9bf2e-152">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="9bf2e-152">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="9bf2e-153">Ajout d’ un nouvel objet qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-153">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="9bf2e-154">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-154">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="9bf2e-155">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="9bf2e-155">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="9bf2e-156">Ajout d’ un nouveau propriétaire qui représente la liste Catégories maître sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-156">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="9bf2e-157">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-157">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="9bf2e-158">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="9bf2e-158">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="9bf2e-159">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un élément.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-159">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="9bf2e-160">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="9bf2e-161">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="9bf2e-161">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="9bf2e-162">Ajouté un nouvel enum qui spécifie les couleurs disponibles à associer à des catégories.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-162">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="9bf2e-163">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="9bf2e-164">Accès délégué</span><span class="sxs-lookup"><span data-stu-id="9bf2e-164">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="9bf2e-165">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="9bf2e-165">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="9bf2e-166">Ajout d’un nouvel objet qui représente les propriétés d’un élément rendez-vous ou message dans un dossier, un calendrier ou une boîte aux lettres partagés.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-166">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="9bf2e-167">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="9bf2e-168">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="9bf2e-168">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="9bf2e-169">Ajout d’une nouvelle méthode qui obtient l’ID d’un rendez-vous ou d’un élément de message enregistré.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-169">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="9bf2e-170">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="9bf2e-171">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="9bf2e-171">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="9bf2e-172">Ajout d’une nouvelle méthode qui obtient un objet qui représente les sharedProperties d’un élément rendez-vous ou message.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-172">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="9bf2e-173">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="9bf2e-174">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="9bf2e-174">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="9bf2e-175">Ajout d’une nouvelle énumération d’indicateur binaire qui spécifie les autorisations accordées aux délégués.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-175">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="9bf2e-176">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="9bf2e-177">Élément de manifeste SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="9bf2e-177">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="9bf2e-178">Ajout d’un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="9bf2e-178">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="9bf2e-179">Définit si le complément est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-179">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="9bf2e-180">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="9bf2e-181">Emplacement amélioré</span><span class="sxs-lookup"><span data-stu-id="9bf2e-181">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="9bf2e-182">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="9bf2e-182">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="9bf2e-183">Ajout d’un nouvel objet représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-183">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="9bf2e-184">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="9bf2e-185">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="9bf2e-185">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="9bf2e-186">Ajout d’un nouvel objet représentant un emplacement.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-186">Added a new object that represents a location.</span></span> <span data-ttu-id="9bf2e-187">En lecture seule.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-187">Read only.</span></span>

<span data-ttu-id="9bf2e-188">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-188">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="9bf2e-189">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="9bf2e-189">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="9bf2e-190">Ajout d’un nouvel objet représentant l’ID d’un emplacement.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-190">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="9bf2e-191">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-191">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="9bf2e-192">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="9bf2e-192">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="9bf2e-193">Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-193">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="9bf2e-194">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-194">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="9bf2e-195">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="9bf2e-195">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="9bf2e-196">Ajout d’une nouvelle énumération qui spécifie le type d’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-196">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="9bf2e-197">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-197">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="9bf2e-198">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="9bf2e-198">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9bf2e-199">Ajout de l’événement `EnhancedLocationsChanged` à `Item`.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-199">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="9bf2e-200">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (moderne), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-200">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9bf2e-201">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="9bf2e-201">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="9bf2e-202">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9bf2e-202">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="9bf2e-203">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="9bf2e-203">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9bf2e-204">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-204">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="9bf2e-205">En-têtes Internet</span><span class="sxs-lookup"><span data-stu-id="9bf2e-205">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="9bf2e-206">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="9bf2e-206">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="9bf2e-207">Ajout d’un nouvel objet représentant les en-têtes Internet personnalisés d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-207">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="9bf2e-208">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-208">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="9bf2e-209">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="9bf2e-209">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="9bf2e-210">Ajout d’une nouvelle propriété représentant les en-têtes Internet personnalisés d’un élément de message.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-210">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="9bf2e-211">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-211">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="9bf2e-212">Thème Office</span><span class="sxs-lookup"><span data-stu-id="9bf2e-212">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="9bf2e-213">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9bf2e-213">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9bf2e-214">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-214">Added ability to get Office theme.</span></span>

<span data-ttu-id="9bf2e-215">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-215">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="9bf2e-216">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9bf2e-216">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9bf2e-217">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-217">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9bf2e-218">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-218">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="9bf2e-219">Authentification unique</span><span class="sxs-lookup"><span data-stu-id="9bf2e-219">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="9bf2e-220">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9bf2e-220">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="9bf2e-221">Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9bf2e-221">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="9bf2e-222">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="9bf2e-222">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="9bf2e-223">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9bf2e-223">See also</span></span>

- [<span data-ttu-id="9bf2e-224">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="9bf2e-224">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="9bf2e-225">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="9bf2e-225">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9bf2e-226">Prise en main</span><span class="sxs-lookup"><span data-stu-id="9bf2e-226">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="9bf2e-227">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="9bf2e-227">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
