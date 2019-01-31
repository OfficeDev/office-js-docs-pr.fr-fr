---
title: Office.context.mailbox.item - ensemble de conditions requises 1.5
description: ''
ms.date: 12/18/2018
localization_priority: Priority
ms.openlocfilehash: 48bc1291e7aa6d8e335c07d16ddd74e6e9455f0d
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389570"
---
# <a name="item"></a><span data-ttu-id="088e5-102">élément</span><span class="sxs-lookup"><span data-stu-id="088e5-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="088e5-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="088e5-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="088e5-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="088e5-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-106">Requirements</span></span>

|<span data-ttu-id="088e5-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-107">Requirement</span></span>| <span data-ttu-id="088e5-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-110">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-110">1.0</span></span>|
|[<span data-ttu-id="088e5-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="088e5-112">Restricted</span></span>|
|[<span data-ttu-id="088e5-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="088e5-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="088e5-115">Members and methods</span></span>

| <span data-ttu-id="088e5-116">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-116">Member</span></span> | <span data-ttu-id="088e5-117">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="088e5-118">attachments</span><span class="sxs-lookup"><span data-stu-id="088e5-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="088e5-119">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-119">Member</span></span> |
| [<span data-ttu-id="088e5-120">bcc</span><span class="sxs-lookup"><span data-stu-id="088e5-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="088e5-121">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-121">Member</span></span> |
| [<span data-ttu-id="088e5-122">body</span><span class="sxs-lookup"><span data-stu-id="088e5-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="088e5-123">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-123">Member</span></span> |
| [<span data-ttu-id="088e5-124">cc</span><span class="sxs-lookup"><span data-stu-id="088e5-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="088e5-125">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-125">Member</span></span> |
| [<span data-ttu-id="088e5-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="088e5-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="088e5-127">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-127">Member</span></span> |
| [<span data-ttu-id="088e5-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="088e5-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="088e5-129">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-129">Member</span></span> |
| [<span data-ttu-id="088e5-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="088e5-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="088e5-131">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-131">Member</span></span> |
| [<span data-ttu-id="088e5-132">end</span><span class="sxs-lookup"><span data-stu-id="088e5-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="088e5-133">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-133">Member</span></span> |
| [<span data-ttu-id="088e5-134">from</span><span class="sxs-lookup"><span data-stu-id="088e5-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="088e5-135">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-135">Member</span></span> |
| [<span data-ttu-id="088e5-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="088e5-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="088e5-137">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-137">Member</span></span> |
| [<span data-ttu-id="088e5-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="088e5-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="088e5-139">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-139">Member</span></span> |
| [<span data-ttu-id="088e5-140">itemId</span><span class="sxs-lookup"><span data-stu-id="088e5-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="088e5-141">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-141">Member</span></span> |
| [<span data-ttu-id="088e5-142">itemType</span><span class="sxs-lookup"><span data-stu-id="088e5-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="088e5-143">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-143">Member</span></span> |
| [<span data-ttu-id="088e5-144">location</span><span class="sxs-lookup"><span data-stu-id="088e5-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="088e5-145">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-145">Member</span></span> |
| [<span data-ttu-id="088e5-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="088e5-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="088e5-147">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-147">Member</span></span> |
| [<span data-ttu-id="088e5-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="088e5-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="088e5-149">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-149">Member</span></span> |
| [<span data-ttu-id="088e5-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="088e5-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="088e5-151">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-151">Member</span></span> |
| [<span data-ttu-id="088e5-152">organizer</span><span class="sxs-lookup"><span data-stu-id="088e5-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="088e5-153">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-153">Member</span></span> |
| [<span data-ttu-id="088e5-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="088e5-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="088e5-155">Member</span><span class="sxs-lookup"><span data-stu-id="088e5-155">Member</span></span> |
| [<span data-ttu-id="088e5-156">sender</span><span class="sxs-lookup"><span data-stu-id="088e5-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="088e5-157">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-157">Member</span></span> |
| [<span data-ttu-id="088e5-158">start</span><span class="sxs-lookup"><span data-stu-id="088e5-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="088e5-159">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-159">Member</span></span> |
| [<span data-ttu-id="088e5-160">subject</span><span class="sxs-lookup"><span data-stu-id="088e5-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="088e5-161">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-161">Member</span></span> |
| [<span data-ttu-id="088e5-162">to</span><span class="sxs-lookup"><span data-stu-id="088e5-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="088e5-163">Membre</span><span class="sxs-lookup"><span data-stu-id="088e5-163">Member</span></span> |
| [<span data-ttu-id="088e5-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="088e5-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-165">Method</span></span> |
| [<span data-ttu-id="088e5-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="088e5-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-167">Method</span></span> |
| [<span data-ttu-id="088e5-168">close</span><span class="sxs-lookup"><span data-stu-id="088e5-168">close</span></span>](#close) | <span data-ttu-id="088e5-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-169">Method</span></span> |
| [<span data-ttu-id="088e5-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="088e5-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="088e5-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-171">Method</span></span> |
| [<span data-ttu-id="088e5-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="088e5-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="088e5-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-173">Method</span></span> |
| [<span data-ttu-id="088e5-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="088e5-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="088e5-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-175">Method</span></span> |
| [<span data-ttu-id="088e5-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="088e5-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="088e5-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-177">Method</span></span> |
| [<span data-ttu-id="088e5-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="088e5-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="088e5-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-179">Method</span></span> |
| [<span data-ttu-id="088e5-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="088e5-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="088e5-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-181">Method</span></span> |
| [<span data-ttu-id="088e5-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="088e5-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="088e5-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-183">Method</span></span> |
| [<span data-ttu-id="088e5-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="088e5-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-185">Method</span></span> |
| [<span data-ttu-id="088e5-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="088e5-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-187">Method</span></span> |
| [<span data-ttu-id="088e5-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="088e5-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-189">Method</span></span> |
| [<span data-ttu-id="088e5-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="088e5-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-191">Method</span></span> |
| [<span data-ttu-id="088e5-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="088e5-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="088e5-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="088e5-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="088e5-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-194">Example</span></span>

<span data-ttu-id="088e5-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="088e5-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="088e5-196">Membres</span><span class="sxs-lookup"><span data-stu-id="088e5-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="088e5-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="088e5-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="088e5-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="088e5-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="088e5-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="088e5-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-202">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-202">Type:</span></span>

*   <span data-ttu-id="088e5-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="088e5-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-204">Requirements</span></span>

|<span data-ttu-id="088e5-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-205">Requirement</span></span>| <span data-ttu-id="088e5-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-208">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-208">1.0</span></span>|
|[<span data-ttu-id="088e5-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-210">ReadItem</span></span>|
|[<span data-ttu-id="088e5-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-213">Example</span></span>

<span data-ttu-id="088e5-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="088e5-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="088e5-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="088e5-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="088e5-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-218">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-218">Type:</span></span>

*   [<span data-ttu-id="088e5-219">Destinataires</span><span class="sxs-lookup"><span data-stu-id="088e5-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="088e5-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-220">Requirements</span></span>

|<span data-ttu-id="088e5-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-221">Requirement</span></span>| <span data-ttu-id="088e5-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-224">1.1</span><span class="sxs-lookup"><span data-stu-id="088e5-224">1.1</span></span>|
|[<span data-ttu-id="088e5-225">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-226">ReadItem</span></span>|
|[<span data-ttu-id="088e5-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-228">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-229">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="088e5-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="088e5-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="088e5-231">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-232">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-232">Type:</span></span>

*   [<span data-ttu-id="088e5-233">Corps</span><span class="sxs-lookup"><span data-stu-id="088e5-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="088e5-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-234">Requirements</span></span>

|<span data-ttu-id="088e5-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-235">Requirement</span></span>| <span data-ttu-id="088e5-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-238">1.1</span><span class="sxs-lookup"><span data-stu-id="088e5-238">1.1</span></span>|
|[<span data-ttu-id="088e5-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-240">ReadItem</span></span>|
|[<span data-ttu-id="088e5-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-242">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-242">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="088e5-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="088e5-244">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="088e5-244">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="088e5-245">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-245">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-246">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-246">Read mode</span></span>

<span data-ttu-id="088e5-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="088e5-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-249">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-249">Compose mode</span></span>

<span data-ttu-id="088e5-250">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="088e5-250">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-251">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-251">Type:</span></span>

*   <span data-ttu-id="088e5-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-253">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-253">Requirements</span></span>

|<span data-ttu-id="088e5-254">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-254">Requirement</span></span>| <span data-ttu-id="088e5-255">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-256">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-257">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-257">1.0</span></span>|
|[<span data-ttu-id="088e5-258">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-258">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-259">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-259">ReadItem</span></span>|
|[<span data-ttu-id="088e5-260">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-260">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-261">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-261">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-262">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-262">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="088e5-263">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="088e5-263">(nullable) conversationId :String</span></span>

<span data-ttu-id="088e5-264">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="088e5-264">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="088e5-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="088e5-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="088e5-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="088e5-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-269">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-269">Type:</span></span>

*   <span data-ttu-id="088e5-270">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-270">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-271">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-271">Requirements</span></span>

|<span data-ttu-id="088e5-272">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-272">Requirement</span></span>| <span data-ttu-id="088e5-273">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-273">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-274">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-275">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-275">1.0</span></span>|
|[<span data-ttu-id="088e5-276">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-277">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-277">ReadItem</span></span>|
|[<span data-ttu-id="088e5-278">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-279">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-279">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="088e5-280">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="088e5-280">dateTimeCreated :Date</span></span>

<span data-ttu-id="088e5-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-283">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-283">Type:</span></span>

*   <span data-ttu-id="088e5-284">Date</span><span class="sxs-lookup"><span data-stu-id="088e5-284">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-285">Requirements</span></span>

|<span data-ttu-id="088e5-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-286">Requirement</span></span>| <span data-ttu-id="088e5-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-289">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-289">1.0</span></span>|
|[<span data-ttu-id="088e5-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-291">ReadItem</span></span>|
|[<span data-ttu-id="088e5-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-293">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-293">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-294">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-294">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="088e5-295">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="088e5-295">dateTimeModified :Date</span></span>

<span data-ttu-id="088e5-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-298">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-298">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-299">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-299">Type:</span></span>

*   <span data-ttu-id="088e5-300">Date</span><span class="sxs-lookup"><span data-stu-id="088e5-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-301">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-301">Requirements</span></span>

|<span data-ttu-id="088e5-302">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-302">Requirement</span></span>| <span data-ttu-id="088e5-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-304">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-305">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-305">1.0</span></span>|
|[<span data-ttu-id="088e5-306">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-307">ReadItem</span></span>|
|[<span data-ttu-id="088e5-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-310">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="088e5-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="088e5-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="088e5-312">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-312">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="088e5-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="088e5-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-315">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-315">Read mode</span></span>

<span data-ttu-id="088e5-316">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="088e5-316">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-317">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-317">Compose mode</span></span>

<span data-ttu-id="088e5-318">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="088e5-318">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="088e5-319">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="088e5-319">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-320">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-320">Type:</span></span>

*   <span data-ttu-id="088e5-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="088e5-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-322">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-322">Requirements</span></span>

|<span data-ttu-id="088e5-323">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-323">Requirement</span></span>| <span data-ttu-id="088e5-324">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-325">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-326">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-326">1.0</span></span>|
|[<span data-ttu-id="088e5-327">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-328">ReadItem</span></span>|
|[<span data-ttu-id="088e5-329">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-330">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-330">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-331">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-331">Example</span></span>

<span data-ttu-id="088e5-332">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="088e5-332">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="088e5-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="088e5-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="088e5-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="088e5-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="088e5-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-338">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="088e5-338">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-339">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-339">Type:</span></span>

*   [<span data-ttu-id="088e5-340">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="088e5-340">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="088e5-341">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-341">Requirements</span></span>

|<span data-ttu-id="088e5-342">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-342">Requirement</span></span>| <span data-ttu-id="088e5-343">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-344">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-345">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-345">1.0</span></span>|
|[<span data-ttu-id="088e5-346">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-347">ReadItem</span></span>|
|[<span data-ttu-id="088e5-348">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-349">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-349">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="088e5-350">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="088e5-350">internetMessageId :String</span></span>

<span data-ttu-id="088e5-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-353">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-353">Type:</span></span>

*   <span data-ttu-id="088e5-354">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-354">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-355">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-355">Requirements</span></span>

|<span data-ttu-id="088e5-356">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-356">Requirement</span></span>| <span data-ttu-id="088e5-357">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-358">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-359">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-359">1.0</span></span>|
|[<span data-ttu-id="088e5-360">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-361">ReadItem</span></span>|
|[<span data-ttu-id="088e5-362">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-363">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-363">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-364">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-364">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="088e5-365">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="088e5-365">itemClass :String</span></span>

<span data-ttu-id="088e5-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="088e5-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="088e5-370">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-370">Type</span></span> | <span data-ttu-id="088e5-371">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-371">Description</span></span> | <span data-ttu-id="088e5-372">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="088e5-372">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="088e5-373">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="088e5-373">Appointment items</span></span> | <span data-ttu-id="088e5-374">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="088e5-374">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="088e5-375">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="088e5-375">Message items</span></span> | <span data-ttu-id="088e5-376">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="088e5-376">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="088e5-377">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="088e5-377">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-378">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-378">Type:</span></span>

*   <span data-ttu-id="088e5-379">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-379">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-380">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-380">Requirements</span></span>

|<span data-ttu-id="088e5-381">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-381">Requirement</span></span>| <span data-ttu-id="088e5-382">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-382">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-383">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-384">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-384">1.0</span></span>|
|[<span data-ttu-id="088e5-385">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-386">ReadItem</span></span>|
|[<span data-ttu-id="088e5-387">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-388">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-388">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-389">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-389">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="088e5-390">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="088e5-390">(nullable) itemId :String</span></span>

<span data-ttu-id="088e5-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-393">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="088e5-393">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="088e5-394">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="088e5-394">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="088e5-395">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="088e5-395">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="088e5-396">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="088e5-396">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="088e5-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-399">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-399">Type:</span></span>

*   <span data-ttu-id="088e5-400">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-400">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-401">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-401">Requirements</span></span>

|<span data-ttu-id="088e5-402">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-402">Requirement</span></span>| <span data-ttu-id="088e5-403">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-403">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-404">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-404">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-405">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-405">1.0</span></span>|
|[<span data-ttu-id="088e5-406">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-406">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-407">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-407">ReadItem</span></span>|
|[<span data-ttu-id="088e5-408">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-408">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-409">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-409">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-410">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-410">Example</span></span>

<span data-ttu-id="088e5-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="088e5-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="088e5-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="088e5-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="088e5-414">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="088e5-414">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="088e5-415">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-415">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-416">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-416">Type:</span></span>

*   [<span data-ttu-id="088e5-417">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="088e5-417">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="088e5-418">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-418">Requirements</span></span>

|<span data-ttu-id="088e5-419">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-419">Requirement</span></span>| <span data-ttu-id="088e5-420">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-421">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-422">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-422">1.0</span></span>|
|[<span data-ttu-id="088e5-423">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-424">ReadItem</span></span>|
|[<span data-ttu-id="088e5-425">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-426">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-426">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-427">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-427">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="088e5-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="088e5-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="088e5-429">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-429">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-430">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-430">Read mode</span></span>

<span data-ttu-id="088e5-431">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-431">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-432">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-432">Compose mode</span></span>

<span data-ttu-id="088e5-433">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-433">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-434">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-434">Type:</span></span>

*   <span data-ttu-id="088e5-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="088e5-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-436">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-436">Requirements</span></span>

|<span data-ttu-id="088e5-437">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-437">Requirement</span></span>| <span data-ttu-id="088e5-438">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-439">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-440">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-440">1.0</span></span>|
|[<span data-ttu-id="088e5-441">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-441">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-442">ReadItem</span></span>|
|[<span data-ttu-id="088e5-443">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-443">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-444">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-444">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-445">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-445">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="088e5-446">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="088e5-446">normalizedSubject :String</span></span>

<span data-ttu-id="088e5-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="088e5-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="088e5-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-451">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-451">Type:</span></span>

*   <span data-ttu-id="088e5-452">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-453">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-453">Requirements</span></span>

|<span data-ttu-id="088e5-454">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-454">Requirement</span></span>| <span data-ttu-id="088e5-455">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-456">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-457">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-457">1.0</span></span>|
|[<span data-ttu-id="088e5-458">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-459">ReadItem</span></span>|
|[<span data-ttu-id="088e5-460">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-461">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-462">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="088e5-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="088e5-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="088e5-464">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-464">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-465">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-465">Type:</span></span>

*   [<span data-ttu-id="088e5-466">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="088e5-466">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="088e5-467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-467">Requirements</span></span>

|<span data-ttu-id="088e5-468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-468">Requirement</span></span>| <span data-ttu-id="088e5-469">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-471">1.3</span><span class="sxs-lookup"><span data-stu-id="088e5-471">1.3</span></span>|
|[<span data-ttu-id="088e5-472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-473">ReadItem</span></span>|
|[<span data-ttu-id="088e5-474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-475">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-475">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="088e5-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="088e5-477">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="088e5-477">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="088e5-478">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-478">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-479">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-479">Read mode</span></span>

<span data-ttu-id="088e5-480">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="088e5-480">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-481">Compose mode</span></span>

<span data-ttu-id="088e5-482">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="088e5-482">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-483">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-483">Type:</span></span>

*   <span data-ttu-id="088e5-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-485">Requirements</span></span>

|<span data-ttu-id="088e5-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-486">Requirement</span></span>| <span data-ttu-id="088e5-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-489">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-489">1.0</span></span>|
|[<span data-ttu-id="088e5-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-491">ReadItem</span></span>|
|[<span data-ttu-id="088e5-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-493">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-493">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-494">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-494">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="088e5-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="088e5-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="088e5-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-498">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-498">Type:</span></span>

*   [<span data-ttu-id="088e5-499">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="088e5-499">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="088e5-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-500">Requirements</span></span>

|<span data-ttu-id="088e5-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-501">Requirement</span></span>| <span data-ttu-id="088e5-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-504">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-504">1.0</span></span>|
|[<span data-ttu-id="088e5-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-506">ReadItem</span></span>|
|[<span data-ttu-id="088e5-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-508">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-509">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-509">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="088e5-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="088e5-511">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="088e5-511">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="088e5-512">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-512">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-513">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-513">Read mode</span></span>

<span data-ttu-id="088e5-514">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="088e5-514">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-515">Compose mode</span></span>

<span data-ttu-id="088e5-516">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="088e5-516">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-517">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-517">Type:</span></span>

*   <span data-ttu-id="088e5-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-519">Requirements</span></span>

|<span data-ttu-id="088e5-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-520">Requirement</span></span>| <span data-ttu-id="088e5-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-523">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-523">1.0</span></span>|
|[<span data-ttu-id="088e5-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-525">ReadItem</span></span>|
|[<span data-ttu-id="088e5-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-527">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-528">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-528">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="088e5-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="088e5-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="088e5-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="088e5-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="088e5-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="088e5-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-534">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="088e5-534">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-535">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-535">Type:</span></span>

*   [<span data-ttu-id="088e5-536">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="088e5-536">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="088e5-537">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-537">Requirements</span></span>

|<span data-ttu-id="088e5-538">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-538">Requirement</span></span>| <span data-ttu-id="088e5-539">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-540">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-541">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-541">1.0</span></span>|
|[<span data-ttu-id="088e5-542">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-543">ReadItem</span></span>|
|[<span data-ttu-id="088e5-544">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-545">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-545">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-546">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-546">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="088e5-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="088e5-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="088e5-548">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-548">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="088e5-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="088e5-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-551">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-551">Read mode</span></span>

<span data-ttu-id="088e5-552">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="088e5-552">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-553">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-553">Compose mode</span></span>

<span data-ttu-id="088e5-554">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="088e5-554">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="088e5-555">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="088e5-555">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-556">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-556">Type:</span></span>

*   <span data-ttu-id="088e5-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="088e5-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-558">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-558">Requirements</span></span>

|<span data-ttu-id="088e5-559">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-559">Requirement</span></span>| <span data-ttu-id="088e5-560">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-561">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-562">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-562">1.0</span></span>|
|[<span data-ttu-id="088e5-563">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-564">ReadItem</span></span>|
|[<span data-ttu-id="088e5-565">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-566">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-567">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-567">Example</span></span>

<span data-ttu-id="088e5-568">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="088e5-568">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="088e5-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="088e5-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="088e5-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="088e5-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="088e5-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-572">Read mode</span></span>

<span data-ttu-id="088e5-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="088e5-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="088e5-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-575">Compose mode</span></span>

<span data-ttu-id="088e5-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="088e5-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="088e5-577">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-577">Type:</span></span>

*   <span data-ttu-id="088e5-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="088e5-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-579">Requirements</span></span>

|<span data-ttu-id="088e5-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-580">Requirement</span></span>| <span data-ttu-id="088e5-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-583">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-583">1.0</span></span>|
|[<span data-ttu-id="088e5-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-585">ReadItem</span></span>|
|[<span data-ttu-id="088e5-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-587">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="088e5-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="088e5-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="088e5-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="088e5-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="088e5-591">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-591">Read mode</span></span>

<span data-ttu-id="088e5-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="088e5-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="088e5-594">Mode composition</span><span class="sxs-lookup"><span data-stu-id="088e5-594">Compose mode</span></span>

<span data-ttu-id="088e5-595">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="088e5-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="088e5-596">Type :</span><span class="sxs-lookup"><span data-stu-id="088e5-596">Type:</span></span>

*   <span data-ttu-id="088e5-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="088e5-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-598">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-598">Requirements</span></span>

|<span data-ttu-id="088e5-599">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-599">Requirement</span></span>| <span data-ttu-id="088e5-600">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-601">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-602">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-602">1.0</span></span>|
|[<span data-ttu-id="088e5-603">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-604">ReadItem</span></span>|
|[<span data-ttu-id="088e5-605">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-606">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-606">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-607">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-607">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="088e5-608">Méthodes</span><span class="sxs-lookup"><span data-stu-id="088e5-608">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="088e5-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="088e5-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="088e5-610">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="088e5-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="088e5-611">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="088e5-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="088e5-612">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="088e5-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-613">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-613">Parameters:</span></span>

|<span data-ttu-id="088e5-614">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-614">Name</span></span>| <span data-ttu-id="088e5-615">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-615">Type</span></span>| <span data-ttu-id="088e5-616">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-616">Attributes</span></span>| <span data-ttu-id="088e5-617">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="088e5-618">String</span><span class="sxs-lookup"><span data-stu-id="088e5-618">String</span></span>||<span data-ttu-id="088e5-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="088e5-621">String</span><span class="sxs-lookup"><span data-stu-id="088e5-621">String</span></span>||<span data-ttu-id="088e5-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="088e5-624">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-624">Object</span></span>| <span data-ttu-id="088e5-625">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-625">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-626">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="088e5-626">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="088e5-627">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-627">Object</span></span> | <span data-ttu-id="088e5-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-628">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-629">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="088e5-630">Boolean</span><span class="sxs-lookup"><span data-stu-id="088e5-630">Boolean</span></span> | <span data-ttu-id="088e5-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-631">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-632">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="088e5-632">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="088e5-633">fonction</span><span class="sxs-lookup"><span data-stu-id="088e5-633">function</span></span>| <span data-ttu-id="088e5-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-634">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-635">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="088e5-636">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="088e5-636">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="088e5-637">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="088e5-637">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="088e5-638">Erreurs</span><span class="sxs-lookup"><span data-stu-id="088e5-638">Errors</span></span>

| <span data-ttu-id="088e5-639">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="088e5-639">Error code</span></span> | <span data-ttu-id="088e5-640">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-640">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="088e5-641">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="088e5-641">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="088e5-642">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="088e5-642">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="088e5-643">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="088e5-643">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="088e5-644">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-644">Requirements</span></span>

|<span data-ttu-id="088e5-645">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-645">Requirement</span></span>| <span data-ttu-id="088e5-646">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-646">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-647">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-648">1.1</span><span class="sxs-lookup"><span data-stu-id="088e5-648">1.1</span></span>|
|[<span data-ttu-id="088e5-649">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-649">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="088e5-650">ReadWriteItem</span></span>|
|[<span data-ttu-id="088e5-651">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-651">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-652">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-652">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="088e5-653">Exemples</span><span class="sxs-lookup"><span data-stu-id="088e5-653">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="088e5-654">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="088e5-654">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="088e5-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="088e5-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="088e5-656">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="088e5-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="088e5-660">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="088e5-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="088e5-661">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="088e5-661">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-662">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-662">Parameters:</span></span>

|<span data-ttu-id="088e5-663">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-663">Name</span></span>| <span data-ttu-id="088e5-664">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-664">Type</span></span>| <span data-ttu-id="088e5-665">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-665">Attributes</span></span>| <span data-ttu-id="088e5-666">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="088e5-667">String</span><span class="sxs-lookup"><span data-stu-id="088e5-667">String</span></span>||<span data-ttu-id="088e5-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="088e5-670">String</span><span class="sxs-lookup"><span data-stu-id="088e5-670">String</span></span>||<span data-ttu-id="088e5-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="088e5-673">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-673">Object</span></span>| <span data-ttu-id="088e5-674">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-674">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-675">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="088e5-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="088e5-676">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-676">Object</span></span>| <span data-ttu-id="088e5-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-677">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-678">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="088e5-679">fonction</span><span class="sxs-lookup"><span data-stu-id="088e5-679">function</span></span>| <span data-ttu-id="088e5-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-680">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-681">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="088e5-682">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="088e5-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="088e5-683">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="088e5-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="088e5-684">Erreurs</span><span class="sxs-lookup"><span data-stu-id="088e5-684">Errors</span></span>

| <span data-ttu-id="088e5-685">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="088e5-685">Error code</span></span> | <span data-ttu-id="088e5-686">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="088e5-687">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="088e5-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="088e5-688">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-688">Requirements</span></span>

|<span data-ttu-id="088e5-689">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-689">Requirement</span></span>| <span data-ttu-id="088e5-690">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-691">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-692">1.1</span><span class="sxs-lookup"><span data-stu-id="088e5-692">1.1</span></span>|
|[<span data-ttu-id="088e5-693">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="088e5-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="088e5-695">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-696">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-697">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-697">Example</span></span>

<span data-ttu-id="088e5-698">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="088e5-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="088e5-699">close()</span><span class="sxs-lookup"><span data-stu-id="088e5-699">close()</span></span>

<span data-ttu-id="088e5-700">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="088e5-700">Closes the current item that is being composed.</span></span>

<span data-ttu-id="088e5-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="088e5-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-703">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-703">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="088e5-704">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="088e5-704">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-705">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-705">Requirements</span></span>

|<span data-ttu-id="088e5-706">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-706">Requirement</span></span>| <span data-ttu-id="088e5-707">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-707">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-708">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-708">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-709">1.3</span><span class="sxs-lookup"><span data-stu-id="088e5-709">1.3</span></span>|
|[<span data-ttu-id="088e5-710">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-710">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-711">Restreinte</span><span class="sxs-lookup"><span data-stu-id="088e5-711">Restricted</span></span>|
|[<span data-ttu-id="088e5-712">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-712">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-713">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-713">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="088e5-714">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="088e5-714">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="088e5-715">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="088e5-715">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-716">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="088e5-717">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="088e5-717">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="088e5-718">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="088e5-718">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="088e5-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="088e5-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-722">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-722">Parameters:</span></span>

| <span data-ttu-id="088e5-723">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-723">Name</span></span> | <span data-ttu-id="088e5-724">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-724">Type</span></span> | <span data-ttu-id="088e5-725">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-725">Attributes</span></span> | <span data-ttu-id="088e5-726">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-726">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="088e5-727">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="088e5-727">String &#124; Object</span></span>| |<span data-ttu-id="088e5-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="088e5-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="088e5-730">**OU**</span><span class="sxs-lookup"><span data-stu-id="088e5-730">**OR**</span></span><br/><span data-ttu-id="088e5-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="088e5-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="088e5-733">String</span><span class="sxs-lookup"><span data-stu-id="088e5-733">String</span></span> | <span data-ttu-id="088e5-734">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-734">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="088e5-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="088e5-737">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-737">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="088e5-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-738">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-739">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-739">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="088e5-740">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-740">String</span></span> | | <span data-ttu-id="088e5-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="088e5-743">String</span><span class="sxs-lookup"><span data-stu-id="088e5-743">String</span></span> | | <span data-ttu-id="088e5-744">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-744">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="088e5-745">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-745">String</span></span> | | <span data-ttu-id="088e5-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="088e5-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="088e5-748">Booléen</span><span class="sxs-lookup"><span data-stu-id="088e5-748">Boolean</span></span> | | <span data-ttu-id="088e5-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="088e5-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="088e5-751">String</span><span class="sxs-lookup"><span data-stu-id="088e5-751">String</span></span> | | <span data-ttu-id="088e5-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="088e5-755">function</span><span class="sxs-lookup"><span data-stu-id="088e5-755">function</span></span> | <span data-ttu-id="088e5-756">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-756">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-757">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-757">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="088e5-758">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-758">Requirements</span></span>

|<span data-ttu-id="088e5-759">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-759">Requirement</span></span>| <span data-ttu-id="088e5-760">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-761">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-762">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-762">1.0</span></span>|
|[<span data-ttu-id="088e5-763">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-764">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-764">ReadItem</span></span>|
|[<span data-ttu-id="088e5-765">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-766">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-766">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="088e5-767">Exemples</span><span class="sxs-lookup"><span data-stu-id="088e5-767">Examples</span></span>

<span data-ttu-id="088e5-768">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="088e5-768">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="088e5-769">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="088e5-769">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="088e5-770">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="088e5-770">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="088e5-771">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="088e5-771">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="088e5-772">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-772">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="088e5-773">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-773">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="088e5-774">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="088e5-774">displayReplyForm(formData)</span></span>

<span data-ttu-id="088e5-775">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="088e5-775">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-776">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-776">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="088e5-777">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="088e5-777">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="088e5-778">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="088e5-778">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="088e5-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="088e5-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-782">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-782">Parameters:</span></span>

| <span data-ttu-id="088e5-783">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-783">Name</span></span> | <span data-ttu-id="088e5-784">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-784">Type</span></span> | <span data-ttu-id="088e5-785">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-785">Attributes</span></span> | <span data-ttu-id="088e5-786">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-786">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="088e5-787">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="088e5-787">String &#124; Object</span></span>| | <span data-ttu-id="088e5-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="088e5-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="088e5-790">**OU**</span><span class="sxs-lookup"><span data-stu-id="088e5-790">**OR**</span></span><br/><span data-ttu-id="088e5-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="088e5-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="088e5-793">String</span><span class="sxs-lookup"><span data-stu-id="088e5-793">String</span></span> | <span data-ttu-id="088e5-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-794">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="088e5-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="088e5-797">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-797">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="088e5-798">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-798">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-799">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-799">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="088e5-800">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-800">String</span></span> | | <span data-ttu-id="088e5-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="088e5-803">String</span><span class="sxs-lookup"><span data-stu-id="088e5-803">String</span></span> | | <span data-ttu-id="088e5-804">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-804">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="088e5-805">Chaîne</span><span class="sxs-lookup"><span data-stu-id="088e5-805">String</span></span> | | <span data-ttu-id="088e5-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="088e5-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="088e5-808">Booléen</span><span class="sxs-lookup"><span data-stu-id="088e5-808">Boolean</span></span> | | <span data-ttu-id="088e5-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="088e5-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="088e5-811">String</span><span class="sxs-lookup"><span data-stu-id="088e5-811">String</span></span> | | <span data-ttu-id="088e5-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="088e5-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="088e5-815">function</span><span class="sxs-lookup"><span data-stu-id="088e5-815">function</span></span> | <span data-ttu-id="088e5-816">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-816">&lt;optional&gt;</span></span> | <span data-ttu-id="088e5-817">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-817">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="088e5-818">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-818">Requirements</span></span>

|<span data-ttu-id="088e5-819">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-819">Requirement</span></span>| <span data-ttu-id="088e5-820">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-820">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-821">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-822">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-822">1.0</span></span>|
|[<span data-ttu-id="088e5-823">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-823">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-824">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-824">ReadItem</span></span>|
|[<span data-ttu-id="088e5-825">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-825">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-826">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-826">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="088e5-827">Exemples</span><span class="sxs-lookup"><span data-stu-id="088e5-827">Examples</span></span>

<span data-ttu-id="088e5-828">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="088e5-828">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="088e5-829">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="088e5-829">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="088e5-830">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="088e5-830">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="088e5-831">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="088e5-831">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="088e5-832">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-832">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="088e5-833">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-833">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="088e5-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="088e5-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="088e5-835">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="088e5-835">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-836">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-836">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-837">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-837">Requirements</span></span>

|<span data-ttu-id="088e5-838">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-838">Requirement</span></span>| <span data-ttu-id="088e5-839">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-840">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-841">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-841">1.0</span></span>|
|[<span data-ttu-id="088e5-842">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-843">ReadItem</span></span>|
|[<span data-ttu-id="088e5-844">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-845">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="088e5-846">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="088e5-846">Returns:</span></span>

<span data-ttu-id="088e5-847">Type : [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="088e5-847">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="088e5-848">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-848">Example</span></span>

<span data-ttu-id="088e5-849">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-849">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="088e5-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="088e5-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="088e5-851">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="088e5-851">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-852">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-852">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-853">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-853">Parameters:</span></span>

|<span data-ttu-id="088e5-854">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-854">Name</span></span>| <span data-ttu-id="088e5-855">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-855">Type</span></span>| <span data-ttu-id="088e5-856">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-856">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="088e5-857">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="088e5-857">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="088e5-858">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="088e5-858">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="088e5-859">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-859">Requirements</span></span>

|<span data-ttu-id="088e5-860">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-860">Requirement</span></span>| <span data-ttu-id="088e5-861">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-861">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-862">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-862">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-863">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-863">1.0</span></span>|
|[<span data-ttu-id="088e5-864">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-864">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-865">Restreinte</span><span class="sxs-lookup"><span data-stu-id="088e5-865">Restricted</span></span>|
|[<span data-ttu-id="088e5-866">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-866">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-867">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-867">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="088e5-868">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="088e5-868">Returns:</span></span>

<span data-ttu-id="088e5-869">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="088e5-869">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="088e5-870">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="088e5-870">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="088e5-871">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="088e5-871">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="088e5-872">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="088e5-872">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="088e5-873">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="088e5-873">Value of `entityType`</span></span> | <span data-ttu-id="088e5-874">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="088e5-874">Type of objects in returned array</span></span> | <span data-ttu-id="088e5-875">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="088e5-875">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="088e5-876">String</span><span class="sxs-lookup"><span data-stu-id="088e5-876">String</span></span> | <span data-ttu-id="088e5-877">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="088e5-877">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="088e5-878">Contact</span><span class="sxs-lookup"><span data-stu-id="088e5-878">Contact</span></span> | <span data-ttu-id="088e5-879">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="088e5-879">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="088e5-880">String</span><span class="sxs-lookup"><span data-stu-id="088e5-880">String</span></span> | <span data-ttu-id="088e5-881">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="088e5-881">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="088e5-882">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="088e5-882">MeetingSuggestion</span></span> | <span data-ttu-id="088e5-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="088e5-883">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="088e5-884">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="088e5-884">PhoneNumber</span></span> | <span data-ttu-id="088e5-885">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="088e5-885">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="088e5-886">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="088e5-886">TaskSuggestion</span></span> | <span data-ttu-id="088e5-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="088e5-887">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="088e5-888">String</span><span class="sxs-lookup"><span data-stu-id="088e5-888">String</span></span> | <span data-ttu-id="088e5-889">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="088e5-889">**Restricted**</span></span> |

<span data-ttu-id="088e5-890">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="088e5-890">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="088e5-891">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-891">Example</span></span>

<span data-ttu-id="088e5-892">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="088e5-892">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="088e5-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="088e5-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="088e5-894">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="088e5-894">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-895">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-895">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="088e5-896">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="088e5-896">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-897">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-897">Parameters:</span></span>

|<span data-ttu-id="088e5-898">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-898">Name</span></span>| <span data-ttu-id="088e5-899">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-899">Type</span></span>| <span data-ttu-id="088e5-900">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-900">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="088e5-901">String</span><span class="sxs-lookup"><span data-stu-id="088e5-901">String</span></span>|<span data-ttu-id="088e5-902">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="088e5-902">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="088e5-903">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-903">Requirements</span></span>

|<span data-ttu-id="088e5-904">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-904">Requirement</span></span>| <span data-ttu-id="088e5-905">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-906">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-907">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-907">1.0</span></span>|
|[<span data-ttu-id="088e5-908">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-909">ReadItem</span></span>|
|[<span data-ttu-id="088e5-910">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-911">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="088e5-912">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="088e5-912">Returns:</span></span>

<span data-ttu-id="088e5-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="088e5-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="088e5-915">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="088e5-915">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="088e5-916">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="088e5-916">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="088e5-917">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="088e5-917">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-918">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-918">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="088e5-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="088e5-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="088e5-922">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="088e5-922">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="088e5-923">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="088e5-923">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="088e5-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="088e5-927">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-927">Requirements</span></span>

|<span data-ttu-id="088e5-928">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-928">Requirement</span></span>| <span data-ttu-id="088e5-929">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-930">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-931">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-931">1.0</span></span>|
|[<span data-ttu-id="088e5-932">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-932">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-933">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-933">ReadItem</span></span>|
|[<span data-ttu-id="088e5-934">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-934">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-935">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-935">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="088e5-936">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="088e5-936">Returns:</span></span>

<span data-ttu-id="088e5-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="088e5-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="088e5-939">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="088e5-939">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="088e5-940">Object</span><span class="sxs-lookup"><span data-stu-id="088e5-940">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="088e5-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-941">Example</span></span>

<span data-ttu-id="088e5-942">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="088e5-942">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="088e5-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="088e5-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="088e5-944">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="088e5-944">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-945">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="088e5-945">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="088e5-946">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="088e5-946">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="088e5-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="088e5-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-949">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-949">Parameters:</span></span>

|<span data-ttu-id="088e5-950">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-950">Name</span></span>| <span data-ttu-id="088e5-951">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-951">Type</span></span>| <span data-ttu-id="088e5-952">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-952">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="088e5-953">String</span><span class="sxs-lookup"><span data-stu-id="088e5-953">String</span></span>|<span data-ttu-id="088e5-954">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="088e5-954">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="088e5-955">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-955">Requirements</span></span>

|<span data-ttu-id="088e5-956">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-956">Requirement</span></span>| <span data-ttu-id="088e5-957">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-958">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-959">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-959">1.0</span></span>|
|[<span data-ttu-id="088e5-960">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-961">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-961">ReadItem</span></span>|
|[<span data-ttu-id="088e5-962">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-963">Lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="088e5-964">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="088e5-964">Returns:</span></span>

<span data-ttu-id="088e5-965">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="088e5-965">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="088e5-966">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="088e5-966">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="088e5-967">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="088e5-967">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="088e5-968">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-968">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="088e5-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="088e5-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="088e5-970">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="088e5-970">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="088e5-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="088e5-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-973">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-973">Parameters:</span></span>

|<span data-ttu-id="088e5-974">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-974">Name</span></span>| <span data-ttu-id="088e5-975">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-975">Type</span></span>| <span data-ttu-id="088e5-976">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-976">Attributes</span></span>| <span data-ttu-id="088e5-977">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-977">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="088e5-978">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="088e5-978">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="088e5-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="088e5-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="088e5-982">Object</span><span class="sxs-lookup"><span data-stu-id="088e5-982">Object</span></span>| <span data-ttu-id="088e5-983">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-983">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-984">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="088e5-984">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="088e5-985">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-985">Object</span></span>| <span data-ttu-id="088e5-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-986">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-987">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-987">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="088e5-988">fonction</span><span class="sxs-lookup"><span data-stu-id="088e5-988">function</span></span>||<span data-ttu-id="088e5-989">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-989">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="088e5-990">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="088e5-990">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="088e5-991">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="088e5-991">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="088e5-992">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-992">Requirements</span></span>

|<span data-ttu-id="088e5-993">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-993">Requirement</span></span>| <span data-ttu-id="088e5-994">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-994">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-995">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-995">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-996">1.2</span><span class="sxs-lookup"><span data-stu-id="088e5-996">1.2</span></span>|
|[<span data-ttu-id="088e5-997">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-997">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-998">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="088e5-998">ReadWriteItem</span></span>|
|[<span data-ttu-id="088e5-999">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-999">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-1000">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-1000">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="088e5-1001">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="088e5-1001">Returns:</span></span>

<span data-ttu-id="088e5-1002">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="088e5-1002">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="088e5-1003">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="088e5-1003">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="088e5-1004">String</span><span class="sxs-lookup"><span data-stu-id="088e5-1004">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="088e5-1005">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-1005">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="088e5-1006">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="088e5-1006">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="088e5-1007">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="088e5-1007">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="088e5-p163">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="088e5-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-1011">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-1011">Parameters:</span></span>

|<span data-ttu-id="088e5-1012">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-1012">Name</span></span>| <span data-ttu-id="088e5-1013">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-1013">Type</span></span>| <span data-ttu-id="088e5-1014">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-1014">Attributes</span></span>| <span data-ttu-id="088e5-1015">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-1015">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="088e5-1016">function</span><span class="sxs-lookup"><span data-stu-id="088e5-1016">function</span></span>||<span data-ttu-id="088e5-1017">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-1017">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="088e5-1018">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="088e5-1018">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="088e5-1019">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="088e5-1019">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="088e5-1020">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1020">Object</span></span>| <span data-ttu-id="088e5-1021">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1021">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1022">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-1022">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="088e5-1023">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-1023">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="088e5-1024">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-1024">Requirements</span></span>

|<span data-ttu-id="088e5-1025">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-1025">Requirement</span></span>| <span data-ttu-id="088e5-1026">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-1026">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-1027">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-1027">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-1028">1.0</span><span class="sxs-lookup"><span data-stu-id="088e5-1028">1.0</span></span>|
|[<span data-ttu-id="088e5-1029">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-1029">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-1030">ReadItem</span><span class="sxs-lookup"><span data-stu-id="088e5-1030">ReadItem</span></span>|
|[<span data-ttu-id="088e5-1031">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-1031">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-1032">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="088e5-1032">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-1033">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-1033">Example</span></span>

<span data-ttu-id="088e5-p166">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="088e5-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="088e5-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="088e5-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="088e5-1038">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="088e5-1038">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="088e5-p167">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="088e5-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-1043">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-1043">Parameters:</span></span>

|<span data-ttu-id="088e5-1044">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-1044">Name</span></span>| <span data-ttu-id="088e5-1045">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-1045">Type</span></span>| <span data-ttu-id="088e5-1046">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-1046">Attributes</span></span>| <span data-ttu-id="088e5-1047">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-1047">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="088e5-1048">String</span><span class="sxs-lookup"><span data-stu-id="088e5-1048">String</span></span>||<span data-ttu-id="088e5-1049">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="088e5-1049">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="088e5-1050">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1050">Object</span></span>| <span data-ttu-id="088e5-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1052">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="088e5-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="088e5-1053">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1053">Object</span></span>| <span data-ttu-id="088e5-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1055">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="088e5-1056">fonction</span><span class="sxs-lookup"><span data-stu-id="088e5-1056">function</span></span>| <span data-ttu-id="088e5-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1058">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="088e5-1059">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="088e5-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="088e5-1060">Erreurs</span><span class="sxs-lookup"><span data-stu-id="088e5-1060">Errors</span></span>

| <span data-ttu-id="088e5-1061">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="088e5-1061">Error code</span></span> | <span data-ttu-id="088e5-1062">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="088e5-1063">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="088e5-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="088e5-1064">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-1064">Requirements</span></span>

|<span data-ttu-id="088e5-1065">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-1065">Requirement</span></span>| <span data-ttu-id="088e5-1066">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-1067">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="088e5-1068">1.1</span></span>|
|[<span data-ttu-id="088e5-1069">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="088e5-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="088e5-1071">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-1072">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-1073">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-1073">Example</span></span>

<span data-ttu-id="088e5-1074">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="088e5-1074">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="088e5-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="088e5-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="088e5-1076">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="088e5-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="088e5-p168">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="088e5-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-1080">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="088e5-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="088e5-1081">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="088e5-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="088e5-p170">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="088e5-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="088e5-1085">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="088e5-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="088e5-1086">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="088e5-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="088e5-1087">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="088e5-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="088e5-1088">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="088e5-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-1089">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-1089">Parameters:</span></span>

|<span data-ttu-id="088e5-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-1090">Name</span></span>| <span data-ttu-id="088e5-1091">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-1091">Type</span></span>| <span data-ttu-id="088e5-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-1092">Attributes</span></span>| <span data-ttu-id="088e5-1093">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="088e5-1094">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1094">Object</span></span>| <span data-ttu-id="088e5-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1096">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="088e5-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="088e5-1097">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1097">Object</span></span>| <span data-ttu-id="088e5-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1099">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="088e5-1100">fonction</span><span class="sxs-lookup"><span data-stu-id="088e5-1100">function</span></span>||<span data-ttu-id="088e5-1101">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="088e5-1102">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="088e5-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="088e5-1103">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-1103">Requirements</span></span>

|<span data-ttu-id="088e5-1104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-1104">Requirement</span></span>| <span data-ttu-id="088e5-1105">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-1106">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="088e5-1107">1.3</span></span>|
|[<span data-ttu-id="088e5-1108">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="088e5-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="088e5-1110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-1111">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="088e5-1112">範例</span><span class="sxs-lookup"><span data-stu-id="088e5-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="088e5-p172">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="088e5-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="088e5-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="088e5-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="088e5-1116">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="088e5-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="088e5-p173">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="088e5-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="088e5-1120">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="088e5-1120">Parameters:</span></span>

|<span data-ttu-id="088e5-1121">Nom</span><span class="sxs-lookup"><span data-stu-id="088e5-1121">Name</span></span>| <span data-ttu-id="088e5-1122">Type</span><span class="sxs-lookup"><span data-stu-id="088e5-1122">Type</span></span>| <span data-ttu-id="088e5-1123">Attributs</span><span class="sxs-lookup"><span data-stu-id="088e5-1123">Attributes</span></span>| <span data-ttu-id="088e5-1124">Description</span><span class="sxs-lookup"><span data-stu-id="088e5-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="088e5-1125">String</span><span class="sxs-lookup"><span data-stu-id="088e5-1125">String</span></span>||<span data-ttu-id="088e5-p174">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="088e5-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="088e5-1129">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1129">Object</span></span>| <span data-ttu-id="088e5-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1131">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="088e5-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="088e5-1132">Objet</span><span class="sxs-lookup"><span data-stu-id="088e5-1132">Object</span></span>| <span data-ttu-id="088e5-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-1134">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="088e5-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="088e5-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="088e5-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="088e5-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="088e5-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="088e5-p175">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="088e5-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="088e5-p176">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="088e5-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="088e5-1141">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="088e5-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="088e5-1142">fonction</span><span class="sxs-lookup"><span data-stu-id="088e5-1142">function</span></span>||<span data-ttu-id="088e5-1143">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="088e5-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="088e5-1144">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="088e5-1144">Requirements</span></span>

|<span data-ttu-id="088e5-1145">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="088e5-1145">Requirement</span></span>| <span data-ttu-id="088e5-1146">Valeur</span><span class="sxs-lookup"><span data-stu-id="088e5-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="088e5-1147">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="088e5-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="088e5-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="088e5-1148">1.2</span></span>|
|[<span data-ttu-id="088e5-1149">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="088e5-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="088e5-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="088e5-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="088e5-1151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="088e5-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="088e5-1152">Composition</span><span class="sxs-lookup"><span data-stu-id="088e5-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="088e5-1153">Exemple</span><span class="sxs-lookup"><span data-stu-id="088e5-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
