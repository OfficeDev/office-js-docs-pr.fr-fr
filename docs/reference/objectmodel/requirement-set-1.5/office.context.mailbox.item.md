---
title: Office.context.mailbox.item - ensemble de conditions requises 1.5
description: ''
ms.date: 01/30/2019
localization_priority: Priority
ms.openlocfilehash: cca0bb4baa15d72a58909ca1417eb52a9bf70a8f
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701840"
---
# <a name="item"></a><span data-ttu-id="00a35-102">élément</span><span class="sxs-lookup"><span data-stu-id="00a35-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="00a35-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="00a35-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="00a35-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="00a35-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-106">Requirements</span></span>

|<span data-ttu-id="00a35-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-107">Requirement</span></span>| <span data-ttu-id="00a35-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-110">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-110">1.0</span></span>|
|[<span data-ttu-id="00a35-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="00a35-112">Restricted</span></span>|
|[<span data-ttu-id="00a35-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="00a35-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="00a35-115">Members and methods</span></span>

| <span data-ttu-id="00a35-116">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-116">Member</span></span> | <span data-ttu-id="00a35-117">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="00a35-118">attachments</span><span class="sxs-lookup"><span data-stu-id="00a35-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="00a35-119">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-119">Member</span></span> |
| [<span data-ttu-id="00a35-120">bcc</span><span class="sxs-lookup"><span data-stu-id="00a35-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="00a35-121">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-121">Member</span></span> |
| [<span data-ttu-id="00a35-122">body</span><span class="sxs-lookup"><span data-stu-id="00a35-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="00a35-123">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-123">Member</span></span> |
| [<span data-ttu-id="00a35-124">cc</span><span class="sxs-lookup"><span data-stu-id="00a35-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="00a35-125">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-125">Member</span></span> |
| [<span data-ttu-id="00a35-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="00a35-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="00a35-127">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-127">Member</span></span> |
| [<span data-ttu-id="00a35-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="00a35-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="00a35-129">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-129">Member</span></span> |
| [<span data-ttu-id="00a35-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="00a35-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="00a35-131">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-131">Member</span></span> |
| [<span data-ttu-id="00a35-132">end</span><span class="sxs-lookup"><span data-stu-id="00a35-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="00a35-133">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-133">Member</span></span> |
| [<span data-ttu-id="00a35-134">from</span><span class="sxs-lookup"><span data-stu-id="00a35-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="00a35-135">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-135">Member</span></span> |
| [<span data-ttu-id="00a35-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="00a35-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="00a35-137">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-137">Member</span></span> |
| [<span data-ttu-id="00a35-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="00a35-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="00a35-139">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-139">Member</span></span> |
| [<span data-ttu-id="00a35-140">itemId</span><span class="sxs-lookup"><span data-stu-id="00a35-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="00a35-141">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-141">Member</span></span> |
| [<span data-ttu-id="00a35-142">itemType</span><span class="sxs-lookup"><span data-stu-id="00a35-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="00a35-143">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-143">Member</span></span> |
| [<span data-ttu-id="00a35-144">location</span><span class="sxs-lookup"><span data-stu-id="00a35-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="00a35-145">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-145">Member</span></span> |
| [<span data-ttu-id="00a35-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="00a35-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="00a35-147">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-147">Member</span></span> |
| [<span data-ttu-id="00a35-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="00a35-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="00a35-149">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-149">Member</span></span> |
| [<span data-ttu-id="00a35-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="00a35-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="00a35-151">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-151">Member</span></span> |
| [<span data-ttu-id="00a35-152">organizer</span><span class="sxs-lookup"><span data-stu-id="00a35-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="00a35-153">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-153">Member</span></span> |
| [<span data-ttu-id="00a35-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="00a35-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="00a35-155">Member</span><span class="sxs-lookup"><span data-stu-id="00a35-155">Member</span></span> |
| [<span data-ttu-id="00a35-156">sender</span><span class="sxs-lookup"><span data-stu-id="00a35-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="00a35-157">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-157">Member</span></span> |
| [<span data-ttu-id="00a35-158">start</span><span class="sxs-lookup"><span data-stu-id="00a35-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="00a35-159">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-159">Member</span></span> |
| [<span data-ttu-id="00a35-160">subject</span><span class="sxs-lookup"><span data-stu-id="00a35-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="00a35-161">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-161">Member</span></span> |
| [<span data-ttu-id="00a35-162">to</span><span class="sxs-lookup"><span data-stu-id="00a35-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="00a35-163">Membre</span><span class="sxs-lookup"><span data-stu-id="00a35-163">Member</span></span> |
| [<span data-ttu-id="00a35-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="00a35-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-165">Method</span></span> |
| [<span data-ttu-id="00a35-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="00a35-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-167">Method</span></span> |
| [<span data-ttu-id="00a35-168">close</span><span class="sxs-lookup"><span data-stu-id="00a35-168">close</span></span>](#close) | <span data-ttu-id="00a35-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-169">Method</span></span> |
| [<span data-ttu-id="00a35-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="00a35-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="00a35-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-171">Method</span></span> |
| [<span data-ttu-id="00a35-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="00a35-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="00a35-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-173">Method</span></span> |
| [<span data-ttu-id="00a35-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="00a35-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="00a35-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-175">Method</span></span> |
| [<span data-ttu-id="00a35-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="00a35-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="00a35-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-177">Method</span></span> |
| [<span data-ttu-id="00a35-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="00a35-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="00a35-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-179">Method</span></span> |
| [<span data-ttu-id="00a35-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="00a35-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="00a35-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-181">Method</span></span> |
| [<span data-ttu-id="00a35-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="00a35-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="00a35-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-183">Method</span></span> |
| [<span data-ttu-id="00a35-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="00a35-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-185">Method</span></span> |
| [<span data-ttu-id="00a35-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="00a35-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-187">Method</span></span> |
| [<span data-ttu-id="00a35-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="00a35-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-189">Method</span></span> |
| [<span data-ttu-id="00a35-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="00a35-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-191">Method</span></span> |
| [<span data-ttu-id="00a35-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="00a35-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="00a35-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="00a35-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="00a35-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-194">Example</span></span>

<span data-ttu-id="00a35-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="00a35-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="00a35-196">Membres</span><span class="sxs-lookup"><span data-stu-id="00a35-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="00a35-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="00a35-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="00a35-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="00a35-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="00a35-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="00a35-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-202">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-202">Type:</span></span>

*   <span data-ttu-id="00a35-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="00a35-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-204">Requirements</span></span>

|<span data-ttu-id="00a35-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-205">Requirement</span></span>| <span data-ttu-id="00a35-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-208">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-208">1.0</span></span>|
|[<span data-ttu-id="00a35-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-210">ReadItem</span></span>|
|[<span data-ttu-id="00a35-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-213">Example</span></span>

<span data-ttu-id="00a35-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00a35-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00a35-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="00a35-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="00a35-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-218">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-218">Type:</span></span>

*   [<span data-ttu-id="00a35-219">Destinataires</span><span class="sxs-lookup"><span data-stu-id="00a35-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="00a35-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-220">Requirements</span></span>

|<span data-ttu-id="00a35-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-221">Requirement</span></span>| <span data-ttu-id="00a35-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-224">1.1</span><span class="sxs-lookup"><span data-stu-id="00a35-224">1.1</span></span>|
|[<span data-ttu-id="00a35-225">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-226">ReadItem</span></span>|
|[<span data-ttu-id="00a35-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-228">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-229">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="00a35-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="00a35-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="00a35-231">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-232">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-232">Type:</span></span>

*   [<span data-ttu-id="00a35-233">Corps</span><span class="sxs-lookup"><span data-stu-id="00a35-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="00a35-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-234">Requirements</span></span>

|<span data-ttu-id="00a35-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-235">Requirement</span></span>| <span data-ttu-id="00a35-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-238">1.1</span><span class="sxs-lookup"><span data-stu-id="00a35-238">1.1</span></span>|
|[<span data-ttu-id="00a35-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-240">ReadItem</span></span>|
|[<span data-ttu-id="00a35-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-242">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-242">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00a35-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00a35-244">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="00a35-244">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="00a35-245">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-245">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-246">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-246">Read mode</span></span>

<span data-ttu-id="00a35-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="00a35-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-249">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-249">Compose mode</span></span>

<span data-ttu-id="00a35-250">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="00a35-250">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-251">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-251">Type:</span></span>

*   <span data-ttu-id="00a35-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-253">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-253">Requirements</span></span>

|<span data-ttu-id="00a35-254">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-254">Requirement</span></span>| <span data-ttu-id="00a35-255">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-256">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-257">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-257">1.0</span></span>|
|[<span data-ttu-id="00a35-258">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-258">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-259">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-259">ReadItem</span></span>|
|[<span data-ttu-id="00a35-260">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-260">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-261">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-261">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-262">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-262">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="00a35-263">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="00a35-263">(nullable) conversationId :String</span></span>

<span data-ttu-id="00a35-264">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="00a35-264">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="00a35-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="00a35-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="00a35-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="00a35-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-269">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-269">Type:</span></span>

*   <span data-ttu-id="00a35-270">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-270">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-271">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-271">Requirements</span></span>

|<span data-ttu-id="00a35-272">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-272">Requirement</span></span>| <span data-ttu-id="00a35-273">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-273">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-274">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-275">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-275">1.0</span></span>|
|[<span data-ttu-id="00a35-276">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-277">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-277">ReadItem</span></span>|
|[<span data-ttu-id="00a35-278">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-279">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-279">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="00a35-280">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="00a35-280">dateTimeCreated :Date</span></span>

<span data-ttu-id="00a35-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-283">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-283">Type:</span></span>

*   <span data-ttu-id="00a35-284">Date</span><span class="sxs-lookup"><span data-stu-id="00a35-284">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-285">Requirements</span></span>

|<span data-ttu-id="00a35-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-286">Requirement</span></span>| <span data-ttu-id="00a35-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-289">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-289">1.0</span></span>|
|[<span data-ttu-id="00a35-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-291">ReadItem</span></span>|
|[<span data-ttu-id="00a35-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-293">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-293">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-294">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-294">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="00a35-295">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="00a35-295">dateTimeModified :Date</span></span>

<span data-ttu-id="00a35-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-298">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-298">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-299">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-299">Type:</span></span>

*   <span data-ttu-id="00a35-300">Date</span><span class="sxs-lookup"><span data-stu-id="00a35-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-301">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-301">Requirements</span></span>

|<span data-ttu-id="00a35-302">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-302">Requirement</span></span>| <span data-ttu-id="00a35-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-304">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-305">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-305">1.0</span></span>|
|[<span data-ttu-id="00a35-306">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-307">ReadItem</span></span>|
|[<span data-ttu-id="00a35-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-310">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="00a35-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00a35-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="00a35-312">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-312">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="00a35-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="00a35-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-315">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-315">Read mode</span></span>

<span data-ttu-id="00a35-316">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="00a35-316">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-317">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-317">Compose mode</span></span>

<span data-ttu-id="00a35-318">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="00a35-318">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="00a35-319">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="00a35-319">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-320">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-320">Type:</span></span>

*   <span data-ttu-id="00a35-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00a35-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-322">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-322">Requirements</span></span>

|<span data-ttu-id="00a35-323">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-323">Requirement</span></span>| <span data-ttu-id="00a35-324">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-325">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-326">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-326">1.0</span></span>|
|[<span data-ttu-id="00a35-327">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-328">ReadItem</span></span>|
|[<span data-ttu-id="00a35-329">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-330">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-330">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-331">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-331">Example</span></span>

<span data-ttu-id="00a35-332">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="00a35-332">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="00a35-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00a35-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="00a35-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="00a35-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="00a35-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-338">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="00a35-338">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-339">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-339">Type:</span></span>

*   [<span data-ttu-id="00a35-340">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00a35-340">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00a35-341">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-341">Requirements</span></span>

|<span data-ttu-id="00a35-342">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-342">Requirement</span></span>| <span data-ttu-id="00a35-343">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-344">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-345">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-345">1.0</span></span>|
|[<span data-ttu-id="00a35-346">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-347">ReadItem</span></span>|
|[<span data-ttu-id="00a35-348">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-349">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-349">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="00a35-350">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="00a35-350">internetMessageId :String</span></span>

<span data-ttu-id="00a35-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-353">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-353">Type:</span></span>

*   <span data-ttu-id="00a35-354">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-354">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-355">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-355">Requirements</span></span>

|<span data-ttu-id="00a35-356">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-356">Requirement</span></span>| <span data-ttu-id="00a35-357">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-358">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-359">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-359">1.0</span></span>|
|[<span data-ttu-id="00a35-360">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-361">ReadItem</span></span>|
|[<span data-ttu-id="00a35-362">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-363">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-363">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-364">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-364">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="00a35-365">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="00a35-365">itemClass :String</span></span>

<span data-ttu-id="00a35-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="00a35-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="00a35-370">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-370">Type</span></span> | <span data-ttu-id="00a35-371">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-371">Description</span></span> | <span data-ttu-id="00a35-372">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="00a35-372">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="00a35-373">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="00a35-373">Appointment items</span></span> | <span data-ttu-id="00a35-374">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="00a35-374">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="00a35-375">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="00a35-375">Message items</span></span> | <span data-ttu-id="00a35-376">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="00a35-376">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="00a35-377">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="00a35-377">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-378">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-378">Type:</span></span>

*   <span data-ttu-id="00a35-379">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-379">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-380">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-380">Requirements</span></span>

|<span data-ttu-id="00a35-381">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-381">Requirement</span></span>| <span data-ttu-id="00a35-382">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-382">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-383">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-384">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-384">1.0</span></span>|
|[<span data-ttu-id="00a35-385">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-386">ReadItem</span></span>|
|[<span data-ttu-id="00a35-387">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-388">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-388">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-389">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-389">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="00a35-390">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="00a35-390">(nullable) itemId :String</span></span>

<span data-ttu-id="00a35-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-393">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="00a35-393">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="00a35-394">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="00a35-394">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="00a35-395">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="00a35-395">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="00a35-396">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="00a35-396">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="00a35-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-399">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-399">Type:</span></span>

*   <span data-ttu-id="00a35-400">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-400">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-401">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-401">Requirements</span></span>

|<span data-ttu-id="00a35-402">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-402">Requirement</span></span>| <span data-ttu-id="00a35-403">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-403">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-404">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-404">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-405">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-405">1.0</span></span>|
|[<span data-ttu-id="00a35-406">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-406">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-407">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-407">ReadItem</span></span>|
|[<span data-ttu-id="00a35-408">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-408">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-409">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-409">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-410">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-410">Example</span></span>

<span data-ttu-id="00a35-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="00a35-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="00a35-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="00a35-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="00a35-414">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="00a35-414">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="00a35-415">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-415">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-416">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-416">Type:</span></span>

*   [<span data-ttu-id="00a35-417">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="00a35-417">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="00a35-418">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-418">Requirements</span></span>

|<span data-ttu-id="00a35-419">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-419">Requirement</span></span>| <span data-ttu-id="00a35-420">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-421">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-422">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-422">1.0</span></span>|
|[<span data-ttu-id="00a35-423">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-424">ReadItem</span></span>|
|[<span data-ttu-id="00a35-425">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-426">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-426">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-427">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-427">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="00a35-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="00a35-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="00a35-429">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-429">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-430">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-430">Read mode</span></span>

<span data-ttu-id="00a35-431">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-431">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-432">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-432">Compose mode</span></span>

<span data-ttu-id="00a35-433">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-433">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-434">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-434">Type:</span></span>

*   <span data-ttu-id="00a35-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="00a35-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-436">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-436">Requirements</span></span>

|<span data-ttu-id="00a35-437">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-437">Requirement</span></span>| <span data-ttu-id="00a35-438">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-439">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-440">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-440">1.0</span></span>|
|[<span data-ttu-id="00a35-441">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-441">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-442">ReadItem</span></span>|
|[<span data-ttu-id="00a35-443">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-443">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-444">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-444">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-445">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-445">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="00a35-446">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="00a35-446">normalizedSubject :String</span></span>

<span data-ttu-id="00a35-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="00a35-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="00a35-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-451">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-451">Type:</span></span>

*   <span data-ttu-id="00a35-452">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-453">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-453">Requirements</span></span>

|<span data-ttu-id="00a35-454">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-454">Requirement</span></span>| <span data-ttu-id="00a35-455">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-456">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-457">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-457">1.0</span></span>|
|[<span data-ttu-id="00a35-458">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-459">ReadItem</span></span>|
|[<span data-ttu-id="00a35-460">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-461">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-462">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="00a35-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="00a35-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="00a35-464">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-464">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-465">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-465">Type:</span></span>

*   [<span data-ttu-id="00a35-466">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="00a35-466">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="00a35-467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-467">Requirements</span></span>

|<span data-ttu-id="00a35-468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-468">Requirement</span></span>| <span data-ttu-id="00a35-469">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-471">1.3</span><span class="sxs-lookup"><span data-stu-id="00a35-471">1.3</span></span>|
|[<span data-ttu-id="00a35-472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-473">ReadItem</span></span>|
|[<span data-ttu-id="00a35-474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-475">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-475">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00a35-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00a35-477">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="00a35-477">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="00a35-478">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-478">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-479">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-479">Read mode</span></span>

<span data-ttu-id="00a35-480">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="00a35-480">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-481">Compose mode</span></span>

<span data-ttu-id="00a35-482">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="00a35-482">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-483">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-483">Type:</span></span>

*   <span data-ttu-id="00a35-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-485">Requirements</span></span>

|<span data-ttu-id="00a35-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-486">Requirement</span></span>| <span data-ttu-id="00a35-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-489">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-489">1.0</span></span>|
|[<span data-ttu-id="00a35-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-491">ReadItem</span></span>|
|[<span data-ttu-id="00a35-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-493">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-493">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-494">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-494">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="00a35-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00a35-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="00a35-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-498">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-498">Type:</span></span>

*   [<span data-ttu-id="00a35-499">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00a35-499">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00a35-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-500">Requirements</span></span>

|<span data-ttu-id="00a35-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-501">Requirement</span></span>| <span data-ttu-id="00a35-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-504">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-504">1.0</span></span>|
|[<span data-ttu-id="00a35-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-506">ReadItem</span></span>|
|[<span data-ttu-id="00a35-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-508">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-509">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-509">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00a35-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00a35-511">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="00a35-511">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="00a35-512">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-512">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-513">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-513">Read mode</span></span>

<span data-ttu-id="00a35-514">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="00a35-514">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-515">Compose mode</span></span>

<span data-ttu-id="00a35-516">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="00a35-516">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-517">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-517">Type:</span></span>

*   <span data-ttu-id="00a35-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-519">Requirements</span></span>

|<span data-ttu-id="00a35-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-520">Requirement</span></span>| <span data-ttu-id="00a35-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-523">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-523">1.0</span></span>|
|[<span data-ttu-id="00a35-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-525">ReadItem</span></span>|
|[<span data-ttu-id="00a35-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-527">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-528">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-528">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="00a35-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00a35-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="00a35-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="00a35-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="00a35-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="00a35-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-534">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="00a35-534">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-535">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-535">Type:</span></span>

*   [<span data-ttu-id="00a35-536">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00a35-536">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00a35-537">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-537">Requirements</span></span>

|<span data-ttu-id="00a35-538">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-538">Requirement</span></span>| <span data-ttu-id="00a35-539">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-540">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-541">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-541">1.0</span></span>|
|[<span data-ttu-id="00a35-542">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-543">ReadItem</span></span>|
|[<span data-ttu-id="00a35-544">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-545">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-545">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-546">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-546">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="00a35-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00a35-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="00a35-548">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-548">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="00a35-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="00a35-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-551">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-551">Read mode</span></span>

<span data-ttu-id="00a35-552">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="00a35-552">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-553">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-553">Compose mode</span></span>

<span data-ttu-id="00a35-554">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="00a35-554">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="00a35-555">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="00a35-555">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-556">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-556">Type:</span></span>

*   <span data-ttu-id="00a35-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00a35-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-558">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-558">Requirements</span></span>

|<span data-ttu-id="00a35-559">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-559">Requirement</span></span>| <span data-ttu-id="00a35-560">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-561">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-562">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-562">1.0</span></span>|
|[<span data-ttu-id="00a35-563">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-564">ReadItem</span></span>|
|[<span data-ttu-id="00a35-565">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-566">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-567">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-567">Example</span></span>

<span data-ttu-id="00a35-568">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="00a35-568">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="00a35-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="00a35-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="00a35-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="00a35-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="00a35-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-572">Read mode</span></span>

<span data-ttu-id="00a35-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="00a35-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="00a35-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-575">Compose mode</span></span>

<span data-ttu-id="00a35-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="00a35-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00a35-577">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-577">Type:</span></span>

*   <span data-ttu-id="00a35-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="00a35-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-579">Requirements</span></span>

|<span data-ttu-id="00a35-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-580">Requirement</span></span>| <span data-ttu-id="00a35-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-583">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-583">1.0</span></span>|
|[<span data-ttu-id="00a35-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-585">ReadItem</span></span>|
|[<span data-ttu-id="00a35-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-587">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00a35-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00a35-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="00a35-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="00a35-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00a35-591">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-591">Read mode</span></span>

<span data-ttu-id="00a35-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="00a35-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00a35-594">Mode composition</span><span class="sxs-lookup"><span data-stu-id="00a35-594">Compose mode</span></span>

<span data-ttu-id="00a35-595">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="00a35-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="00a35-596">Type :</span><span class="sxs-lookup"><span data-stu-id="00a35-596">Type:</span></span>

*   <span data-ttu-id="00a35-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00a35-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-598">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-598">Requirements</span></span>

|<span data-ttu-id="00a35-599">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-599">Requirement</span></span>| <span data-ttu-id="00a35-600">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-601">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-602">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-602">1.0</span></span>|
|[<span data-ttu-id="00a35-603">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-604">ReadItem</span></span>|
|[<span data-ttu-id="00a35-605">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-606">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-606">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-607">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-607">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="00a35-608">Méthodes</span><span class="sxs-lookup"><span data-stu-id="00a35-608">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="00a35-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00a35-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="00a35-610">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="00a35-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="00a35-611">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="00a35-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="00a35-612">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="00a35-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-613">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-613">Parameters:</span></span>

|<span data-ttu-id="00a35-614">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-614">Name</span></span>| <span data-ttu-id="00a35-615">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-615">Type</span></span>| <span data-ttu-id="00a35-616">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-616">Attributes</span></span>| <span data-ttu-id="00a35-617">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="00a35-618">String</span><span class="sxs-lookup"><span data-stu-id="00a35-618">String</span></span>||<span data-ttu-id="00a35-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="00a35-621">String</span><span class="sxs-lookup"><span data-stu-id="00a35-621">String</span></span>||<span data-ttu-id="00a35-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="00a35-624">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-624">Object</span></span>| <span data-ttu-id="00a35-625">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-625">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-626">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="00a35-626">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="00a35-627">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-627">Object</span></span> | <span data-ttu-id="00a35-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-628">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-629">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="00a35-630">Boolean</span><span class="sxs-lookup"><span data-stu-id="00a35-630">Boolean</span></span> | <span data-ttu-id="00a35-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-631">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-632">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="00a35-632">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="00a35-633">fonction</span><span class="sxs-lookup"><span data-stu-id="00a35-633">function</span></span>| <span data-ttu-id="00a35-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-634">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-635">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00a35-636">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00a35-636">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="00a35-637">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="00a35-637">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00a35-638">Erreurs</span><span class="sxs-lookup"><span data-stu-id="00a35-638">Errors</span></span>

| <span data-ttu-id="00a35-639">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="00a35-639">Error code</span></span> | <span data-ttu-id="00a35-640">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-640">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="00a35-641">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="00a35-641">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="00a35-642">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="00a35-642">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="00a35-643">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="00a35-643">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00a35-644">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-644">Requirements</span></span>

|<span data-ttu-id="00a35-645">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-645">Requirement</span></span>| <span data-ttu-id="00a35-646">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-646">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-647">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-648">1.1</span><span class="sxs-lookup"><span data-stu-id="00a35-648">1.1</span></span>|
|[<span data-ttu-id="00a35-649">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-649">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00a35-650">ReadWriteItem</span></span>|
|[<span data-ttu-id="00a35-651">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-651">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-652">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-652">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="00a35-653">Exemples</span><span class="sxs-lookup"><span data-stu-id="00a35-653">Examples</span></span>

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

<span data-ttu-id="00a35-654">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="00a35-654">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="00a35-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00a35-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="00a35-656">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="00a35-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="00a35-660">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="00a35-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="00a35-661">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="00a35-661">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-662">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-662">Parameters:</span></span>

|<span data-ttu-id="00a35-663">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-663">Name</span></span>| <span data-ttu-id="00a35-664">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-664">Type</span></span>| <span data-ttu-id="00a35-665">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-665">Attributes</span></span>| <span data-ttu-id="00a35-666">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="00a35-667">String</span><span class="sxs-lookup"><span data-stu-id="00a35-667">String</span></span>||<span data-ttu-id="00a35-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="00a35-670">String</span><span class="sxs-lookup"><span data-stu-id="00a35-670">String</span></span>||<span data-ttu-id="00a35-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="00a35-673">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-673">Object</span></span>| <span data-ttu-id="00a35-674">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-674">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-675">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="00a35-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00a35-676">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-676">Object</span></span>| <span data-ttu-id="00a35-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-677">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-678">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00a35-679">fonction</span><span class="sxs-lookup"><span data-stu-id="00a35-679">function</span></span>| <span data-ttu-id="00a35-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-680">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-681">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00a35-682">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00a35-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="00a35-683">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="00a35-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00a35-684">Erreurs</span><span class="sxs-lookup"><span data-stu-id="00a35-684">Errors</span></span>

| <span data-ttu-id="00a35-685">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="00a35-685">Error code</span></span> | <span data-ttu-id="00a35-686">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="00a35-687">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="00a35-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00a35-688">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-688">Requirements</span></span>

|<span data-ttu-id="00a35-689">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-689">Requirement</span></span>| <span data-ttu-id="00a35-690">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-691">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-692">1.1</span><span class="sxs-lookup"><span data-stu-id="00a35-692">1.1</span></span>|
|[<span data-ttu-id="00a35-693">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00a35-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="00a35-695">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-696">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-697">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-697">Example</span></span>

<span data-ttu-id="00a35-698">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="00a35-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="00a35-699">close()</span><span class="sxs-lookup"><span data-stu-id="00a35-699">close()</span></span>

<span data-ttu-id="00a35-700">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="00a35-700">Closes the current item that is being composed.</span></span>

<span data-ttu-id="00a35-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="00a35-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-703">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-703">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="00a35-704">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="00a35-704">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-705">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-705">Requirements</span></span>

|<span data-ttu-id="00a35-706">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-706">Requirement</span></span>| <span data-ttu-id="00a35-707">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-707">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-708">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-708">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-709">1.3</span><span class="sxs-lookup"><span data-stu-id="00a35-709">1.3</span></span>|
|[<span data-ttu-id="00a35-710">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-710">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-711">Restreinte</span><span class="sxs-lookup"><span data-stu-id="00a35-711">Restricted</span></span>|
|[<span data-ttu-id="00a35-712">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-712">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-713">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-713">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="00a35-714">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="00a35-714">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="00a35-715">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="00a35-715">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-716">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00a35-717">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="00a35-717">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="00a35-718">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="00a35-718">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="00a35-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="00a35-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-722">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-722">Parameters:</span></span>

| <span data-ttu-id="00a35-723">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-723">Name</span></span> | <span data-ttu-id="00a35-724">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-724">Type</span></span> | <span data-ttu-id="00a35-725">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-725">Attributes</span></span> | <span data-ttu-id="00a35-726">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-726">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="00a35-727">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="00a35-727">String &#124; Object</span></span>| |<span data-ttu-id="00a35-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="00a35-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="00a35-730">**OU**</span><span class="sxs-lookup"><span data-stu-id="00a35-730">**OR**</span></span><br/><span data-ttu-id="00a35-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="00a35-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="00a35-733">String</span><span class="sxs-lookup"><span data-stu-id="00a35-733">String</span></span> | <span data-ttu-id="00a35-734">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-734">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="00a35-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="00a35-737">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-737">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="00a35-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-738">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-739">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-739">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="00a35-740">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-740">String</span></span> | | <span data-ttu-id="00a35-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="00a35-743">String</span><span class="sxs-lookup"><span data-stu-id="00a35-743">String</span></span> | | <span data-ttu-id="00a35-744">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-744">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="00a35-745">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-745">String</span></span> | | <span data-ttu-id="00a35-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="00a35-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="00a35-748">Booléen</span><span class="sxs-lookup"><span data-stu-id="00a35-748">Boolean</span></span> | | <span data-ttu-id="00a35-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="00a35-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="00a35-751">String</span><span class="sxs-lookup"><span data-stu-id="00a35-751">String</span></span> | | <span data-ttu-id="00a35-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="00a35-755">function</span><span class="sxs-lookup"><span data-stu-id="00a35-755">function</span></span> | <span data-ttu-id="00a35-756">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-756">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-757">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-757">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00a35-758">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-758">Requirements</span></span>

|<span data-ttu-id="00a35-759">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-759">Requirement</span></span>| <span data-ttu-id="00a35-760">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-761">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-762">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-762">1.0</span></span>|
|[<span data-ttu-id="00a35-763">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-764">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-764">ReadItem</span></span>|
|[<span data-ttu-id="00a35-765">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-766">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-766">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="00a35-767">Exemples</span><span class="sxs-lookup"><span data-stu-id="00a35-767">Examples</span></span>

<span data-ttu-id="00a35-768">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="00a35-768">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="00a35-769">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="00a35-769">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="00a35-770">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="00a35-770">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="00a35-771">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="00a35-771">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="00a35-772">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-772">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="00a35-773">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-773">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="00a35-774">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="00a35-774">displayReplyForm(formData)</span></span>

<span data-ttu-id="00a35-775">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="00a35-775">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-776">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-776">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00a35-777">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="00a35-777">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="00a35-778">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="00a35-778">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="00a35-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="00a35-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-782">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-782">Parameters:</span></span>

| <span data-ttu-id="00a35-783">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-783">Name</span></span> | <span data-ttu-id="00a35-784">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-784">Type</span></span> | <span data-ttu-id="00a35-785">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-785">Attributes</span></span> | <span data-ttu-id="00a35-786">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-786">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="00a35-787">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="00a35-787">String &#124; Object</span></span>| | <span data-ttu-id="00a35-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="00a35-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="00a35-790">**OU**</span><span class="sxs-lookup"><span data-stu-id="00a35-790">**OR**</span></span><br/><span data-ttu-id="00a35-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="00a35-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="00a35-793">String</span><span class="sxs-lookup"><span data-stu-id="00a35-793">String</span></span> | <span data-ttu-id="00a35-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-794">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="00a35-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="00a35-797">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-797">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="00a35-798">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-798">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-799">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-799">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="00a35-800">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-800">String</span></span> | | <span data-ttu-id="00a35-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="00a35-803">String</span><span class="sxs-lookup"><span data-stu-id="00a35-803">String</span></span> | | <span data-ttu-id="00a35-804">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-804">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="00a35-805">Chaîne</span><span class="sxs-lookup"><span data-stu-id="00a35-805">String</span></span> | | <span data-ttu-id="00a35-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="00a35-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="00a35-808">Booléen</span><span class="sxs-lookup"><span data-stu-id="00a35-808">Boolean</span></span> | | <span data-ttu-id="00a35-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="00a35-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="00a35-811">String</span><span class="sxs-lookup"><span data-stu-id="00a35-811">String</span></span> | | <span data-ttu-id="00a35-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="00a35-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="00a35-815">function</span><span class="sxs-lookup"><span data-stu-id="00a35-815">function</span></span> | <span data-ttu-id="00a35-816">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-816">&lt;optional&gt;</span></span> | <span data-ttu-id="00a35-817">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-817">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00a35-818">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-818">Requirements</span></span>

|<span data-ttu-id="00a35-819">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-819">Requirement</span></span>| <span data-ttu-id="00a35-820">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-820">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-821">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-822">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-822">1.0</span></span>|
|[<span data-ttu-id="00a35-823">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-823">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-824">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-824">ReadItem</span></span>|
|[<span data-ttu-id="00a35-825">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-825">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-826">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-826">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="00a35-827">Exemples</span><span class="sxs-lookup"><span data-stu-id="00a35-827">Examples</span></span>

<span data-ttu-id="00a35-828">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="00a35-828">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="00a35-829">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="00a35-829">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="00a35-830">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="00a35-830">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="00a35-831">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="00a35-831">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="00a35-832">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-832">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="00a35-833">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-833">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="00a35-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="00a35-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="00a35-835">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="00a35-835">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-836">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-836">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-837">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-837">Requirements</span></span>

|<span data-ttu-id="00a35-838">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-838">Requirement</span></span>| <span data-ttu-id="00a35-839">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-840">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-841">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-841">1.0</span></span>|
|[<span data-ttu-id="00a35-842">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-843">ReadItem</span></span>|
|[<span data-ttu-id="00a35-844">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-845">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00a35-846">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="00a35-846">Returns:</span></span>

<span data-ttu-id="00a35-847">Type : [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="00a35-847">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="00a35-848">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-848">Example</span></span>

<span data-ttu-id="00a35-849">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-849">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="00a35-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="00a35-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="00a35-851">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="00a35-851">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-852">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-852">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-853">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-853">Parameters:</span></span>

|<span data-ttu-id="00a35-854">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-854">Name</span></span>| <span data-ttu-id="00a35-855">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-855">Type</span></span>| <span data-ttu-id="00a35-856">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-856">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="00a35-857">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="00a35-857">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="00a35-858">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="00a35-858">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00a35-859">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-859">Requirements</span></span>

|<span data-ttu-id="00a35-860">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-860">Requirement</span></span>| <span data-ttu-id="00a35-861">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-861">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-862">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-862">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-863">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-863">1.0</span></span>|
|[<span data-ttu-id="00a35-864">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-864">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-865">Restreinte</span><span class="sxs-lookup"><span data-stu-id="00a35-865">Restricted</span></span>|
|[<span data-ttu-id="00a35-866">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-866">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-867">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-867">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00a35-868">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="00a35-868">Returns:</span></span>

<span data-ttu-id="00a35-869">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="00a35-869">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="00a35-870">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="00a35-870">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="00a35-871">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="00a35-871">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="00a35-872">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="00a35-872">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="00a35-873">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="00a35-873">Value of `entityType`</span></span> | <span data-ttu-id="00a35-874">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="00a35-874">Type of objects in returned array</span></span> | <span data-ttu-id="00a35-875">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="00a35-875">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="00a35-876">String</span><span class="sxs-lookup"><span data-stu-id="00a35-876">String</span></span> | <span data-ttu-id="00a35-877">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="00a35-877">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="00a35-878">Contact</span><span class="sxs-lookup"><span data-stu-id="00a35-878">Contact</span></span> | <span data-ttu-id="00a35-879">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00a35-879">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="00a35-880">String</span><span class="sxs-lookup"><span data-stu-id="00a35-880">String</span></span> | <span data-ttu-id="00a35-881">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00a35-881">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="00a35-882">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="00a35-882">MeetingSuggestion</span></span> | <span data-ttu-id="00a35-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00a35-883">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="00a35-884">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="00a35-884">PhoneNumber</span></span> | <span data-ttu-id="00a35-885">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="00a35-885">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="00a35-886">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="00a35-886">TaskSuggestion</span></span> | <span data-ttu-id="00a35-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00a35-887">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="00a35-888">String</span><span class="sxs-lookup"><span data-stu-id="00a35-888">String</span></span> | <span data-ttu-id="00a35-889">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="00a35-889">**Restricted**</span></span> |

<span data-ttu-id="00a35-890">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="00a35-890">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="00a35-891">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-891">Example</span></span>

<span data-ttu-id="00a35-892">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="00a35-892">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="00a35-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="00a35-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="00a35-894">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="00a35-894">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-895">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-895">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00a35-896">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="00a35-896">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-897">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-897">Parameters:</span></span>

|<span data-ttu-id="00a35-898">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-898">Name</span></span>| <span data-ttu-id="00a35-899">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-899">Type</span></span>| <span data-ttu-id="00a35-900">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-900">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="00a35-901">String</span><span class="sxs-lookup"><span data-stu-id="00a35-901">String</span></span>|<span data-ttu-id="00a35-902">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="00a35-902">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00a35-903">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-903">Requirements</span></span>

|<span data-ttu-id="00a35-904">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-904">Requirement</span></span>| <span data-ttu-id="00a35-905">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-906">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-907">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-907">1.0</span></span>|
|[<span data-ttu-id="00a35-908">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-909">ReadItem</span></span>|
|[<span data-ttu-id="00a35-910">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-911">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00a35-912">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="00a35-912">Returns:</span></span>

<span data-ttu-id="00a35-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="00a35-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="00a35-915">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="00a35-915">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="00a35-916">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="00a35-916">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="00a35-917">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="00a35-917">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-918">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-918">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00a35-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="00a35-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="00a35-922">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="00a35-922">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="00a35-923">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="00a35-923">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="00a35-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00a35-927">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-927">Requirements</span></span>

|<span data-ttu-id="00a35-928">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-928">Requirement</span></span>| <span data-ttu-id="00a35-929">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-930">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-931">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-931">1.0</span></span>|
|[<span data-ttu-id="00a35-932">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-932">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-933">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-933">ReadItem</span></span>|
|[<span data-ttu-id="00a35-934">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-934">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-935">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-935">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00a35-936">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="00a35-936">Returns:</span></span>

<span data-ttu-id="00a35-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="00a35-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="00a35-939">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="00a35-939">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00a35-940">Object</span><span class="sxs-lookup"><span data-stu-id="00a35-940">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00a35-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-941">Example</span></span>

<span data-ttu-id="00a35-942">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="00a35-942">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="00a35-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="00a35-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="00a35-944">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="00a35-944">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-945">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="00a35-945">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00a35-946">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="00a35-946">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="00a35-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="00a35-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-949">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-949">Parameters:</span></span>

|<span data-ttu-id="00a35-950">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-950">Name</span></span>| <span data-ttu-id="00a35-951">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-951">Type</span></span>| <span data-ttu-id="00a35-952">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-952">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="00a35-953">String</span><span class="sxs-lookup"><span data-stu-id="00a35-953">String</span></span>|<span data-ttu-id="00a35-954">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="00a35-954">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00a35-955">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-955">Requirements</span></span>

|<span data-ttu-id="00a35-956">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-956">Requirement</span></span>| <span data-ttu-id="00a35-957">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-958">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-959">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-959">1.0</span></span>|
|[<span data-ttu-id="00a35-960">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-961">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-961">ReadItem</span></span>|
|[<span data-ttu-id="00a35-962">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-963">Lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00a35-964">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="00a35-964">Returns:</span></span>

<span data-ttu-id="00a35-965">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="00a35-965">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="00a35-966">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="00a35-966">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00a35-967">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="00a35-967">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00a35-968">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-968">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="00a35-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="00a35-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="00a35-970">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="00a35-970">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="00a35-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="00a35-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-973">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-973">Parameters:</span></span>

|<span data-ttu-id="00a35-974">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-974">Name</span></span>| <span data-ttu-id="00a35-975">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-975">Type</span></span>| <span data-ttu-id="00a35-976">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-976">Attributes</span></span>| <span data-ttu-id="00a35-977">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-977">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="00a35-978">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="00a35-978">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="00a35-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="00a35-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="00a35-982">Object</span><span class="sxs-lookup"><span data-stu-id="00a35-982">Object</span></span>| <span data-ttu-id="00a35-983">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-983">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-984">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="00a35-984">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00a35-985">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-985">Object</span></span>| <span data-ttu-id="00a35-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-986">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-987">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-987">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00a35-988">fonction</span><span class="sxs-lookup"><span data-stu-id="00a35-988">function</span></span>||<span data-ttu-id="00a35-989">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-989">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00a35-990">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="00a35-990">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="00a35-991">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="00a35-991">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00a35-992">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-992">Requirements</span></span>

|<span data-ttu-id="00a35-993">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-993">Requirement</span></span>| <span data-ttu-id="00a35-994">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-994">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-995">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-995">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-996">1.2</span><span class="sxs-lookup"><span data-stu-id="00a35-996">1.2</span></span>|
|[<span data-ttu-id="00a35-997">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-997">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-998">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00a35-998">ReadWriteItem</span></span>|
|[<span data-ttu-id="00a35-999">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-999">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-1000">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-1000">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="00a35-1001">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="00a35-1001">Returns:</span></span>

<span data-ttu-id="00a35-1002">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="00a35-1002">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="00a35-1003">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="00a35-1003">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00a35-1004">String</span><span class="sxs-lookup"><span data-stu-id="00a35-1004">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00a35-1005">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-1005">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="00a35-1006">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="00a35-1006">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="00a35-1007">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="00a35-1007">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="00a35-p163">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="00a35-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-1011">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-1011">Parameters:</span></span>

|<span data-ttu-id="00a35-1012">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-1012">Name</span></span>| <span data-ttu-id="00a35-1013">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-1013">Type</span></span>| <span data-ttu-id="00a35-1014">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-1014">Attributes</span></span>| <span data-ttu-id="00a35-1015">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-1015">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="00a35-1016">function</span><span class="sxs-lookup"><span data-stu-id="00a35-1016">function</span></span>||<span data-ttu-id="00a35-1017">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-1017">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00a35-1018">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00a35-1018">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="00a35-1019">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="00a35-1019">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="00a35-1020">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1020">Object</span></span>| <span data-ttu-id="00a35-1021">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1021">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1022">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-1022">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="00a35-1023">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-1023">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00a35-1024">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-1024">Requirements</span></span>

|<span data-ttu-id="00a35-1025">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-1025">Requirement</span></span>| <span data-ttu-id="00a35-1026">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-1026">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-1027">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-1027">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-1028">1.0</span><span class="sxs-lookup"><span data-stu-id="00a35-1028">1.0</span></span>|
|[<span data-ttu-id="00a35-1029">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-1029">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-1030">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00a35-1030">ReadItem</span></span>|
|[<span data-ttu-id="00a35-1031">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-1031">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-1032">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="00a35-1032">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-1033">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-1033">Example</span></span>

<span data-ttu-id="00a35-p166">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="00a35-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="00a35-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00a35-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="00a35-1038">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="00a35-1038">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="00a35-p167">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="00a35-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-1043">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-1043">Parameters:</span></span>

|<span data-ttu-id="00a35-1044">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-1044">Name</span></span>| <span data-ttu-id="00a35-1045">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-1045">Type</span></span>| <span data-ttu-id="00a35-1046">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-1046">Attributes</span></span>| <span data-ttu-id="00a35-1047">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-1047">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="00a35-1048">String</span><span class="sxs-lookup"><span data-stu-id="00a35-1048">String</span></span>||<span data-ttu-id="00a35-1049">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="00a35-1049">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="00a35-1050">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1050">Object</span></span>| <span data-ttu-id="00a35-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1052">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="00a35-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00a35-1053">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1053">Object</span></span>| <span data-ttu-id="00a35-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1055">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00a35-1056">fonction</span><span class="sxs-lookup"><span data-stu-id="00a35-1056">function</span></span>| <span data-ttu-id="00a35-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1058">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00a35-1059">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="00a35-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00a35-1060">Erreurs</span><span class="sxs-lookup"><span data-stu-id="00a35-1060">Errors</span></span>

| <span data-ttu-id="00a35-1061">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="00a35-1061">Error code</span></span> | <span data-ttu-id="00a35-1062">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="00a35-1063">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="00a35-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00a35-1064">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-1064">Requirements</span></span>

|<span data-ttu-id="00a35-1065">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-1065">Requirement</span></span>| <span data-ttu-id="00a35-1066">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-1067">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="00a35-1068">1.1</span></span>|
|[<span data-ttu-id="00a35-1069">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00a35-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="00a35-1071">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-1072">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-1073">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-1073">Example</span></span>

<span data-ttu-id="00a35-1074">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="00a35-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="00a35-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="00a35-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="00a35-1076">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="00a35-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="00a35-p168">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="00a35-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-1080">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="00a35-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="00a35-1081">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="00a35-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="00a35-p170">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="00a35-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="00a35-1085">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="00a35-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="00a35-1086">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="00a35-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="00a35-1087">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="00a35-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="00a35-1088">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="00a35-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-1089">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-1089">Parameters:</span></span>

|<span data-ttu-id="00a35-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-1090">Name</span></span>| <span data-ttu-id="00a35-1091">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-1091">Type</span></span>| <span data-ttu-id="00a35-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-1092">Attributes</span></span>| <span data-ttu-id="00a35-1093">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="00a35-1094">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1094">Object</span></span>| <span data-ttu-id="00a35-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1096">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="00a35-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00a35-1097">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1097">Object</span></span>| <span data-ttu-id="00a35-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1099">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00a35-1100">fonction</span><span class="sxs-lookup"><span data-stu-id="00a35-1100">function</span></span>||<span data-ttu-id="00a35-1101">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00a35-1102">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00a35-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00a35-1103">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-1103">Requirements</span></span>

|<span data-ttu-id="00a35-1104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-1104">Requirement</span></span>| <span data-ttu-id="00a35-1105">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-1106">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="00a35-1107">1.3</span></span>|
|[<span data-ttu-id="00a35-1108">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00a35-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="00a35-1110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-1111">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="00a35-1112">範例</span><span class="sxs-lookup"><span data-stu-id="00a35-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="00a35-p172">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="00a35-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="00a35-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="00a35-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="00a35-1116">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="00a35-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="00a35-p173">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="00a35-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00a35-1120">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="00a35-1120">Parameters:</span></span>

|<span data-ttu-id="00a35-1121">Nom</span><span class="sxs-lookup"><span data-stu-id="00a35-1121">Name</span></span>| <span data-ttu-id="00a35-1122">Type</span><span class="sxs-lookup"><span data-stu-id="00a35-1122">Type</span></span>| <span data-ttu-id="00a35-1123">Attributs</span><span class="sxs-lookup"><span data-stu-id="00a35-1123">Attributes</span></span>| <span data-ttu-id="00a35-1124">Description</span><span class="sxs-lookup"><span data-stu-id="00a35-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="00a35-1125">String</span><span class="sxs-lookup"><span data-stu-id="00a35-1125">String</span></span>||<span data-ttu-id="00a35-p174">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="00a35-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="00a35-1129">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1129">Object</span></span>| <span data-ttu-id="00a35-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1131">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="00a35-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00a35-1132">Objet</span><span class="sxs-lookup"><span data-stu-id="00a35-1132">Object</span></span>| <span data-ttu-id="00a35-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-1134">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="00a35-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="00a35-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="00a35-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="00a35-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="00a35-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="00a35-p175">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="00a35-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="00a35-p176">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="00a35-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="00a35-1141">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="00a35-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="00a35-1142">fonction</span><span class="sxs-lookup"><span data-stu-id="00a35-1142">function</span></span>||<span data-ttu-id="00a35-1143">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00a35-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00a35-1144">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00a35-1144">Requirements</span></span>

|<span data-ttu-id="00a35-1145">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00a35-1145">Requirement</span></span>| <span data-ttu-id="00a35-1146">Valeur</span><span class="sxs-lookup"><span data-stu-id="00a35-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="00a35-1147">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00a35-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00a35-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="00a35-1148">1.2</span></span>|
|[<span data-ttu-id="00a35-1149">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00a35-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00a35-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00a35-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="00a35-1151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00a35-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00a35-1152">Composition</span><span class="sxs-lookup"><span data-stu-id="00a35-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00a35-1153">Exemple</span><span class="sxs-lookup"><span data-stu-id="00a35-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
