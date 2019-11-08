---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,3
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: d0a4d5244a3abeed20282b8b548dedf8f60e7ba5
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066122"
---
# <a name="item"></a><span data-ttu-id="6e76d-102">élément</span><span class="sxs-lookup"><span data-stu-id="6e76d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="6e76d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="6e76d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="6e76d-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="6e76d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-106">Requirements</span></span>

|<span data-ttu-id="6e76d-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-107">Requirement</span></span>| <span data-ttu-id="6e76d-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-110">1.0</span></span>|
|[<span data-ttu-id="6e76d-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="6e76d-112">Restricted</span></span>|
|[<span data-ttu-id="6e76d-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6e76d-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="6e76d-115">Members and methods</span></span>

| <span data-ttu-id="6e76d-116">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-116">Member</span></span> | <span data-ttu-id="6e76d-117">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6e76d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="6e76d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="6e76d-119">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-119">Member</span></span> |
| [<span data-ttu-id="6e76d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="6e76d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="6e76d-121">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-121">Member</span></span> |
| [<span data-ttu-id="6e76d-122">body</span><span class="sxs-lookup"><span data-stu-id="6e76d-122">body</span></span>](#body-body) | <span data-ttu-id="6e76d-123">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-123">Member</span></span> |
| [<span data-ttu-id="6e76d-124">cc</span><span class="sxs-lookup"><span data-stu-id="6e76d-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6e76d-125">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-125">Member</span></span> |
| [<span data-ttu-id="6e76d-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="6e76d-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="6e76d-127">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-127">Member</span></span> |
| [<span data-ttu-id="6e76d-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="6e76d-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="6e76d-129">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-129">Member</span></span> |
| [<span data-ttu-id="6e76d-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="6e76d-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="6e76d-131">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-131">Member</span></span> |
| [<span data-ttu-id="6e76d-132">end</span><span class="sxs-lookup"><span data-stu-id="6e76d-132">end</span></span>](#end-datetime) | <span data-ttu-id="6e76d-133">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-133">Member</span></span> |
| [<span data-ttu-id="6e76d-134">from</span><span class="sxs-lookup"><span data-stu-id="6e76d-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="6e76d-135">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-135">Member</span></span> |
| [<span data-ttu-id="6e76d-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="6e76d-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="6e76d-137">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-137">Member</span></span> |
| [<span data-ttu-id="6e76d-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="6e76d-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="6e76d-139">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-139">Member</span></span> |
| [<span data-ttu-id="6e76d-140">itemId</span><span class="sxs-lookup"><span data-stu-id="6e76d-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="6e76d-141">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-141">Member</span></span> |
| [<span data-ttu-id="6e76d-142">itemType</span><span class="sxs-lookup"><span data-stu-id="6e76d-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="6e76d-143">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-143">Member</span></span> |
| [<span data-ttu-id="6e76d-144">location</span><span class="sxs-lookup"><span data-stu-id="6e76d-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="6e76d-145">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-145">Member</span></span> |
| [<span data-ttu-id="6e76d-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="6e76d-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="6e76d-147">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-147">Member</span></span> |
| [<span data-ttu-id="6e76d-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="6e76d-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="6e76d-149">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-149">Member</span></span> |
| [<span data-ttu-id="6e76d-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="6e76d-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6e76d-151">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-151">Member</span></span> |
| [<span data-ttu-id="6e76d-152">organizer</span><span class="sxs-lookup"><span data-stu-id="6e76d-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="6e76d-153">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-153">Member</span></span> |
| [<span data-ttu-id="6e76d-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="6e76d-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6e76d-155">Member</span><span class="sxs-lookup"><span data-stu-id="6e76d-155">Member</span></span> |
| [<span data-ttu-id="6e76d-156">sender</span><span class="sxs-lookup"><span data-stu-id="6e76d-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="6e76d-157">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-157">Member</span></span> |
| [<span data-ttu-id="6e76d-158">start</span><span class="sxs-lookup"><span data-stu-id="6e76d-158">start</span></span>](#start-datetime) | <span data-ttu-id="6e76d-159">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-159">Member</span></span> |
| [<span data-ttu-id="6e76d-160">subject</span><span class="sxs-lookup"><span data-stu-id="6e76d-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="6e76d-161">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-161">Member</span></span> |
| [<span data-ttu-id="6e76d-162">to</span><span class="sxs-lookup"><span data-stu-id="6e76d-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6e76d-163">Membre</span><span class="sxs-lookup"><span data-stu-id="6e76d-163">Member</span></span> |
| [<span data-ttu-id="6e76d-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="6e76d-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-165">Method</span></span> |
| [<span data-ttu-id="6e76d-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="6e76d-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-167">Method</span></span> |
| [<span data-ttu-id="6e76d-168">close</span><span class="sxs-lookup"><span data-stu-id="6e76d-168">close</span></span>](#close) | <span data-ttu-id="6e76d-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-169">Method</span></span> |
| [<span data-ttu-id="6e76d-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="6e76d-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="6e76d-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-171">Method</span></span> |
| [<span data-ttu-id="6e76d-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="6e76d-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="6e76d-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-173">Method</span></span> |
| [<span data-ttu-id="6e76d-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="6e76d-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="6e76d-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-175">Method</span></span> |
| [<span data-ttu-id="6e76d-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="6e76d-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="6e76d-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-177">Method</span></span> |
| [<span data-ttu-id="6e76d-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="6e76d-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="6e76d-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-179">Method</span></span> |
| [<span data-ttu-id="6e76d-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="6e76d-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="6e76d-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-181">Method</span></span> |
| [<span data-ttu-id="6e76d-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="6e76d-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="6e76d-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-183">Method</span></span> |
| [<span data-ttu-id="6e76d-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="6e76d-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-185">Method</span></span> |
| [<span data-ttu-id="6e76d-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="6e76d-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-187">Method</span></span> |
| [<span data-ttu-id="6e76d-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="6e76d-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-189">Method</span></span> |
| [<span data-ttu-id="6e76d-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="6e76d-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-191">Method</span></span> |
| [<span data-ttu-id="6e76d-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6e76d-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="6e76d-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="6e76d-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="6e76d-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-194">Example</span></span>

<span data-ttu-id="6e76d-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="6e76d-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="6e76d-196">Members</span><span class="sxs-lookup"><span data-stu-id="6e76d-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-13"></a><span data-ttu-id="6e76d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="6e76d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

<span data-ttu-id="6e76d-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="6e76d-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="6e76d-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="6e76d-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-202">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-202">Type</span></span>

*   <span data-ttu-id="6e76d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="6e76d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-204">Requirements</span></span>

|<span data-ttu-id="6e76d-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-205">Requirement</span></span>| <span data-ttu-id="6e76d-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-208">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-208">1.0</span></span>|
|[<span data-ttu-id="6e76d-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-210">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-213">Example</span></span>

<span data-ttu-id="6e76d-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="6e76d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="6e76d-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-217">Compose mode only.</span></span>

<span data-ttu-id="6e76d-218">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-219">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="6e76d-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="6e76d-220">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="6e76d-221">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="6e76d-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-222">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-222">Type</span></span>

*   [<span data-ttu-id="6e76d-223">Destinataires</span><span class="sxs-lookup"><span data-stu-id="6e76d-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-224">Requirements</span></span>

|<span data-ttu-id="6e76d-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-225">Requirement</span></span>| <span data-ttu-id="6e76d-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-228">1.1</span><span class="sxs-lookup"><span data-stu-id="6e76d-228">1.1</span></span>|
|[<span data-ttu-id="6e76d-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-230">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-232">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-233">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-13"></a><span data-ttu-id="6e76d-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-235">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-236">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-236">Type</span></span>

*   [<span data-ttu-id="6e76d-237">Body</span><span class="sxs-lookup"><span data-stu-id="6e76d-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-238">Requirements</span></span>

|<span data-ttu-id="6e76d-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-239">Requirement</span></span>| <span data-ttu-id="6e76d-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-242">1.1</span><span class="sxs-lookup"><span data-stu-id="6e76d-242">1.1</span></span>|
|[<span data-ttu-id="6e76d-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-244">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-247">Example</span></span>

<span data-ttu-id="6e76d-248">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="6e76d-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="6e76d-249">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="6e76d-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-251">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="6e76d-252">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-253">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-253">Read mode</span></span>

<span data-ttu-id="6e76d-254">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="6e76d-255">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-256">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-257">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-257">Compose mode</span></span>

<span data-ttu-id="6e76d-258">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="6e76d-259">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-260">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="6e76d-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="6e76d-261">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="6e76d-262">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="6e76d-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

<br>

---
---

##### <a name="type"></a><span data-ttu-id="6e76d-263">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-263">Type</span></span>

*   <span data-ttu-id="6e76d-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-265">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-265">Requirements</span></span>

|<span data-ttu-id="6e76d-266">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-266">Requirement</span></span>| <span data-ttu-id="6e76d-267">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-268">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-269">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-269">1.0</span></span>|
|[<span data-ttu-id="6e76d-270">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-271">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-272">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-273">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="6e76d-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="6e76d-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="6e76d-275">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="6e76d-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="6e76d-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="6e76d-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-280">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-280">Type</span></span>

*   <span data-ttu-id="6e76d-281">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-282">Requirements</span></span>

|<span data-ttu-id="6e76d-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-283">Requirement</span></span>| <span data-ttu-id="6e76d-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-286">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-286">1.0</span></span>|
|[<span data-ttu-id="6e76d-287">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-288">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-289">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-290">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-291">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="6e76d-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="6e76d-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="6e76d-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-295">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-295">Type</span></span>

*   <span data-ttu-id="6e76d-296">Date</span><span class="sxs-lookup"><span data-stu-id="6e76d-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-297">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-297">Requirements</span></span>

|<span data-ttu-id="6e76d-298">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-298">Requirement</span></span>| <span data-ttu-id="6e76d-299">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-300">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-301">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-301">1.0</span></span>|
|[<span data-ttu-id="6e76d-302">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-303">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-304">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-305">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-306">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="6e76d-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="6e76d-307">dateTimeModified: Date</span></span>

<span data-ttu-id="6e76d-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-310">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-311">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-311">Type</span></span>

*   <span data-ttu-id="6e76d-312">Date</span><span class="sxs-lookup"><span data-stu-id="6e76d-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-313">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-313">Requirements</span></span>

|<span data-ttu-id="6e76d-314">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-314">Requirement</span></span>| <span data-ttu-id="6e76d-315">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-316">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-317">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-317">1.0</span></span>|
|[<span data-ttu-id="6e76d-318">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-319">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-320">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-321">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-322">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="6e76d-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-324">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="6e76d-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-327">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-327">Read mode</span></span>

<span data-ttu-id="6e76d-328">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-329">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-329">Compose mode</span></span>

<span data-ttu-id="6e76d-330">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="6e76d-331">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="6e76d-332">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="6e76d-333">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-333">Type</span></span>

*   <span data-ttu-id="6e76d-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-335">Requirements</span></span>

|<span data-ttu-id="6e76d-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-336">Requirement</span></span>| <span data-ttu-id="6e76d-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-339">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-339">1.0</span></span>|
|[<span data-ttu-id="6e76d-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-341">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-343">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="6e76d-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-p114">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="6e76d-p115">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-349">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-350">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-350">Type</span></span>

*   [<span data-ttu-id="6e76d-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6e76d-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-352">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-352">Requirements</span></span>

|<span data-ttu-id="6e76d-353">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-353">Requirement</span></span>| <span data-ttu-id="6e76d-354">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-355">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-356">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-356">1.0</span></span>|
|[<span data-ttu-id="6e76d-357">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-358">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-359">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-360">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-361">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="6e76d-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="6e76d-362">internetMessageId: String</span></span>

<span data-ttu-id="6e76d-p116">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-365">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-365">Type</span></span>

*   <span data-ttu-id="6e76d-366">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-367">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-367">Requirements</span></span>

|<span data-ttu-id="6e76d-368">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-368">Requirement</span></span>| <span data-ttu-id="6e76d-369">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-370">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-371">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-371">1.0</span></span>|
|[<span data-ttu-id="6e76d-372">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-373">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-374">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-375">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-376">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="6e76d-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="6e76d-377">itemClass: String</span></span>

<span data-ttu-id="6e76d-p117">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="6e76d-p118">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="6e76d-382">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-382">Type</span></span> | <span data-ttu-id="6e76d-383">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-383">Description</span></span> | <span data-ttu-id="6e76d-384">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="6e76d-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="6e76d-385">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="6e76d-385">Appointment items</span></span> | <span data-ttu-id="6e76d-386">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="6e76d-387">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="6e76d-387">Message items</span></span> | <span data-ttu-id="6e76d-388">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="6e76d-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="6e76d-389">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-390">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-390">Type</span></span>

*   <span data-ttu-id="6e76d-391">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-392">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-392">Requirements</span></span>

|<span data-ttu-id="6e76d-393">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-393">Requirement</span></span>| <span data-ttu-id="6e76d-394">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-395">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-396">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-396">1.0</span></span>|
|[<span data-ttu-id="6e76d-397">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-398">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-400">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-401">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="6e76d-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="6e76d-402">(nullable) itemId: String</span></span>

<span data-ttu-id="6e76d-p119">Obtient l' [identificateur d’élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actuel. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-405">L’identificateur renvoyé par la `itemId` propriété est identique à l’identificateur d' [élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="6e76d-405">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="6e76d-406">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="6e76d-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="6e76d-407">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="6e76d-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="6e76d-408">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="6e76d-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="6e76d-p121">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-411">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-411">Type</span></span>

*   <span data-ttu-id="6e76d-412">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-413">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-413">Requirements</span></span>

|<span data-ttu-id="6e76d-414">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-414">Requirement</span></span>| <span data-ttu-id="6e76d-415">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-416">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-417">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-417">1.0</span></span>|
|[<span data-ttu-id="6e76d-418">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-419">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-420">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-421">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-422">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-422">Example</span></span>

<span data-ttu-id="6e76d-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-13"></a><span data-ttu-id="6e76d-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-426">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="6e76d-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="6e76d-427">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-428">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-428">Type</span></span>

*   [<span data-ttu-id="6e76d-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="6e76d-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-430">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-430">Requirements</span></span>

|<span data-ttu-id="6e76d-431">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-431">Requirement</span></span>| <span data-ttu-id="6e76d-432">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-433">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-434">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-434">1.0</span></span>|
|[<span data-ttu-id="6e76d-435">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-436">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-437">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-438">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-439">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-439">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-13"></a><span data-ttu-id="6e76d-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-441">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-442">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-442">Read mode</span></span>

<span data-ttu-id="6e76d-443">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-444">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-444">Compose mode</span></span>

<span data-ttu-id="6e76d-445">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6e76d-446">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-446">Type</span></span>

*   <span data-ttu-id="6e76d-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-448">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-448">Requirements</span></span>

|<span data-ttu-id="6e76d-449">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-449">Requirement</span></span>| <span data-ttu-id="6e76d-450">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-451">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-452">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-452">1.0</span></span>|
|[<span data-ttu-id="6e76d-453">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-454">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-455">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-456">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="6e76d-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="6e76d-457">normalizedSubject: String</span></span>

<span data-ttu-id="6e76d-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="6e76d-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="6e76d-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-462">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-462">Type</span></span>

*   <span data-ttu-id="6e76d-463">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-464">Requirements</span></span>

|<span data-ttu-id="6e76d-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-465">Requirement</span></span>| <span data-ttu-id="6e76d-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-468">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-468">1.0</span></span>|
|[<span data-ttu-id="6e76d-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-470">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-472">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-13"></a><span data-ttu-id="6e76d-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-475">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-476">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-476">Type</span></span>

*   [<span data-ttu-id="6e76d-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="6e76d-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-478">Requirements</span></span>

|<span data-ttu-id="6e76d-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-479">Requirement</span></span>| <span data-ttu-id="6e76d-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-482">1.3</span><span class="sxs-lookup"><span data-stu-id="6e76d-482">1.3</span></span>|
|[<span data-ttu-id="6e76d-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-484">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-486">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-487">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="6e76d-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-489">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="6e76d-490">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-491">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-491">Read mode</span></span>

<span data-ttu-id="6e76d-492">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="6e76d-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="6e76d-493">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-494">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-495">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-495">Compose mode</span></span>

<span data-ttu-id="6e76d-496">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="6e76d-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="6e76d-497">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-498">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="6e76d-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="6e76d-499">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="6e76d-500">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="6e76d-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6e76d-501">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-501">Type</span></span>

*   <span data-ttu-id="6e76d-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-503">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-503">Requirements</span></span>

|<span data-ttu-id="6e76d-504">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-504">Requirement</span></span>| <span data-ttu-id="6e76d-505">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-506">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-507">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-507">1.0</span></span>|
|[<span data-ttu-id="6e76d-508">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-509">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-510">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-511">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="6e76d-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-p128">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-515">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-515">Type</span></span>

*   [<span data-ttu-id="6e76d-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6e76d-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-517">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-517">Requirements</span></span>

|<span data-ttu-id="6e76d-518">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-518">Requirement</span></span>| <span data-ttu-id="6e76d-519">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-520">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-521">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-521">1.0</span></span>|
|[<span data-ttu-id="6e76d-522">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-523">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-524">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-525">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-526">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="6e76d-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-528">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="6e76d-529">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-530">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-530">Read mode</span></span>

<span data-ttu-id="6e76d-531">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="6e76d-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="6e76d-532">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-533">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-534">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-534">Compose mode</span></span>

<span data-ttu-id="6e76d-535">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="6e76d-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="6e76d-536">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-537">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="6e76d-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="6e76d-538">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="6e76d-539">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="6e76d-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="6e76d-540">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-540">Type</span></span>

*   <span data-ttu-id="6e76d-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-542">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-542">Requirements</span></span>

|<span data-ttu-id="6e76d-543">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-543">Requirement</span></span>| <span data-ttu-id="6e76d-544">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-545">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-546">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-546">1.0</span></span>|
|[<span data-ttu-id="6e76d-547">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-548">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-549">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-550">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="6e76d-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-p132">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="6e76d-p133">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-556">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6e76d-557">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-557">Type</span></span>

*   [<span data-ttu-id="6e76d-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6e76d-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="6e76d-559">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-559">Requirements</span></span>

|<span data-ttu-id="6e76d-560">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-560">Requirement</span></span>| <span data-ttu-id="6e76d-561">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-562">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-563">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-563">1.0</span></span>|
|[<span data-ttu-id="6e76d-564">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-565">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-566">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-567">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-568">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="6e76d-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-570">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="6e76d-p134">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-573">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-573">Read mode</span></span>

<span data-ttu-id="6e76d-574">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-575">Compose mode</span></span>

<span data-ttu-id="6e76d-576">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="6e76d-577">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="6e76d-578">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="6e76d-579">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-579">Type</span></span>

*   <span data-ttu-id="6e76d-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-581">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-581">Requirements</span></span>

|<span data-ttu-id="6e76d-582">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-582">Requirement</span></span>| <span data-ttu-id="6e76d-583">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-584">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-585">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-585">1.0</span></span>|
|[<span data-ttu-id="6e76d-586">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-587">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-588">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-589">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-13"></a><span data-ttu-id="6e76d-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-591">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="6e76d-592">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="6e76d-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-593">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-593">Read mode</span></span>

<span data-ttu-id="6e76d-p135">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-596">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-596">Compose mode</span></span>

<span data-ttu-id="6e76d-597">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="6e76d-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="6e76d-598">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-598">Type</span></span>

*   <span data-ttu-id="6e76d-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-600">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-600">Requirements</span></span>

|<span data-ttu-id="6e76d-601">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-601">Requirement</span></span>| <span data-ttu-id="6e76d-602">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-603">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-604">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-604">1.0</span></span>|
|[<span data-ttu-id="6e76d-605">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-606">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-607">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-608">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="6e76d-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="6e76d-610">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="6e76d-611">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6e76d-612">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-612">Read mode</span></span>

<span data-ttu-id="6e76d-613">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="6e76d-614">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-615">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="6e76d-616">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-616">Compose mode</span></span>

<span data-ttu-id="6e76d-617">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="6e76d-618">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6e76d-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="6e76d-619">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="6e76d-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="6e76d-620">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="6e76d-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="6e76d-621">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="6e76d-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6e76d-622">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-622">Type</span></span>

*   <span data-ttu-id="6e76d-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-624">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-624">Requirements</span></span>

|<span data-ttu-id="6e76d-625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-625">Requirement</span></span>| <span data-ttu-id="6e76d-626">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-627">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-628">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-628">1.0</span></span>|
|[<span data-ttu-id="6e76d-629">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-630">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-631">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-632">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="6e76d-633">Méthodes</span><span class="sxs-lookup"><span data-stu-id="6e76d-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="6e76d-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6e76d-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6e76d-635">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="6e76d-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="6e76d-636">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="6e76d-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="6e76d-637">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="6e76d-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-638">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e76d-638">Parameters</span></span>

|<span data-ttu-id="6e76d-639">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-639">Name</span></span>| <span data-ttu-id="6e76d-640">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-640">Type</span></span>| <span data-ttu-id="6e76d-641">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-641">Attributes</span></span>| <span data-ttu-id="6e76d-642">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="6e76d-643">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-643">String</span></span>||<span data-ttu-id="6e76d-p139">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6e76d-646">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-646">String</span></span>||<span data-ttu-id="6e76d-p140">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6e76d-649">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-649">Object</span></span>| <span data-ttu-id="6e76d-650">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-650">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-651">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-651">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6e76d-652">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-652">Object</span></span>| <span data-ttu-id="6e76d-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-653">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-654">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-654">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6e76d-655">fonction</span><span class="sxs-lookup"><span data-stu-id="6e76d-655">function</span></span>| <span data-ttu-id="6e76d-656">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-656">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-657">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-657">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6e76d-658">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-658">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6e76d-659">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-659">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6e76d-660">Erreurs</span><span class="sxs-lookup"><span data-stu-id="6e76d-660">Errors</span></span>

| <span data-ttu-id="6e76d-661">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="6e76d-661">Error code</span></span> | <span data-ttu-id="6e76d-662">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-662">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="6e76d-663">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="6e76d-663">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="6e76d-664">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="6e76d-664">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6e76d-665">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-665">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6e76d-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-666">Requirements</span></span>

|<span data-ttu-id="6e76d-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-667">Requirement</span></span>| <span data-ttu-id="6e76d-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-670">1.1</span><span class="sxs-lookup"><span data-stu-id="6e76d-670">1.1</span></span>|
|[<span data-ttu-id="6e76d-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-672">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-672">ReadWriteItem</span></span>|
|[<span data-ttu-id="6e76d-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-674">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-674">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-675">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-675">Example</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="6e76d-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6e76d-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6e76d-677">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-677">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="6e76d-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="6e76d-681">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="6e76d-681">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="6e76d-682">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="6e76d-682">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-683">Parameters</span><span class="sxs-lookup"><span data-stu-id="6e76d-683">Parameters</span></span>

|<span data-ttu-id="6e76d-684">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-684">Name</span></span>| <span data-ttu-id="6e76d-685">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-685">Type</span></span>| <span data-ttu-id="6e76d-686">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-686">Attributes</span></span>| <span data-ttu-id="6e76d-687">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-687">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="6e76d-688">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-688">String</span></span>||<span data-ttu-id="6e76d-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6e76d-691">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-691">String</span></span>||<span data-ttu-id="6e76d-692">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="6e76d-692">The subject of the item to be attached.</span></span> <span data-ttu-id="6e76d-693">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-693">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6e76d-694">Object</span><span class="sxs-lookup"><span data-stu-id="6e76d-694">Object</span></span>| <span data-ttu-id="6e76d-695">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-695">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-696">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-696">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6e76d-697">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-697">Object</span></span>| <span data-ttu-id="6e76d-698">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-698">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-699">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-699">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6e76d-700">fonction</span><span class="sxs-lookup"><span data-stu-id="6e76d-700">function</span></span>| <span data-ttu-id="6e76d-701">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-701">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-702">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6e76d-703">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6e76d-704">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-704">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6e76d-705">Erreurs</span><span class="sxs-lookup"><span data-stu-id="6e76d-705">Errors</span></span>

| <span data-ttu-id="6e76d-706">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="6e76d-706">Error code</span></span> | <span data-ttu-id="6e76d-707">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-707">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6e76d-708">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-708">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6e76d-709">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-709">Requirements</span></span>

|<span data-ttu-id="6e76d-710">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-710">Requirement</span></span>| <span data-ttu-id="6e76d-711">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-712">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-713">1.1</span><span class="sxs-lookup"><span data-stu-id="6e76d-713">1.1</span></span>|
|[<span data-ttu-id="6e76d-714">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-714">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-715">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-715">ReadWriteItem</span></span>|
|[<span data-ttu-id="6e76d-716">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-716">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-717">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-717">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-718">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-718">Example</span></span>

<span data-ttu-id="6e76d-719">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-719">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="close"></a><span data-ttu-id="6e76d-720">close()</span><span class="sxs-lookup"><span data-stu-id="6e76d-720">close()</span></span>

<span data-ttu-id="6e76d-721">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="6e76d-721">Closes the current item that is being composed.</span></span>

<span data-ttu-id="6e76d-p144">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-724">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-724">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="6e76d-725">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="6e76d-725">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-726">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-726">Requirements</span></span>

|<span data-ttu-id="6e76d-727">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-727">Requirement</span></span>| <span data-ttu-id="6e76d-728">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-729">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-730">1.3</span><span class="sxs-lookup"><span data-stu-id="6e76d-730">1.3</span></span>|
|[<span data-ttu-id="6e76d-731">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-732">Restreinte</span><span class="sxs-lookup"><span data-stu-id="6e76d-732">Restricted</span></span>|
|[<span data-ttu-id="6e76d-733">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-734">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-734">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="6e76d-735">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="6e76d-735">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="6e76d-736">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6e76d-736">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-737">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6e76d-738">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-738">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6e76d-739">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="6e76d-739">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="6e76d-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-743">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e76d-743">Parameters</span></span>

|<span data-ttu-id="6e76d-744">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-744">Name</span></span>| <span data-ttu-id="6e76d-745">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-745">Type</span></span>| <span data-ttu-id="6e76d-746">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-746">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6e76d-747">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6e76d-747">String &#124; Object</span></span>| |<span data-ttu-id="6e76d-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6e76d-750">**OU**</span><span class="sxs-lookup"><span data-stu-id="6e76d-750">**OR**</span></span><br/><span data-ttu-id="6e76d-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="6e76d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6e76d-753">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-753">String</span></span> | <span data-ttu-id="6e76d-754">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-754">&lt;optional&gt;</span></span> | <span data-ttu-id="6e76d-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6e76d-757">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-757">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6e76d-758">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-758">&lt;optional&gt;</span></span> | <span data-ttu-id="6e76d-759">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-759">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6e76d-760">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-760">String</span></span> | | <span data-ttu-id="6e76d-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6e76d-763">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-763">String</span></span> | | <span data-ttu-id="6e76d-764">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-764">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6e76d-765">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-765">String</span></span> | | <span data-ttu-id="6e76d-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6e76d-768">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-768">String</span></span> | | <span data-ttu-id="6e76d-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6e76d-772">function</span><span class="sxs-lookup"><span data-stu-id="6e76d-772">function</span></span> | <span data-ttu-id="6e76d-773">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-773">&lt;optional&gt;</span></span> | <span data-ttu-id="6e76d-774">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-774">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6e76d-775">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-775">Requirements</span></span>

|<span data-ttu-id="6e76d-776">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-776">Requirement</span></span>| <span data-ttu-id="6e76d-777">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-777">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-778">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-778">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-779">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-779">1.0</span></span>|
|[<span data-ttu-id="6e76d-780">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-780">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-781">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-781">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-782">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-782">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-783">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-783">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6e76d-784">Exemples</span><span class="sxs-lookup"><span data-stu-id="6e76d-784">Examples</span></span>

<span data-ttu-id="6e76d-785">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-785">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="6e76d-786">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="6e76d-786">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="6e76d-787">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="6e76d-787">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6e76d-788">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="6e76d-788">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6e76d-789">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-789">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6e76d-790">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-790">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="6e76d-791">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="6e76d-791">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="6e76d-792">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6e76d-792">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-793">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-793">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6e76d-794">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-794">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6e76d-795">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="6e76d-795">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="6e76d-p152">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-799">Parameters</span><span class="sxs-lookup"><span data-stu-id="6e76d-799">Parameters</span></span>

|<span data-ttu-id="6e76d-800">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-800">Name</span></span>| <span data-ttu-id="6e76d-801">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-801">Type</span></span>| <span data-ttu-id="6e76d-802">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-802">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6e76d-803">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6e76d-803">String &#124; Object</span></span>| | <span data-ttu-id="6e76d-p153">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6e76d-806">**OU**</span><span class="sxs-lookup"><span data-stu-id="6e76d-806">**OR**</span></span><br/><span data-ttu-id="6e76d-p154">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="6e76d-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6e76d-809">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-809">String</span></span> | <span data-ttu-id="6e76d-810">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-810">&lt;optional&gt;</span></span> | <span data-ttu-id="6e76d-p155">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6e76d-813">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-813">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6e76d-814">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-814">&lt;optional&gt;</span></span> | <span data-ttu-id="6e76d-815">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-815">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6e76d-816">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-816">String</span></span> | | <span data-ttu-id="6e76d-p156">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6e76d-819">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-819">String</span></span> | | <span data-ttu-id="6e76d-820">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-820">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6e76d-821">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-821">String</span></span> | | <span data-ttu-id="6e76d-p157">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6e76d-824">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-824">String</span></span> | | <span data-ttu-id="6e76d-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6e76d-828">function</span><span class="sxs-lookup"><span data-stu-id="6e76d-828">function</span></span> | <span data-ttu-id="6e76d-829">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-829">&lt;optional&gt;</span></span> | <span data-ttu-id="6e76d-830">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-830">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6e76d-831">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-831">Requirements</span></span>

|<span data-ttu-id="6e76d-832">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-832">Requirement</span></span>| <span data-ttu-id="6e76d-833">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-834">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-835">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-835">1.0</span></span>|
|[<span data-ttu-id="6e76d-836">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-837">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-837">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-838">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-839">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-839">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6e76d-840">Exemples</span><span class="sxs-lookup"><span data-stu-id="6e76d-840">Examples</span></span>

<span data-ttu-id="6e76d-841">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-841">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="6e76d-842">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="6e76d-842">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="6e76d-843">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="6e76d-843">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6e76d-844">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="6e76d-844">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6e76d-845">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-845">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6e76d-846">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-846">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-13"></a><span data-ttu-id="6e76d-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="6e76d-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="6e76d-848">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6e76d-848">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-849">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-849">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-850">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-850">Requirements</span></span>

|<span data-ttu-id="6e76d-851">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-851">Requirement</span></span>| <span data-ttu-id="6e76d-852">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-853">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-854">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-854">1.0</span></span>|
|[<span data-ttu-id="6e76d-855">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-855">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-856">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-856">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-857">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-857">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-858">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-858">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6e76d-859">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6e76d-859">Returns:</span></span>

<span data-ttu-id="6e76d-860">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="6e76d-860">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span></span>

##### <a name="example"></a><span data-ttu-id="6e76d-861">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-861">Example</span></span>

<span data-ttu-id="6e76d-862">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-862">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="6e76d-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="6e76d-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="6e76d-864">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6e76d-864">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-865">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-866">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e76d-866">Parameters</span></span>

|<span data-ttu-id="6e76d-867">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-867">Name</span></span>| <span data-ttu-id="6e76d-868">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-868">Type</span></span>| <span data-ttu-id="6e76d-869">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-869">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="6e76d-870">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="6e76d-870">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.3)|<span data-ttu-id="6e76d-871">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="6e76d-871">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6e76d-872">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-872">Requirements</span></span>

|<span data-ttu-id="6e76d-873">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-873">Requirement</span></span>| <span data-ttu-id="6e76d-874">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-874">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-875">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-875">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-876">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-876">1.0</span></span>|
|[<span data-ttu-id="6e76d-877">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-877">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-878">Restreinte</span><span class="sxs-lookup"><span data-stu-id="6e76d-878">Restricted</span></span>|
|[<span data-ttu-id="6e76d-879">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-879">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-880">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-880">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6e76d-881">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6e76d-881">Returns:</span></span>

<span data-ttu-id="6e76d-882">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="6e76d-882">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="6e76d-883">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="6e76d-883">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="6e76d-884">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-884">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="6e76d-885">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="6e76d-885">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="6e76d-886">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="6e76d-886">Value of `entityType`</span></span> | <span data-ttu-id="6e76d-887">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="6e76d-887">Type of objects in returned array</span></span> | <span data-ttu-id="6e76d-888">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="6e76d-888">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="6e76d-889">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-889">String</span></span> | <span data-ttu-id="6e76d-890">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6e76d-890">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="6e76d-891">Contact</span><span class="sxs-lookup"><span data-stu-id="6e76d-891">Contact</span></span> | <span data-ttu-id="6e76d-892">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6e76d-892">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="6e76d-893">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-893">String</span></span> | <span data-ttu-id="6e76d-894">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6e76d-894">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="6e76d-895">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="6e76d-895">MeetingSuggestion</span></span> | <span data-ttu-id="6e76d-896">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6e76d-896">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="6e76d-897">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="6e76d-897">PhoneNumber</span></span> | <span data-ttu-id="6e76d-898">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6e76d-898">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="6e76d-899">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="6e76d-899">TaskSuggestion</span></span> | <span data-ttu-id="6e76d-900">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6e76d-900">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="6e76d-901">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-901">String</span></span> | <span data-ttu-id="6e76d-902">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6e76d-902">**Restricted**</span></span> |

<span data-ttu-id="6e76d-903">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="6e76d-903">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

##### <a name="example"></a><span data-ttu-id="6e76d-904">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-904">Example</span></span>

<span data-ttu-id="6e76d-905">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6e76d-905">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="6e76d-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="6e76d-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="6e76d-907">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6e76d-907">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-908">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-908">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6e76d-909">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="6e76d-909">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-910">Parameters</span><span class="sxs-lookup"><span data-stu-id="6e76d-910">Parameters</span></span>

|<span data-ttu-id="6e76d-911">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-911">Name</span></span>| <span data-ttu-id="6e76d-912">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-912">Type</span></span>| <span data-ttu-id="6e76d-913">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-913">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6e76d-914">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-914">String</span></span>|<span data-ttu-id="6e76d-915">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="6e76d-915">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6e76d-916">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-916">Requirements</span></span>

|<span data-ttu-id="6e76d-917">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-917">Requirement</span></span>| <span data-ttu-id="6e76d-918">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-918">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-919">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-919">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-920">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-920">1.0</span></span>|
|[<span data-ttu-id="6e76d-921">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-921">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-922">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-922">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-923">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-923">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-924">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-924">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6e76d-925">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6e76d-925">Returns:</span></span>

<span data-ttu-id="6e76d-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="6e76d-928">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="6e76d-928">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="6e76d-929">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6e76d-929">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="6e76d-930">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6e76d-930">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-931">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6e76d-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6e76d-935">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="6e76d-935">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6e76d-936">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-936">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="6e76d-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e76d-940">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-940">Requirements</span></span>

|<span data-ttu-id="6e76d-941">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-941">Requirement</span></span>| <span data-ttu-id="6e76d-942">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-943">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-944">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-944">1.0</span></span>|
|[<span data-ttu-id="6e76d-945">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-946">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-947">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-948">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-948">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6e76d-949">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6e76d-949">Returns:</span></span>

<span data-ttu-id="6e76d-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="6e76d-952">Type : Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-952">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="6e76d-953">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-953">Example</span></span>

<span data-ttu-id="6e76d-954">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="6e76d-954">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="6e76d-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="6e76d-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="6e76d-956">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6e76d-956">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-957">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6e76d-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6e76d-958">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="6e76d-958">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="6e76d-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-961">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e76d-961">Parameters</span></span>

|<span data-ttu-id="6e76d-962">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-962">Name</span></span>| <span data-ttu-id="6e76d-963">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-963">Type</span></span>| <span data-ttu-id="6e76d-964">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-964">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6e76d-965">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6e76d-965">String</span></span>|<span data-ttu-id="6e76d-966">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="6e76d-966">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6e76d-967">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-967">Requirements</span></span>

|<span data-ttu-id="6e76d-968">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-968">Requirement</span></span>| <span data-ttu-id="6e76d-969">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-969">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-970">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-970">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-971">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-971">1.0</span></span>|
|[<span data-ttu-id="6e76d-972">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-972">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-973">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-973">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-974">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-974">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-975">Lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-975">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6e76d-976">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6e76d-976">Returns:</span></span>

<span data-ttu-id="6e76d-977">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6e76d-977">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="6e76d-978">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="6e76d-978">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="6e76d-979">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-979">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="6e76d-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="6e76d-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="6e76d-981">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="6e76d-981">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="6e76d-982">S’il n’y a aucune sélection, mais que le curseur se trouve dans le corps ou l’objet, la méthode renvoie une chaîne vide pour les données sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="6e76d-982">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="6e76d-983">Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-983">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-984">Dans Outlook sur le Web, la méthode renvoie la chaîne « NULL » si aucun texte n’est sélectionné, mais que le curseur se trouve dans le corps.</span><span class="sxs-lookup"><span data-stu-id="6e76d-984">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="6e76d-985">Pour vérifier cette situation, reportez-vous à l’exemple plus loin dans cette section.</span><span class="sxs-lookup"><span data-stu-id="6e76d-985">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-986">Parameters</span><span class="sxs-lookup"><span data-stu-id="6e76d-986">Parameters</span></span>

|<span data-ttu-id="6e76d-987">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-987">Name</span></span>| <span data-ttu-id="6e76d-988">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-988">Type</span></span>| <span data-ttu-id="6e76d-989">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-989">Attributes</span></span>| <span data-ttu-id="6e76d-990">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-990">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="6e76d-991">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6e76d-991">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="6e76d-p167">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p167">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="6e76d-995">Object</span><span class="sxs-lookup"><span data-stu-id="6e76d-995">Object</span></span>| <span data-ttu-id="6e76d-996">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-996">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-997">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-997">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6e76d-998">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-998">Object</span></span>| <span data-ttu-id="6e76d-999">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-999">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1000">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1000">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6e76d-1001">fonction</span><span class="sxs-lookup"><span data-stu-id="6e76d-1001">function</span></span>||<span data-ttu-id="6e76d-1002">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-1002">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6e76d-1003">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1003">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="6e76d-1004">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1004">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6e76d-1005">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-1005">Requirements</span></span>

|<span data-ttu-id="6e76d-1006">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-1006">Requirement</span></span>| <span data-ttu-id="6e76d-1007">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-1007">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-1008">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1008">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-1009">1.2</span><span class="sxs-lookup"><span data-stu-id="6e76d-1009">1.2</span></span>|
|[<span data-ttu-id="6e76d-1010">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-1010">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-1011">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-1011">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-1012">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-1012">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-1013">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-1013">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="6e76d-1014">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6e76d-1014">Returns:</span></span>

<span data-ttu-id="6e76d-1015">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1015">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="6e76d-1016">Type : String</span><span class="sxs-lookup"><span data-stu-id="6e76d-1016">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="6e76d-1017">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-1017">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  // Handle where Outlook on the web erroneously returns "null" instead of empty string.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookWebApp'
      && asyncResult.value.endPosition === asyncResult.value.startPosition) {
    text = "";
  }

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="6e76d-1018">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6e76d-1018">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="6e76d-1019">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1019">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="6e76d-p169">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p169">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-1023">Parameters</span><span class="sxs-lookup"><span data-stu-id="6e76d-1023">Parameters</span></span>

|<span data-ttu-id="6e76d-1024">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-1024">Name</span></span>| <span data-ttu-id="6e76d-1025">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-1025">Type</span></span>| <span data-ttu-id="6e76d-1026">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-1026">Attributes</span></span>| <span data-ttu-id="6e76d-1027">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-1027">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6e76d-1028">function</span><span class="sxs-lookup"><span data-stu-id="6e76d-1028">function</span></span>||<span data-ttu-id="6e76d-1029">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6e76d-1030">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1030">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="6e76d-1031">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1031">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="6e76d-1032">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1032">Object</span></span>| <span data-ttu-id="6e76d-1033">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1034">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1034">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="6e76d-1035">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1035">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6e76d-1036">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-1036">Requirements</span></span>

|<span data-ttu-id="6e76d-1037">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-1037">Requirement</span></span>| <span data-ttu-id="6e76d-1038">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-1038">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-1039">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1039">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-1040">1.0</span><span class="sxs-lookup"><span data-stu-id="6e76d-1040">1.0</span></span>|
|[<span data-ttu-id="6e76d-1041">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-1041">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-1042">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-1042">ReadItem</span></span>|
|[<span data-ttu-id="6e76d-1043">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-1043">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-1044">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6e76d-1044">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-1045">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-1045">Example</span></span>

<span data-ttu-id="6e76d-p172">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p172">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="6e76d-1049">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6e76d-1049">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="6e76d-1050">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1050">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="6e76d-1051">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1051">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="6e76d-1052">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1052">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="6e76d-1053">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1053">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="6e76d-1054">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1054">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-1055">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1055">Parameters</span></span>

|<span data-ttu-id="6e76d-1056">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-1056">Name</span></span>| <span data-ttu-id="6e76d-1057">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-1057">Type</span></span>| <span data-ttu-id="6e76d-1058">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-1058">Attributes</span></span>| <span data-ttu-id="6e76d-1059">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-1059">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="6e76d-1060">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-1060">String</span></span>||<span data-ttu-id="6e76d-1061">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1061">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="6e76d-1062">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1062">Object</span></span>| <span data-ttu-id="6e76d-1063">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1064">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6e76d-1065">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1065">Object</span></span>| <span data-ttu-id="6e76d-1066">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1067">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6e76d-1068">fonction</span><span class="sxs-lookup"><span data-stu-id="6e76d-1068">function</span></span>| <span data-ttu-id="6e76d-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1070">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6e76d-1071">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1071">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6e76d-1072">Erreurs</span><span class="sxs-lookup"><span data-stu-id="6e76d-1072">Errors</span></span>

| <span data-ttu-id="6e76d-1073">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="6e76d-1073">Error code</span></span> | <span data-ttu-id="6e76d-1074">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-1074">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="6e76d-1075">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1075">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6e76d-1076">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-1076">Requirements</span></span>

|<span data-ttu-id="6e76d-1077">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-1077">Requirement</span></span>| <span data-ttu-id="6e76d-1078">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-1078">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-1079">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1079">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-1080">1.1</span><span class="sxs-lookup"><span data-stu-id="6e76d-1080">1.1</span></span>|
|[<span data-ttu-id="6e76d-1081">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-1081">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-1082">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-1082">ReadWriteItem</span></span>|
|[<span data-ttu-id="6e76d-1083">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-1083">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-1084">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-1084">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-1085">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-1085">Example</span></span>

<span data-ttu-id="6e76d-1086">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="6e76d-1086">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="6e76d-1087">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="6e76d-1087">saveAsync([options], callback)</span></span>

<span data-ttu-id="6e76d-1088">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1088">Asynchronously saves an item.</span></span>

<span data-ttu-id="6e76d-1089">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1089">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="6e76d-1090">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1090">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="6e76d-1091">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1091">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-1092">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1092">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="6e76d-1093">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1093">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="6e76d-p176">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="6e76d-1097">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="6e76d-1097">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="6e76d-1098">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1098">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="6e76d-1099">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1099">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="6e76d-1100">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="6e76d-1100">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="6e76d-1101">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1101">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-1102">Parameters</span><span class="sxs-lookup"><span data-stu-id="6e76d-1102">Parameters</span></span>

|<span data-ttu-id="6e76d-1103">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-1103">Name</span></span>| <span data-ttu-id="6e76d-1104">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-1104">Type</span></span>| <span data-ttu-id="6e76d-1105">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-1105">Attributes</span></span>| <span data-ttu-id="6e76d-1106">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-1106">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="6e76d-1107">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1107">Object</span></span>| <span data-ttu-id="6e76d-1108">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1108">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1109">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1109">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6e76d-1110">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1110">Object</span></span>| <span data-ttu-id="6e76d-1111">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1111">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1112">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1112">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6e76d-1113">fonction</span><span class="sxs-lookup"><span data-stu-id="6e76d-1113">function</span></span>||<span data-ttu-id="6e76d-1114">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-1114">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6e76d-1115">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1115">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6e76d-1116">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-1116">Requirements</span></span>

|<span data-ttu-id="6e76d-1117">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-1117">Requirement</span></span>| <span data-ttu-id="6e76d-1118">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-1119">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-1120">1.3</span><span class="sxs-lookup"><span data-stu-id="6e76d-1120">1.3</span></span>|
|[<span data-ttu-id="6e76d-1121">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-1122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-1122">ReadWriteItem</span></span>|
|[<span data-ttu-id="6e76d-1123">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-1124">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-1124">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="6e76d-1125">範例</span><span class="sxs-lookup"><span data-stu-id="6e76d-1125">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="6e76d-p178">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="6e76d-1128">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="6e76d-1128">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="6e76d-1129">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1129">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="6e76d-p179">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6e76d-1133">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1133">Parameters</span></span>

|<span data-ttu-id="6e76d-1134">Nom</span><span class="sxs-lookup"><span data-stu-id="6e76d-1134">Name</span></span>| <span data-ttu-id="6e76d-1135">Type</span><span class="sxs-lookup"><span data-stu-id="6e76d-1135">Type</span></span>| <span data-ttu-id="6e76d-1136">Attributs</span><span class="sxs-lookup"><span data-stu-id="6e76d-1136">Attributes</span></span>| <span data-ttu-id="6e76d-1137">Description</span><span class="sxs-lookup"><span data-stu-id="6e76d-1137">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6e76d-1138">String</span><span class="sxs-lookup"><span data-stu-id="6e76d-1138">String</span></span>||<span data-ttu-id="6e76d-p180">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="6e76d-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="6e76d-1142">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1142">Object</span></span>| <span data-ttu-id="6e76d-1143">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1143">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1144">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1144">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6e76d-1145">Objet</span><span class="sxs-lookup"><span data-stu-id="6e76d-1145">Object</span></span>| <span data-ttu-id="6e76d-1146">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1146">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1147">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1147">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="6e76d-1148">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6e76d-1148">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="6e76d-1149">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6e76d-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="6e76d-1150">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1150">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="6e76d-1151">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1151">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="6e76d-1152">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1152">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="6e76d-1153">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1153">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="6e76d-1154">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="6e76d-1154">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="6e76d-1155">fonction</span><span class="sxs-lookup"><span data-stu-id="6e76d-1155">function</span></span>||<span data-ttu-id="6e76d-1156">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6e76d-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6e76d-1157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6e76d-1157">Requirements</span></span>

|<span data-ttu-id="6e76d-1158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6e76d-1158">Requirement</span></span>| <span data-ttu-id="6e76d-1159">Valeur</span><span class="sxs-lookup"><span data-stu-id="6e76d-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e76d-1160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6e76d-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e76d-1161">1.2</span><span class="sxs-lookup"><span data-stu-id="6e76d-1161">1.2</span></span>|
|[<span data-ttu-id="6e76d-1162">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6e76d-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e76d-1163">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6e76d-1163">ReadWriteItem</span></span>|
|[<span data-ttu-id="6e76d-1164">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6e76d-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e76d-1165">Composition</span><span class="sxs-lookup"><span data-stu-id="6e76d-1165">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6e76d-1166">Exemple</span><span class="sxs-lookup"><span data-stu-id="6e76d-1166">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
