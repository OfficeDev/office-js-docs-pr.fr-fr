---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,6
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 4aa9b5ae086b9879842a6f1cdd7125b74aa0c54d
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066143"
---
# <a name="item"></a><span data-ttu-id="557e1-102">élément</span><span class="sxs-lookup"><span data-stu-id="557e1-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="557e1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="557e1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="557e1-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="557e1-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-106">Requirements</span></span>

|<span data-ttu-id="557e1-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-107">Requirement</span></span>| <span data-ttu-id="557e1-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-110">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-110">1.0</span></span>|
|[<span data-ttu-id="557e1-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="557e1-112">Restricted</span></span>|
|[<span data-ttu-id="557e1-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="557e1-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="557e1-115">Members and methods</span></span>

| <span data-ttu-id="557e1-116">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-116">Member</span></span> | <span data-ttu-id="557e1-117">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="557e1-118">attachments</span><span class="sxs-lookup"><span data-stu-id="557e1-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="557e1-119">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-119">Member</span></span> |
| [<span data-ttu-id="557e1-120">bcc</span><span class="sxs-lookup"><span data-stu-id="557e1-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="557e1-121">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-121">Member</span></span> |
| [<span data-ttu-id="557e1-122">body</span><span class="sxs-lookup"><span data-stu-id="557e1-122">body</span></span>](#body-body) | <span data-ttu-id="557e1-123">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-123">Member</span></span> |
| [<span data-ttu-id="557e1-124">cc</span><span class="sxs-lookup"><span data-stu-id="557e1-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="557e1-125">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-125">Member</span></span> |
| [<span data-ttu-id="557e1-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="557e1-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="557e1-127">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-127">Member</span></span> |
| [<span data-ttu-id="557e1-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="557e1-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="557e1-129">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-129">Member</span></span> |
| [<span data-ttu-id="557e1-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="557e1-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="557e1-131">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-131">Member</span></span> |
| [<span data-ttu-id="557e1-132">end</span><span class="sxs-lookup"><span data-stu-id="557e1-132">end</span></span>](#end-datetime) | <span data-ttu-id="557e1-133">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-133">Member</span></span> |
| [<span data-ttu-id="557e1-134">from</span><span class="sxs-lookup"><span data-stu-id="557e1-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="557e1-135">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-135">Member</span></span> |
| [<span data-ttu-id="557e1-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="557e1-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="557e1-137">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-137">Member</span></span> |
| [<span data-ttu-id="557e1-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="557e1-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="557e1-139">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-139">Member</span></span> |
| [<span data-ttu-id="557e1-140">itemId</span><span class="sxs-lookup"><span data-stu-id="557e1-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="557e1-141">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-141">Member</span></span> |
| [<span data-ttu-id="557e1-142">itemType</span><span class="sxs-lookup"><span data-stu-id="557e1-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="557e1-143">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-143">Member</span></span> |
| [<span data-ttu-id="557e1-144">location</span><span class="sxs-lookup"><span data-stu-id="557e1-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="557e1-145">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-145">Member</span></span> |
| [<span data-ttu-id="557e1-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="557e1-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="557e1-147">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-147">Member</span></span> |
| [<span data-ttu-id="557e1-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="557e1-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="557e1-149">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-149">Member</span></span> |
| [<span data-ttu-id="557e1-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="557e1-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="557e1-151">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-151">Member</span></span> |
| [<span data-ttu-id="557e1-152">organizer</span><span class="sxs-lookup"><span data-stu-id="557e1-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="557e1-153">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-153">Member</span></span> |
| [<span data-ttu-id="557e1-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="557e1-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="557e1-155">Member</span><span class="sxs-lookup"><span data-stu-id="557e1-155">Member</span></span> |
| [<span data-ttu-id="557e1-156">sender</span><span class="sxs-lookup"><span data-stu-id="557e1-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="557e1-157">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-157">Member</span></span> |
| [<span data-ttu-id="557e1-158">start</span><span class="sxs-lookup"><span data-stu-id="557e1-158">start</span></span>](#start-datetime) | <span data-ttu-id="557e1-159">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-159">Member</span></span> |
| [<span data-ttu-id="557e1-160">subject</span><span class="sxs-lookup"><span data-stu-id="557e1-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="557e1-161">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-161">Member</span></span> |
| [<span data-ttu-id="557e1-162">to</span><span class="sxs-lookup"><span data-stu-id="557e1-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="557e1-163">Membre</span><span class="sxs-lookup"><span data-stu-id="557e1-163">Member</span></span> |
| [<span data-ttu-id="557e1-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="557e1-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-165">Method</span></span> |
| [<span data-ttu-id="557e1-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="557e1-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-167">Method</span></span> |
| [<span data-ttu-id="557e1-168">close</span><span class="sxs-lookup"><span data-stu-id="557e1-168">close</span></span>](#close) | <span data-ttu-id="557e1-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-169">Method</span></span> |
| [<span data-ttu-id="557e1-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="557e1-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="557e1-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-171">Method</span></span> |
| [<span data-ttu-id="557e1-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="557e1-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="557e1-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-173">Method</span></span> |
| [<span data-ttu-id="557e1-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="557e1-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="557e1-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-175">Method</span></span> |
| [<span data-ttu-id="557e1-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="557e1-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="557e1-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-177">Method</span></span> |
| [<span data-ttu-id="557e1-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="557e1-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="557e1-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-179">Method</span></span> |
| [<span data-ttu-id="557e1-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="557e1-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="557e1-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-181">Method</span></span> |
| [<span data-ttu-id="557e1-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="557e1-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="557e1-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-183">Method</span></span> |
| [<span data-ttu-id="557e1-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="557e1-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-185">Method</span></span> |
| [<span data-ttu-id="557e1-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="557e1-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="557e1-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-187">Method</span></span> |
| [<span data-ttu-id="557e1-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="557e1-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="557e1-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-189">Method</span></span> |
| [<span data-ttu-id="557e1-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="557e1-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-191">Method</span></span> |
| [<span data-ttu-id="557e1-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="557e1-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-193">Method</span></span> |
| [<span data-ttu-id="557e1-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="557e1-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-195">Method</span></span> |
| [<span data-ttu-id="557e1-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="557e1-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="557e1-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="557e1-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="557e1-198">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-198">Example</span></span>

<span data-ttu-id="557e1-199">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="557e1-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="557e1-200">Members</span><span class="sxs-lookup"><span data-stu-id="557e1-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="557e1-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="557e1-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="557e1-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-204">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="557e1-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="557e1-205">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="557e1-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-206">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-206">Type</span></span>

*   <span data-ttu-id="557e1-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="557e1-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-208">Requirements</span></span>

|<span data-ttu-id="557e1-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-209">Requirement</span></span>| <span data-ttu-id="557e1-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-212">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-212">1.0</span></span>|
|[<span data-ttu-id="557e1-213">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-214">ReadItem</span></span>|
|[<span data-ttu-id="557e1-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-216">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-217">Example</span></span>

<span data-ttu-id="557e1-218">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="557e1-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-220">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="557e1-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="557e1-221">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-221">Compose mode only.</span></span>

<span data-ttu-id="557e1-222">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-223">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="557e1-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="557e1-224">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="557e1-225">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="557e1-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-226">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-226">Type</span></span>

*   [<span data-ttu-id="557e1-227">Destinataires</span><span class="sxs-lookup"><span data-stu-id="557e1-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="557e1-228">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-228">Requirements</span></span>

|<span data-ttu-id="557e1-229">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-229">Requirement</span></span>| <span data-ttu-id="557e1-230">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-231">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-232">1.1</span><span class="sxs-lookup"><span data-stu-id="557e1-232">1.1</span></span>|
|[<span data-ttu-id="557e1-233">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-234">ReadItem</span></span>|
|[<span data-ttu-id="557e1-235">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-236">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-237">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="557e1-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-239">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-240">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-240">Type</span></span>

*   [<span data-ttu-id="557e1-241">Body</span><span class="sxs-lookup"><span data-stu-id="557e1-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="557e1-242">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-242">Requirements</span></span>

|<span data-ttu-id="557e1-243">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-243">Requirement</span></span>| <span data-ttu-id="557e1-244">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-245">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-246">1.1</span><span class="sxs-lookup"><span data-stu-id="557e1-246">1.1</span></span>|
|[<span data-ttu-id="557e1-247">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-248">ReadItem</span></span>|
|[<span data-ttu-id="557e1-249">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-250">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-251">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-251">Example</span></span>

<span data-ttu-id="557e1-252">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="557e1-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="557e1-253">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="557e1-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-255">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="557e1-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="557e1-256">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-257">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-257">Read mode</span></span>

<span data-ttu-id="557e1-258">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="557e1-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="557e1-259">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-260">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-261">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-261">Compose mode</span></span>

<span data-ttu-id="557e1-262">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="557e1-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="557e1-263">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-264">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="557e1-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="557e1-265">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="557e1-266">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="557e1-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="557e1-267">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-267">Type</span></span>

*   <span data-ttu-id="557e1-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-269">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-269">Requirements</span></span>

|<span data-ttu-id="557e1-270">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-270">Requirement</span></span>| <span data-ttu-id="557e1-271">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-272">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-273">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-273">1.0</span></span>|
|[<span data-ttu-id="557e1-274">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-275">ReadItem</span></span>|
|[<span data-ttu-id="557e1-276">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-277">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="557e1-278">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="557e1-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="557e1-279">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="557e1-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="557e1-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="557e1-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="557e1-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="557e1-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-284">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-284">Type</span></span>

*   <span data-ttu-id="557e1-285">String</span><span class="sxs-lookup"><span data-stu-id="557e1-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-286">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-286">Requirements</span></span>

|<span data-ttu-id="557e1-287">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-287">Requirement</span></span>| <span data-ttu-id="557e1-288">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-289">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-290">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-290">1.0</span></span>|
|[<span data-ttu-id="557e1-291">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-292">ReadItem</span></span>|
|[<span data-ttu-id="557e1-293">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-294">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-295">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="557e1-296">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="557e1-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="557e1-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-299">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-299">Type</span></span>

*   <span data-ttu-id="557e1-300">Date</span><span class="sxs-lookup"><span data-stu-id="557e1-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-301">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-301">Requirements</span></span>

|<span data-ttu-id="557e1-302">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-302">Requirement</span></span>| <span data-ttu-id="557e1-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-304">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-305">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-305">1.0</span></span>|
|[<span data-ttu-id="557e1-306">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-307">ReadItem</span></span>|
|[<span data-ttu-id="557e1-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="557e1-311">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="557e1-311">dateTimeModified: Date</span></span>

<span data-ttu-id="557e1-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-314">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-315">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-315">Type</span></span>

*   <span data-ttu-id="557e1-316">Date</span><span class="sxs-lookup"><span data-stu-id="557e1-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-317">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-317">Requirements</span></span>

|<span data-ttu-id="557e1-318">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-318">Requirement</span></span>| <span data-ttu-id="557e1-319">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-320">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-321">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-321">1.0</span></span>|
|[<span data-ttu-id="557e1-322">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-323">ReadItem</span></span>|
|[<span data-ttu-id="557e1-324">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-325">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-326">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="557e1-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-328">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="557e1-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="557e1-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-331">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-331">Read mode</span></span>

<span data-ttu-id="557e1-332">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="557e1-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-333">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-333">Compose mode</span></span>

<span data-ttu-id="557e1-334">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="557e1-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="557e1-335">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="557e1-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="557e1-336">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="557e1-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="557e1-337">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-337">Type</span></span>

*   <span data-ttu-id="557e1-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-339">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-339">Requirements</span></span>

|<span data-ttu-id="557e1-340">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-340">Requirement</span></span>| <span data-ttu-id="557e1-341">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-342">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-343">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-343">1.0</span></span>|
|[<span data-ttu-id="557e1-344">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-345">ReadItem</span></span>|
|[<span data-ttu-id="557e1-346">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-347">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="557e1-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-p114">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="557e1-p115">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="557e1-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-353">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="557e1-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-354">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-354">Type</span></span>

*   [<span data-ttu-id="557e1-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="557e1-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="557e1-356">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="557e1-357">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-357">Requirements</span></span>

|<span data-ttu-id="557e1-358">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-358">Requirement</span></span>| <span data-ttu-id="557e1-359">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-360">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-361">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-361">1.0</span></span>|
|[<span data-ttu-id="557e1-362">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-363">ReadItem</span></span>|
|[<span data-ttu-id="557e1-364">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-365">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="557e1-366">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="557e1-366">internetMessageId: String</span></span>

<span data-ttu-id="557e1-p116">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-369">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-369">Type</span></span>

*   <span data-ttu-id="557e1-370">String</span><span class="sxs-lookup"><span data-stu-id="557e1-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-371">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-371">Requirements</span></span>

|<span data-ttu-id="557e1-372">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-372">Requirement</span></span>| <span data-ttu-id="557e1-373">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-374">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-375">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-375">1.0</span></span>|
|[<span data-ttu-id="557e1-376">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-377">ReadItem</span></span>|
|[<span data-ttu-id="557e1-378">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-379">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-380">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="557e1-381">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="557e1-381">itemClass: String</span></span>

<span data-ttu-id="557e1-p117">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="557e1-p118">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="557e1-386">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-386">Type</span></span> | <span data-ttu-id="557e1-387">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-387">Description</span></span> | <span data-ttu-id="557e1-388">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="557e1-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="557e1-389">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="557e1-389">Appointment items</span></span> | <span data-ttu-id="557e1-390">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="557e1-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="557e1-391">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="557e1-391">Message items</span></span> | <span data-ttu-id="557e1-392">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="557e1-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="557e1-393">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="557e1-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-394">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-394">Type</span></span>

*   <span data-ttu-id="557e1-395">String</span><span class="sxs-lookup"><span data-stu-id="557e1-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-396">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-396">Requirements</span></span>

|<span data-ttu-id="557e1-397">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-397">Requirement</span></span>| <span data-ttu-id="557e1-398">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-399">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-400">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-400">1.0</span></span>|
|[<span data-ttu-id="557e1-401">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-402">ReadItem</span></span>|
|[<span data-ttu-id="557e1-403">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-404">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-405">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="557e1-406">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="557e1-406">(nullable) itemId: String</span></span>

<span data-ttu-id="557e1-407">Obtient l' [identificateur d’élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="557e1-407">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="557e1-408">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-408">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-409">L’identificateur renvoyé par la `itemId` propriété est identique à l’identificateur d' [élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="557e1-409">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="557e1-410">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="557e1-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="557e1-411">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="557e1-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="557e1-412">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="557e1-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="557e1-p121">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-415">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-415">Type</span></span>

*   <span data-ttu-id="557e1-416">String</span><span class="sxs-lookup"><span data-stu-id="557e1-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-417">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-417">Requirements</span></span>

|<span data-ttu-id="557e1-418">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-418">Requirement</span></span>| <span data-ttu-id="557e1-419">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-420">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-421">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-421">1.0</span></span>|
|[<span data-ttu-id="557e1-422">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-423">ReadItem</span></span>|
|[<span data-ttu-id="557e1-424">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-425">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-426">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-426">Example</span></span>

<span data-ttu-id="557e1-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="557e1-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="557e1-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-430">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="557e1-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="557e1-431">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-432">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-432">Type</span></span>

*   [<span data-ttu-id="557e1-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="557e1-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="557e1-434">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-434">Requirements</span></span>

|<span data-ttu-id="557e1-435">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-435">Requirement</span></span>| <span data-ttu-id="557e1-436">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-437">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-438">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-438">1.0</span></span>|
|[<span data-ttu-id="557e1-439">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-440">ReadItem</span></span>|
|[<span data-ttu-id="557e1-441">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-442">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-443">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="557e1-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-445">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-446">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-446">Read mode</span></span>

<span data-ttu-id="557e1-447">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-448">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-448">Compose mode</span></span>

<span data-ttu-id="557e1-449">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="557e1-450">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-450">Type</span></span>

*   <span data-ttu-id="557e1-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-452">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-452">Requirements</span></span>

|<span data-ttu-id="557e1-453">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-453">Requirement</span></span>| <span data-ttu-id="557e1-454">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-455">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-456">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-456">1.0</span></span>|
|[<span data-ttu-id="557e1-457">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-458">ReadItem</span></span>|
|[<span data-ttu-id="557e1-459">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-460">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="557e1-461">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="557e1-461">normalizedSubject: String</span></span>

<span data-ttu-id="557e1-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="557e1-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="557e1-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-466">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-466">Type</span></span>

*   <span data-ttu-id="557e1-467">String</span><span class="sxs-lookup"><span data-stu-id="557e1-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-468">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-468">Requirements</span></span>

|<span data-ttu-id="557e1-469">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-469">Requirement</span></span>| <span data-ttu-id="557e1-470">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-471">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-472">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-472">1.0</span></span>|
|[<span data-ttu-id="557e1-473">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-474">ReadItem</span></span>|
|[<span data-ttu-id="557e1-475">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-476">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-477">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="557e1-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-479">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-480">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-480">Type</span></span>

*   [<span data-ttu-id="557e1-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="557e1-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="557e1-482">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-482">Requirements</span></span>

|<span data-ttu-id="557e1-483">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-483">Requirement</span></span>| <span data-ttu-id="557e1-484">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-485">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-486">1.3</span><span class="sxs-lookup"><span data-stu-id="557e1-486">1.3</span></span>|
|[<span data-ttu-id="557e1-487">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-488">ReadItem</span></span>|
|[<span data-ttu-id="557e1-489">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-490">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-491">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="557e1-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-493">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="557e1-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="557e1-494">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-495">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-495">Read mode</span></span>

<span data-ttu-id="557e1-496">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="557e1-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="557e1-497">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-498">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-499">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-499">Compose mode</span></span>

<span data-ttu-id="557e1-500">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="557e1-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="557e1-501">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-502">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="557e1-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="557e1-503">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="557e1-504">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="557e1-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="557e1-505">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-505">Type</span></span>

*   <span data-ttu-id="557e1-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-507">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-507">Requirements</span></span>

|<span data-ttu-id="557e1-508">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-508">Requirement</span></span>| <span data-ttu-id="557e1-509">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-510">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-511">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-511">1.0</span></span>|
|[<span data-ttu-id="557e1-512">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-513">ReadItem</span></span>|
|[<span data-ttu-id="557e1-514">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-515">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="557e1-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-p128">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-519">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-519">Type</span></span>

*   [<span data-ttu-id="557e1-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="557e1-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="557e1-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-521">Requirements</span></span>

|<span data-ttu-id="557e1-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-522">Requirement</span></span>| <span data-ttu-id="557e1-523">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-525">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-525">1.0</span></span>|
|[<span data-ttu-id="557e1-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-527">ReadItem</span></span>|
|[<span data-ttu-id="557e1-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-529">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-530">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="557e1-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-532">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="557e1-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="557e1-533">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-534">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-534">Read mode</span></span>

<span data-ttu-id="557e1-535">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="557e1-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="557e1-536">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-537">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-538">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-538">Compose mode</span></span>

<span data-ttu-id="557e1-539">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="557e1-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="557e1-540">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-541">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="557e1-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="557e1-542">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="557e1-543">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="557e1-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="557e1-544">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-544">Type</span></span>

*   <span data-ttu-id="557e1-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-546">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-546">Requirements</span></span>

|<span data-ttu-id="557e1-547">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-547">Requirement</span></span>| <span data-ttu-id="557e1-548">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-549">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-550">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-550">1.0</span></span>|
|[<span data-ttu-id="557e1-551">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-552">ReadItem</span></span>|
|[<span data-ttu-id="557e1-553">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-554">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="557e1-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-p132">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="557e1-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="557e1-p133">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="557e1-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-560">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="557e1-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="557e1-561">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-561">Type</span></span>

*   [<span data-ttu-id="557e1-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="557e1-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="557e1-563">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-563">Requirements</span></span>

|<span data-ttu-id="557e1-564">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-564">Requirement</span></span>| <span data-ttu-id="557e1-565">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-566">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-567">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-567">1.0</span></span>|
|[<span data-ttu-id="557e1-568">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-569">ReadItem</span></span>|
|[<span data-ttu-id="557e1-570">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-571">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-572">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="557e1-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-574">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="557e1-p134">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="557e1-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-577">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-577">Read mode</span></span>

<span data-ttu-id="557e1-578">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="557e1-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-579">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-579">Compose mode</span></span>

<span data-ttu-id="557e1-580">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="557e1-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="557e1-581">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="557e1-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="557e1-582">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="557e1-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="557e1-583">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-583">Type</span></span>

*   <span data-ttu-id="557e1-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-585">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-585">Requirements</span></span>

|<span data-ttu-id="557e1-586">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-586">Requirement</span></span>| <span data-ttu-id="557e1-587">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-588">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-589">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-589">1.0</span></span>|
|[<span data-ttu-id="557e1-590">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-591">ReadItem</span></span>|
|[<span data-ttu-id="557e1-592">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-593">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="557e1-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-595">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="557e1-596">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="557e1-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-597">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-597">Read mode</span></span>

<span data-ttu-id="557e1-p135">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="557e1-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-600">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-600">Compose mode</span></span>

<span data-ttu-id="557e1-601">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="557e1-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="557e1-602">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-602">Type</span></span>

*   <span data-ttu-id="557e1-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-604">Requirements</span></span>

|<span data-ttu-id="557e1-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-605">Requirement</span></span>| <span data-ttu-id="557e1-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-607">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-608">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-608">1.0</span></span>|
|[<span data-ttu-id="557e1-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-610">ReadItem</span></span>|
|[<span data-ttu-id="557e1-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-612">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="557e1-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="557e1-614">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="557e1-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="557e1-615">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="557e1-616">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-616">Read mode</span></span>

<span data-ttu-id="557e1-617">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="557e1-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="557e1-618">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-619">Toutefois, sous Windows et Mac, vous pouvez configurer pour obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="557e1-620">Mode composition</span><span class="sxs-lookup"><span data-stu-id="557e1-620">Compose mode</span></span>

<span data-ttu-id="557e1-621">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="557e1-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="557e1-622">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="557e1-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="557e1-623">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="557e1-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="557e1-624">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="557e1-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="557e1-625">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="557e1-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="557e1-626">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-626">Type</span></span>

*   <span data-ttu-id="557e1-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-628">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-628">Requirements</span></span>

|<span data-ttu-id="557e1-629">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-629">Requirement</span></span>| <span data-ttu-id="557e1-630">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-631">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-632">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-632">1.0</span></span>|
|[<span data-ttu-id="557e1-633">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-634">ReadItem</span></span>|
|[<span data-ttu-id="557e1-635">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-636">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="557e1-637">Méthodes</span><span class="sxs-lookup"><span data-stu-id="557e1-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="557e1-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="557e1-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="557e1-639">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="557e1-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="557e1-640">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="557e1-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="557e1-641">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="557e1-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-642">Parameters</span><span class="sxs-lookup"><span data-stu-id="557e1-642">Parameters</span></span>

|<span data-ttu-id="557e1-643">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-643">Name</span></span>| <span data-ttu-id="557e1-644">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-644">Type</span></span>| <span data-ttu-id="557e1-645">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-645">Attributes</span></span>| <span data-ttu-id="557e1-646">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="557e1-647">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-647">String</span></span>||<span data-ttu-id="557e1-p139">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="557e1-650">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-650">String</span></span>||<span data-ttu-id="557e1-p140">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="557e1-653">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-653">Object</span></span>| <span data-ttu-id="557e1-654">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-654">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-655">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="557e1-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="557e1-656">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-656">Object</span></span> | <span data-ttu-id="557e1-657">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-657">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-658">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="557e1-659">Boolean</span><span class="sxs-lookup"><span data-stu-id="557e1-659">Boolean</span></span> | <span data-ttu-id="557e1-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-660">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-661">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="557e1-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="557e1-662">fonction</span><span class="sxs-lookup"><span data-stu-id="557e1-662">function</span></span>| <span data-ttu-id="557e1-663">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-663">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-664">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="557e1-665">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="557e1-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="557e1-666">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="557e1-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="557e1-667">Erreurs</span><span class="sxs-lookup"><span data-stu-id="557e1-667">Errors</span></span>

| <span data-ttu-id="557e1-668">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="557e1-668">Error code</span></span> | <span data-ttu-id="557e1-669">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="557e1-670">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="557e1-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="557e1-671">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="557e1-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="557e1-672">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="557e1-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="557e1-673">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-673">Requirements</span></span>

|<span data-ttu-id="557e1-674">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-674">Requirement</span></span>| <span data-ttu-id="557e1-675">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-676">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-677">1.1</span><span class="sxs-lookup"><span data-stu-id="557e1-677">1.1</span></span>|
|[<span data-ttu-id="557e1-678">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="557e1-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="557e1-680">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-681">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="557e1-682">Exemples</span><span class="sxs-lookup"><span data-stu-id="557e1-682">Examples</span></span>

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

<span data-ttu-id="557e1-683">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="557e1-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="557e1-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="557e1-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="557e1-685">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="557e1-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="557e1-689">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="557e1-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="557e1-690">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="557e1-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-691">Paramètres</span><span class="sxs-lookup"><span data-stu-id="557e1-691">Parameters</span></span>

|<span data-ttu-id="557e1-692">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-692">Name</span></span>| <span data-ttu-id="557e1-693">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-693">Type</span></span>| <span data-ttu-id="557e1-694">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-694">Attributes</span></span>| <span data-ttu-id="557e1-695">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="557e1-696">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-696">String</span></span>||<span data-ttu-id="557e1-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="557e1-699">String</span><span class="sxs-lookup"><span data-stu-id="557e1-699">String</span></span>||<span data-ttu-id="557e1-700">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="557e1-700">The subject of the item to be attached.</span></span> <span data-ttu-id="557e1-701">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="557e1-702">Object</span><span class="sxs-lookup"><span data-stu-id="557e1-702">Object</span></span>| <span data-ttu-id="557e1-703">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-703">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-704">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="557e1-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="557e1-705">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-705">Object</span></span>| <span data-ttu-id="557e1-706">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-706">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-707">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="557e1-708">fonction</span><span class="sxs-lookup"><span data-stu-id="557e1-708">function</span></span>| <span data-ttu-id="557e1-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-709">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-710">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="557e1-711">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="557e1-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="557e1-712">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="557e1-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="557e1-713">Erreurs</span><span class="sxs-lookup"><span data-stu-id="557e1-713">Errors</span></span>

| <span data-ttu-id="557e1-714">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="557e1-714">Error code</span></span> | <span data-ttu-id="557e1-715">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="557e1-716">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="557e1-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="557e1-717">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-717">Requirements</span></span>

|<span data-ttu-id="557e1-718">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-718">Requirement</span></span>| <span data-ttu-id="557e1-719">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-720">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-721">1.1</span><span class="sxs-lookup"><span data-stu-id="557e1-721">1.1</span></span>|
|[<span data-ttu-id="557e1-722">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="557e1-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="557e1-724">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-725">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-726">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-726">Example</span></span>

<span data-ttu-id="557e1-727">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="557e1-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="557e1-728">close()</span><span class="sxs-lookup"><span data-stu-id="557e1-728">close()</span></span>

<span data-ttu-id="557e1-729">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="557e1-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="557e1-p144">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="557e1-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-732">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="557e1-733">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="557e1-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-734">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-734">Requirements</span></span>

|<span data-ttu-id="557e1-735">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-735">Requirement</span></span>| <span data-ttu-id="557e1-736">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-737">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-738">1.3</span><span class="sxs-lookup"><span data-stu-id="557e1-738">1.3</span></span>|
|[<span data-ttu-id="557e1-739">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-740">Restreinte</span><span class="sxs-lookup"><span data-stu-id="557e1-740">Restricted</span></span>|
|[<span data-ttu-id="557e1-741">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-742">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="557e1-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="557e1-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="557e1-744">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="557e1-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-745">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="557e1-746">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="557e1-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="557e1-747">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="557e1-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="557e1-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="557e1-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-751">Paramètres</span><span class="sxs-lookup"><span data-stu-id="557e1-751">Parameters</span></span>

| <span data-ttu-id="557e1-752">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-752">Name</span></span> | <span data-ttu-id="557e1-753">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-753">Type</span></span> | <span data-ttu-id="557e1-754">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-754">Attributes</span></span> | <span data-ttu-id="557e1-755">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="557e1-756">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="557e1-756">String &#124; Object</span></span>| |<span data-ttu-id="557e1-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="557e1-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="557e1-759">**OU**</span><span class="sxs-lookup"><span data-stu-id="557e1-759">**OR**</span></span><br/><span data-ttu-id="557e1-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="557e1-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="557e1-762">String</span><span class="sxs-lookup"><span data-stu-id="557e1-762">String</span></span> | <span data-ttu-id="557e1-763">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-763">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="557e1-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="557e1-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="557e1-767">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-767">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-768">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="557e1-769">String</span><span class="sxs-lookup"><span data-stu-id="557e1-769">String</span></span> | | <span data-ttu-id="557e1-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="557e1-772">String</span><span class="sxs-lookup"><span data-stu-id="557e1-772">String</span></span> | | <span data-ttu-id="557e1-773">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="557e1-774">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-774">String</span></span> | | <span data-ttu-id="557e1-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="557e1-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="557e1-777">Booléen</span><span class="sxs-lookup"><span data-stu-id="557e1-777">Boolean</span></span> | | <span data-ttu-id="557e1-p151">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="557e1-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="557e1-780">String</span><span class="sxs-lookup"><span data-stu-id="557e1-780">String</span></span> | | <span data-ttu-id="557e1-p152">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="557e1-784">function</span><span class="sxs-lookup"><span data-stu-id="557e1-784">function</span></span> | <span data-ttu-id="557e1-785">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-785">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-786">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="557e1-787">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-787">Requirements</span></span>

|<span data-ttu-id="557e1-788">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-788">Requirement</span></span>| <span data-ttu-id="557e1-789">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-790">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-791">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-791">1.0</span></span>|
|[<span data-ttu-id="557e1-792">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-793">ReadItem</span></span>|
|[<span data-ttu-id="557e1-794">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-795">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="557e1-796">Exemples</span><span class="sxs-lookup"><span data-stu-id="557e1-796">Examples</span></span>

<span data-ttu-id="557e1-797">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="557e1-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="557e1-798">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="557e1-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="557e1-799">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="557e1-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="557e1-800">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="557e1-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="557e1-801">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="557e1-802">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="557e1-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="557e1-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="557e1-804">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="557e1-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-805">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="557e1-806">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="557e1-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="557e1-807">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="557e1-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="557e1-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="557e1-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-811">Paramètres</span><span class="sxs-lookup"><span data-stu-id="557e1-811">Parameters</span></span>

| <span data-ttu-id="557e1-812">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-812">Name</span></span> | <span data-ttu-id="557e1-813">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-813">Type</span></span> | <span data-ttu-id="557e1-814">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-814">Attributes</span></span> | <span data-ttu-id="557e1-815">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="557e1-816">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="557e1-816">String &#124; Object</span></span>| | <span data-ttu-id="557e1-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="557e1-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="557e1-819">**OU**</span><span class="sxs-lookup"><span data-stu-id="557e1-819">**OR**</span></span><br/><span data-ttu-id="557e1-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="557e1-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="557e1-822">String</span><span class="sxs-lookup"><span data-stu-id="557e1-822">String</span></span> | <span data-ttu-id="557e1-823">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-823">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="557e1-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="557e1-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="557e1-827">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-827">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-828">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="557e1-829">String</span><span class="sxs-lookup"><span data-stu-id="557e1-829">String</span></span> | | <span data-ttu-id="557e1-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="557e1-832">String</span><span class="sxs-lookup"><span data-stu-id="557e1-832">String</span></span> | | <span data-ttu-id="557e1-833">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="557e1-834">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-834">String</span></span> | | <span data-ttu-id="557e1-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="557e1-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="557e1-837">Booléen</span><span class="sxs-lookup"><span data-stu-id="557e1-837">Boolean</span></span> | | <span data-ttu-id="557e1-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="557e1-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="557e1-840">String</span><span class="sxs-lookup"><span data-stu-id="557e1-840">String</span></span> | | <span data-ttu-id="557e1-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="557e1-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="557e1-844">function</span><span class="sxs-lookup"><span data-stu-id="557e1-844">function</span></span> | <span data-ttu-id="557e1-845">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-845">&lt;optional&gt;</span></span> | <span data-ttu-id="557e1-846">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="557e1-847">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-847">Requirements</span></span>

|<span data-ttu-id="557e1-848">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-848">Requirement</span></span>| <span data-ttu-id="557e1-849">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-850">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-851">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-851">1.0</span></span>|
|[<span data-ttu-id="557e1-852">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-853">ReadItem</span></span>|
|[<span data-ttu-id="557e1-854">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-855">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="557e1-856">Exemples</span><span class="sxs-lookup"><span data-stu-id="557e1-856">Examples</span></span>

<span data-ttu-id="557e1-857">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="557e1-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="557e1-858">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="557e1-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="557e1-859">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="557e1-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="557e1-860">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="557e1-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="557e1-861">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="557e1-862">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="557e1-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="557e1-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="557e1-864">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="557e1-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-865">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-866">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-866">Requirements</span></span>

|<span data-ttu-id="557e1-867">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-867">Requirement</span></span>| <span data-ttu-id="557e1-868">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-869">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-870">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-870">1.0</span></span>|
|[<span data-ttu-id="557e1-871">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-872">ReadItem</span></span>|
|[<span data-ttu-id="557e1-873">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-874">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-875">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-875">Returns:</span></span>

<span data-ttu-id="557e1-876">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="557e1-877">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-877">Example</span></span>

<span data-ttu-id="557e1-878">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="557e1-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="557e1-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="557e1-880">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="557e1-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-881">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-882">Paramètres</span><span class="sxs-lookup"><span data-stu-id="557e1-882">Parameters</span></span>

|<span data-ttu-id="557e1-883">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-883">Name</span></span>| <span data-ttu-id="557e1-884">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-884">Type</span></span>| <span data-ttu-id="557e1-885">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="557e1-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="557e1-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="557e1-887">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="557e1-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="557e1-888">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-888">Requirements</span></span>

|<span data-ttu-id="557e1-889">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-889">Requirement</span></span>| <span data-ttu-id="557e1-890">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-891">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-892">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-892">1.0</span></span>|
|[<span data-ttu-id="557e1-893">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-894">Restreinte</span><span class="sxs-lookup"><span data-stu-id="557e1-894">Restricted</span></span>|
|[<span data-ttu-id="557e1-895">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-896">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-897">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-897">Returns:</span></span>

<span data-ttu-id="557e1-898">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="557e1-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="557e1-899">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="557e1-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="557e1-900">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="557e1-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="557e1-901">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="557e1-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="557e1-902">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="557e1-902">Value of `entityType`</span></span> | <span data-ttu-id="557e1-903">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="557e1-903">Type of objects in returned array</span></span> | <span data-ttu-id="557e1-904">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="557e1-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="557e1-905">String</span><span class="sxs-lookup"><span data-stu-id="557e1-905">String</span></span> | <span data-ttu-id="557e1-906">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="557e1-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="557e1-907">Contact</span><span class="sxs-lookup"><span data-stu-id="557e1-907">Contact</span></span> | <span data-ttu-id="557e1-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="557e1-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="557e1-909">String</span><span class="sxs-lookup"><span data-stu-id="557e1-909">String</span></span> | <span data-ttu-id="557e1-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="557e1-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="557e1-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="557e1-911">MeetingSuggestion</span></span> | <span data-ttu-id="557e1-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="557e1-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="557e1-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="557e1-913">PhoneNumber</span></span> | <span data-ttu-id="557e1-914">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="557e1-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="557e1-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="557e1-915">TaskSuggestion</span></span> | <span data-ttu-id="557e1-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="557e1-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="557e1-917">String</span><span class="sxs-lookup"><span data-stu-id="557e1-917">String</span></span> | <span data-ttu-id="557e1-918">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="557e1-918">**Restricted**</span></span> |

<span data-ttu-id="557e1-919">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="557e1-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="557e1-920">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-920">Example</span></span>

<span data-ttu-id="557e1-921">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="557e1-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="557e1-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="557e1-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="557e1-923">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="557e1-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-924">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="557e1-925">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="557e1-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-926">Parameters</span><span class="sxs-lookup"><span data-stu-id="557e1-926">Parameters</span></span>

|<span data-ttu-id="557e1-927">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-927">Name</span></span>| <span data-ttu-id="557e1-928">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-928">Type</span></span>| <span data-ttu-id="557e1-929">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="557e1-930">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-930">String</span></span>|<span data-ttu-id="557e1-931">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="557e1-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="557e1-932">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-932">Requirements</span></span>

|<span data-ttu-id="557e1-933">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-933">Requirement</span></span>| <span data-ttu-id="557e1-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-935">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-936">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-936">1.0</span></span>|
|[<span data-ttu-id="557e1-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-938">ReadItem</span></span>|
|[<span data-ttu-id="557e1-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-940">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-941">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-941">Returns:</span></span>

<span data-ttu-id="557e1-p162">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="557e1-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="557e1-944">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="557e1-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="557e1-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="557e1-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="557e1-946">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="557e1-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-947">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="557e1-p163">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="557e1-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="557e1-951">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="557e1-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="557e1-952">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="557e1-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="557e1-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-956">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-956">Requirements</span></span>

|<span data-ttu-id="557e1-957">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-957">Requirement</span></span>| <span data-ttu-id="557e1-958">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-959">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-960">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-960">1.0</span></span>|
|[<span data-ttu-id="557e1-961">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-962">ReadItem</span></span>|
|[<span data-ttu-id="557e1-963">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-964">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-965">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-965">Returns:</span></span>

<span data-ttu-id="557e1-p165">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="557e1-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="557e1-968">Type : Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="557e1-969">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-969">Example</span></span>

<span data-ttu-id="557e1-970">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="557e1-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="557e1-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="557e1-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="557e1-972">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="557e1-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-973">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="557e1-974">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="557e1-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="557e1-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="557e1-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-977">Parameters</span><span class="sxs-lookup"><span data-stu-id="557e1-977">Parameters</span></span>

|<span data-ttu-id="557e1-978">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-978">Name</span></span>| <span data-ttu-id="557e1-979">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-979">Type</span></span>| <span data-ttu-id="557e1-980">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="557e1-981">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-981">String</span></span>|<span data-ttu-id="557e1-982">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="557e1-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="557e1-983">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-983">Requirements</span></span>

|<span data-ttu-id="557e1-984">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-984">Requirement</span></span>| <span data-ttu-id="557e1-985">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-986">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-987">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-987">1.0</span></span>|
|[<span data-ttu-id="557e1-988">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-989">ReadItem</span></span>|
|[<span data-ttu-id="557e1-990">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-991">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-992">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-992">Returns:</span></span>

<span data-ttu-id="557e1-993">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="557e1-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="557e1-994">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="557e1-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="557e1-995">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="557e1-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="557e1-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="557e1-997">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="557e1-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="557e1-998">S’il n’y a aucune sélection, mais que le curseur se trouve dans le corps ou l’objet, la méthode renvoie une chaîne vide pour les données sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="557e1-998">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="557e1-999">Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="557e1-999">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-1000">Dans Outlook sur le Web, la méthode renvoie la chaîne « NULL » si aucun texte n’est sélectionné, mais que le curseur se trouve dans le corps.</span><span class="sxs-lookup"><span data-stu-id="557e1-1000">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="557e1-1001">Pour vérifier cette situation, reportez-vous à l’exemple plus loin dans cette section.</span><span class="sxs-lookup"><span data-stu-id="557e1-1001">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-1002">Parameters</span><span class="sxs-lookup"><span data-stu-id="557e1-1002">Parameters</span></span>

|<span data-ttu-id="557e1-1003">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-1003">Name</span></span>| <span data-ttu-id="557e1-1004">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-1004">Type</span></span>| <span data-ttu-id="557e1-1005">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-1005">Attributes</span></span>| <span data-ttu-id="557e1-1006">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-1006">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="557e1-1007">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="557e1-1007">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="557e1-p169">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="557e1-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="557e1-1011">Object</span><span class="sxs-lookup"><span data-stu-id="557e1-1011">Object</span></span>| <span data-ttu-id="557e1-1012">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1012">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1013">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="557e1-1013">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="557e1-1014">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1014">Object</span></span>| <span data-ttu-id="557e1-1015">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1015">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1016">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1016">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="557e1-1017">fonction</span><span class="sxs-lookup"><span data-stu-id="557e1-1017">function</span></span>||<span data-ttu-id="557e1-1018">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-1018">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="557e1-1019">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="557e1-1019">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="557e1-1020">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="557e1-1020">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="557e1-1021">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1021">Requirements</span></span>

|<span data-ttu-id="557e1-1022">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1022">Requirement</span></span>| <span data-ttu-id="557e1-1023">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1024">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1025">1.2</span><span class="sxs-lookup"><span data-stu-id="557e1-1025">1.2</span></span>|
|[<span data-ttu-id="557e1-1026">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1026">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1027">ReadItem</span></span>|
|[<span data-ttu-id="557e1-1028">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1028">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1029">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-1029">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-1030">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-1030">Returns:</span></span>

<span data-ttu-id="557e1-1031">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="557e1-1031">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="557e1-1032">Type : String</span><span class="sxs-lookup"><span data-stu-id="557e1-1032">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="557e1-1033">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-1033">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="557e1-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="557e1-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="557e1-1035">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="557e1-1035">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="557e1-1036">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="557e1-1036">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-1037">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-1037">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-1038">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1038">Requirements</span></span>

|<span data-ttu-id="557e1-1039">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1039">Requirement</span></span>| <span data-ttu-id="557e1-1040">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1041">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="557e1-1042">1.6</span></span> |
|[<span data-ttu-id="557e1-1043">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1044">ReadItem</span></span>|
|[<span data-ttu-id="557e1-1045">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1046">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-1047">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-1047">Returns:</span></span>

<span data-ttu-id="557e1-1048">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="557e1-1048">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="557e1-1049">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-1049">Example</span></span>

<span data-ttu-id="557e1-1050">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="557e1-1050">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="557e1-1051">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="557e1-1051">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="557e1-p172">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="557e1-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-1054">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="557e1-1054">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="557e1-p173">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="557e1-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="557e1-1058">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="557e1-1058">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="557e1-1059">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="557e1-1059">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="557e1-p174">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="557e1-1063">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1063">Requirements</span></span>

|<span data-ttu-id="557e1-1064">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1064">Requirement</span></span>| <span data-ttu-id="557e1-1065">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1066">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1067">1.6</span><span class="sxs-lookup"><span data-stu-id="557e1-1067">1.6</span></span> |
|[<span data-ttu-id="557e1-1068">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1069">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1069">ReadItem</span></span>|
|[<span data-ttu-id="557e1-1070">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1071">Lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-1071">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="557e1-1072">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="557e1-1072">Returns:</span></span>

<span data-ttu-id="557e1-p175">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="557e1-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="557e1-1075">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-1075">Example</span></span>

<span data-ttu-id="557e1-1076">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="557e1-1076">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="557e1-1077">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="557e1-1077">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="557e1-1078">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="557e1-1078">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="557e1-p176">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="557e1-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-1082">Paramètres</span><span class="sxs-lookup"><span data-stu-id="557e1-1082">Parameters</span></span>

|<span data-ttu-id="557e1-1083">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-1083">Name</span></span>| <span data-ttu-id="557e1-1084">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-1084">Type</span></span>| <span data-ttu-id="557e1-1085">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-1085">Attributes</span></span>| <span data-ttu-id="557e1-1086">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-1086">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="557e1-1087">function</span><span class="sxs-lookup"><span data-stu-id="557e1-1087">function</span></span>||<span data-ttu-id="557e1-1088">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-1088">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="557e1-1089">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="557e1-1089">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="557e1-1090">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="557e1-1090">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="557e1-1091">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1091">Object</span></span>| <span data-ttu-id="557e1-1092">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1093">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1093">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="557e1-1094">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1094">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="557e1-1095">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1095">Requirements</span></span>

|<span data-ttu-id="557e1-1096">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1096">Requirement</span></span>| <span data-ttu-id="557e1-1097">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1098">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1099">1.0</span><span class="sxs-lookup"><span data-stu-id="557e1-1099">1.0</span></span>|
|[<span data-ttu-id="557e1-1100">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1101">ReadItem</span></span>|
|[<span data-ttu-id="557e1-1102">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1103">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="557e1-1103">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-1104">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-1104">Example</span></span>

<span data-ttu-id="557e1-p179">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="557e1-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="557e1-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="557e1-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="557e1-1109">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="557e1-1109">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="557e1-1110">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-1110">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="557e1-1111">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="557e1-1111">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="557e1-1112">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="557e1-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="557e1-1113">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="557e1-1113">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-1114">Paramètres</span><span class="sxs-lookup"><span data-stu-id="557e1-1114">Parameters</span></span>

|<span data-ttu-id="557e1-1115">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-1115">Name</span></span>| <span data-ttu-id="557e1-1116">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-1116">Type</span></span>| <span data-ttu-id="557e1-1117">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-1117">Attributes</span></span>| <span data-ttu-id="557e1-1118">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="557e1-1119">Chaîne</span><span class="sxs-lookup"><span data-stu-id="557e1-1119">String</span></span>||<span data-ttu-id="557e1-1120">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="557e1-1120">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="557e1-1121">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1121">Object</span></span>| <span data-ttu-id="557e1-1122">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1123">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="557e1-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="557e1-1124">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1124">Object</span></span>| <span data-ttu-id="557e1-1125">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1126">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="557e1-1127">fonction</span><span class="sxs-lookup"><span data-stu-id="557e1-1127">function</span></span>| <span data-ttu-id="557e1-1128">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1129">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="557e1-1130">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="557e1-1130">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="557e1-1131">Erreurs</span><span class="sxs-lookup"><span data-stu-id="557e1-1131">Errors</span></span>

| <span data-ttu-id="557e1-1132">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="557e1-1132">Error code</span></span> | <span data-ttu-id="557e1-1133">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-1133">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="557e1-1134">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="557e1-1134">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="557e1-1135">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1135">Requirements</span></span>

|<span data-ttu-id="557e1-1136">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1136">Requirement</span></span>| <span data-ttu-id="557e1-1137">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1138">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1139">1.1</span><span class="sxs-lookup"><span data-stu-id="557e1-1139">1.1</span></span>|
|[<span data-ttu-id="557e1-1140">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1141">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1141">ReadWriteItem</span></span>|
|[<span data-ttu-id="557e1-1142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1143">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-1143">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-1144">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-1144">Example</span></span>

<span data-ttu-id="557e1-1145">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="557e1-1145">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="557e1-1146">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="557e1-1146">saveAsync([options], callback)</span></span>

<span data-ttu-id="557e1-1147">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="557e1-1147">Asynchronously saves an item.</span></span>

<span data-ttu-id="557e1-1148">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1148">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="557e1-1149">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="557e1-1149">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="557e1-1150">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="557e1-1150">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-1151">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="557e1-1151">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="557e1-1152">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="557e1-1152">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="557e1-p183">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="557e1-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="557e1-1156">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="557e1-1156">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="557e1-1157">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="557e1-1157">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="557e1-1158">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="557e1-1158">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="557e1-1159">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="557e1-1159">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="557e1-1160">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="557e1-1160">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-1161">Parameters</span><span class="sxs-lookup"><span data-stu-id="557e1-1161">Parameters</span></span>

|<span data-ttu-id="557e1-1162">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-1162">Name</span></span>| <span data-ttu-id="557e1-1163">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-1163">Type</span></span>| <span data-ttu-id="557e1-1164">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-1164">Attributes</span></span>| <span data-ttu-id="557e1-1165">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-1165">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="557e1-1166">Object</span><span class="sxs-lookup"><span data-stu-id="557e1-1166">Object</span></span>| <span data-ttu-id="557e1-1167">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1167">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1168">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="557e1-1168">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="557e1-1169">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1169">Object</span></span>| <span data-ttu-id="557e1-1170">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1170">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1171">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1171">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="557e1-1172">fonction</span><span class="sxs-lookup"><span data-stu-id="557e1-1172">function</span></span>||<span data-ttu-id="557e1-1173">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-1173">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="557e1-1174">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="557e1-1174">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="557e1-1175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1175">Requirements</span></span>

|<span data-ttu-id="557e1-1176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1176">Requirement</span></span>| <span data-ttu-id="557e1-1177">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1179">1.3</span><span class="sxs-lookup"><span data-stu-id="557e1-1179">1.3</span></span>|
|[<span data-ttu-id="557e1-1180">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1181">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1181">ReadWriteItem</span></span>|
|[<span data-ttu-id="557e1-1182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1183">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-1183">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="557e1-1184">範例</span><span class="sxs-lookup"><span data-stu-id="557e1-1184">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="557e1-p185">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="557e1-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="557e1-1187">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="557e1-1187">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="557e1-1188">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="557e1-1188">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="557e1-p186">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="557e1-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="557e1-1192">Parameters</span><span class="sxs-lookup"><span data-stu-id="557e1-1192">Parameters</span></span>

|<span data-ttu-id="557e1-1193">Nom</span><span class="sxs-lookup"><span data-stu-id="557e1-1193">Name</span></span>| <span data-ttu-id="557e1-1194">Type</span><span class="sxs-lookup"><span data-stu-id="557e1-1194">Type</span></span>| <span data-ttu-id="557e1-1195">Attributs</span><span class="sxs-lookup"><span data-stu-id="557e1-1195">Attributes</span></span>| <span data-ttu-id="557e1-1196">Description</span><span class="sxs-lookup"><span data-stu-id="557e1-1196">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="557e1-1197">String</span><span class="sxs-lookup"><span data-stu-id="557e1-1197">String</span></span>||<span data-ttu-id="557e1-p187">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="557e1-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="557e1-1201">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1201">Object</span></span>| <span data-ttu-id="557e1-1202">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1203">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="557e1-1203">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="557e1-1204">Objet</span><span class="sxs-lookup"><span data-stu-id="557e1-1204">Object</span></span>| <span data-ttu-id="557e1-1205">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1206">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="557e1-1206">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="557e1-1207">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="557e1-1207">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="557e1-1208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="557e1-1208">&lt;optional&gt;</span></span>|<span data-ttu-id="557e1-1209">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="557e1-1209">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="557e1-1210">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="557e1-1210">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="557e1-1211">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="557e1-1211">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="557e1-1212">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="557e1-1212">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="557e1-1213">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="557e1-1213">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="557e1-1214">fonction</span><span class="sxs-lookup"><span data-stu-id="557e1-1214">function</span></span>||<span data-ttu-id="557e1-1215">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="557e1-1215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="557e1-1216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="557e1-1216">Requirements</span></span>

|<span data-ttu-id="557e1-1217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="557e1-1217">Requirement</span></span>| <span data-ttu-id="557e1-1218">Valeur</span><span class="sxs-lookup"><span data-stu-id="557e1-1218">Value</span></span>|
|---|---|
|[<span data-ttu-id="557e1-1219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="557e1-1219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="557e1-1220">1.2</span><span class="sxs-lookup"><span data-stu-id="557e1-1220">1.2</span></span>|
|[<span data-ttu-id="557e1-1221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="557e1-1221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="557e1-1222">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="557e1-1222">ReadWriteItem</span></span>|
|[<span data-ttu-id="557e1-1223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="557e1-1223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="557e1-1224">Composition</span><span class="sxs-lookup"><span data-stu-id="557e1-1224">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="557e1-1225">Exemple</span><span class="sxs-lookup"><span data-stu-id="557e1-1225">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
