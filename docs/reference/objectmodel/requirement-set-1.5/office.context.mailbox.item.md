---
title: Office.context.mailbox.item - ensemble de conditions requises 1.5
description: ''
ms.date: 10/23/2019
localization_priority: Priority
ms.openlocfilehash: 7d585d3fd60d51b68d86b632701e8ac512fe708c
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682521"
---
# <a name="item"></a><span data-ttu-id="7a23e-102">élément</span><span class="sxs-lookup"><span data-stu-id="7a23e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="7a23e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="7a23e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="7a23e-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="7a23e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-106">Requirements</span></span>

|<span data-ttu-id="7a23e-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-107">Requirement</span></span>| <span data-ttu-id="7a23e-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-110">1.0</span></span>|
|[<span data-ttu-id="7a23e-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7a23e-112">Restricted</span></span>|
|[<span data-ttu-id="7a23e-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7a23e-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="7a23e-115">Members and methods</span></span>

| <span data-ttu-id="7a23e-116">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-116">Member</span></span> | <span data-ttu-id="7a23e-117">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7a23e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="7a23e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="7a23e-119">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-119">Member</span></span> |
| [<span data-ttu-id="7a23e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="7a23e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="7a23e-121">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-121">Member</span></span> |
| [<span data-ttu-id="7a23e-122">body</span><span class="sxs-lookup"><span data-stu-id="7a23e-122">body</span></span>](#body-body) | <span data-ttu-id="7a23e-123">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-123">Member</span></span> |
| [<span data-ttu-id="7a23e-124">cc</span><span class="sxs-lookup"><span data-stu-id="7a23e-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7a23e-125">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-125">Member</span></span> |
| [<span data-ttu-id="7a23e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="7a23e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="7a23e-127">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-127">Member</span></span> |
| [<span data-ttu-id="7a23e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="7a23e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="7a23e-129">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-129">Member</span></span> |
| [<span data-ttu-id="7a23e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="7a23e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="7a23e-131">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-131">Member</span></span> |
| [<span data-ttu-id="7a23e-132">end</span><span class="sxs-lookup"><span data-stu-id="7a23e-132">end</span></span>](#end-datetime) | <span data-ttu-id="7a23e-133">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-133">Member</span></span> |
| [<span data-ttu-id="7a23e-134">from</span><span class="sxs-lookup"><span data-stu-id="7a23e-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="7a23e-135">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-135">Member</span></span> |
| [<span data-ttu-id="7a23e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="7a23e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="7a23e-137">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-137">Member</span></span> |
| [<span data-ttu-id="7a23e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="7a23e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="7a23e-139">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-139">Member</span></span> |
| [<span data-ttu-id="7a23e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="7a23e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="7a23e-141">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-141">Member</span></span> |
| [<span data-ttu-id="7a23e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="7a23e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="7a23e-143">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-143">Member</span></span> |
| [<span data-ttu-id="7a23e-144">location</span><span class="sxs-lookup"><span data-stu-id="7a23e-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="7a23e-145">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-145">Member</span></span> |
| [<span data-ttu-id="7a23e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="7a23e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="7a23e-147">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-147">Member</span></span> |
| [<span data-ttu-id="7a23e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="7a23e-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="7a23e-149">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-149">Member</span></span> |
| [<span data-ttu-id="7a23e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="7a23e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7a23e-151">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-151">Member</span></span> |
| [<span data-ttu-id="7a23e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="7a23e-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="7a23e-153">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-153">Member</span></span> |
| [<span data-ttu-id="7a23e-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="7a23e-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7a23e-155">Member</span><span class="sxs-lookup"><span data-stu-id="7a23e-155">Member</span></span> |
| [<span data-ttu-id="7a23e-156">sender</span><span class="sxs-lookup"><span data-stu-id="7a23e-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="7a23e-157">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-157">Member</span></span> |
| [<span data-ttu-id="7a23e-158">start</span><span class="sxs-lookup"><span data-stu-id="7a23e-158">start</span></span>](#start-datetime) | <span data-ttu-id="7a23e-159">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-159">Member</span></span> |
| [<span data-ttu-id="7a23e-160">subject</span><span class="sxs-lookup"><span data-stu-id="7a23e-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="7a23e-161">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-161">Member</span></span> |
| [<span data-ttu-id="7a23e-162">to</span><span class="sxs-lookup"><span data-stu-id="7a23e-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7a23e-163">Membre</span><span class="sxs-lookup"><span data-stu-id="7a23e-163">Member</span></span> |
| [<span data-ttu-id="7a23e-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="7a23e-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-165">Method</span></span> |
| [<span data-ttu-id="7a23e-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="7a23e-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-167">Method</span></span> |
| [<span data-ttu-id="7a23e-168">close</span><span class="sxs-lookup"><span data-stu-id="7a23e-168">close</span></span>](#close) | <span data-ttu-id="7a23e-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-169">Method</span></span> |
| [<span data-ttu-id="7a23e-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="7a23e-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="7a23e-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-171">Method</span></span> |
| [<span data-ttu-id="7a23e-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="7a23e-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="7a23e-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-173">Method</span></span> |
| [<span data-ttu-id="7a23e-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="7a23e-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="7a23e-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-175">Method</span></span> |
| [<span data-ttu-id="7a23e-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="7a23e-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="7a23e-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-177">Method</span></span> |
| [<span data-ttu-id="7a23e-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="7a23e-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="7a23e-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-179">Method</span></span> |
| [<span data-ttu-id="7a23e-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="7a23e-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="7a23e-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-181">Method</span></span> |
| [<span data-ttu-id="7a23e-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="7a23e-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="7a23e-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-183">Method</span></span> |
| [<span data-ttu-id="7a23e-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="7a23e-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-185">Method</span></span> |
| [<span data-ttu-id="7a23e-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="7a23e-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-187">Method</span></span> |
| [<span data-ttu-id="7a23e-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="7a23e-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-189">Method</span></span> |
| [<span data-ttu-id="7a23e-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="7a23e-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-191">Method</span></span> |
| [<span data-ttu-id="7a23e-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7a23e-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="7a23e-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="7a23e-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="7a23e-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-194">Example</span></span>

<span data-ttu-id="7a23e-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="7a23e-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="7a23e-196">Members</span><span class="sxs-lookup"><span data-stu-id="7a23e-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="7a23e-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="7a23e-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="7a23e-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="7a23e-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="7a23e-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="7a23e-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-202">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-202">Type</span></span>

*   <span data-ttu-id="7a23e-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="7a23e-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-204">Requirements</span></span>

|<span data-ttu-id="7a23e-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-205">Requirement</span></span>| <span data-ttu-id="7a23e-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-208">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-208">1.0</span></span>|
|[<span data-ttu-id="7a23e-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-210">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-213">Example</span></span>

<span data-ttu-id="7a23e-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="7a23e-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="7a23e-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-217">Compose mode only.</span></span>

<span data-ttu-id="7a23e-218">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-218">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-219">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="7a23e-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7a23e-220">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="7a23e-221">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="7a23e-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-222">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-222">Type</span></span>

*   [<span data-ttu-id="7a23e-223">Destinataires</span><span class="sxs-lookup"><span data-stu-id="7a23e-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-224">Requirements</span></span>

|<span data-ttu-id="7a23e-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-225">Requirement</span></span>| <span data-ttu-id="7a23e-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-228">1.1</span><span class="sxs-lookup"><span data-stu-id="7a23e-228">1.1</span></span>|
|[<span data-ttu-id="7a23e-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-230">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-232">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-233">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="7a23e-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-235">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-236">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-236">Type</span></span>

*   [<span data-ttu-id="7a23e-237">Body</span><span class="sxs-lookup"><span data-stu-id="7a23e-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-238">Requirements</span></span>

|<span data-ttu-id="7a23e-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-239">Requirement</span></span>| <span data-ttu-id="7a23e-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-242">1.1</span><span class="sxs-lookup"><span data-stu-id="7a23e-242">1.1</span></span>|
|[<span data-ttu-id="7a23e-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-244">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-247">Example</span></span>

<span data-ttu-id="7a23e-248">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="7a23e-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="7a23e-249">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="7a23e-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-251">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="7a23e-252">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-253">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-253">Read mode</span></span>

<span data-ttu-id="7a23e-254">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="7a23e-255">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-255">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-256">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-257">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-257">Compose mode</span></span>

<span data-ttu-id="7a23e-258">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="7a23e-259">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-259">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-260">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="7a23e-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7a23e-261">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="7a23e-262">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="7a23e-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7a23e-263">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-263">Type</span></span>

*   <span data-ttu-id="7a23e-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-265">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-265">Requirements</span></span>

|<span data-ttu-id="7a23e-266">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-266">Requirement</span></span>| <span data-ttu-id="7a23e-267">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-268">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-269">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-269">1.0</span></span>|
|[<span data-ttu-id="7a23e-270">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-271">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-272">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-273">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="7a23e-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="7a23e-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="7a23e-275">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="7a23e-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="7a23e-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="7a23e-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-280">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-280">Type</span></span>

*   <span data-ttu-id="7a23e-281">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-282">Requirements</span></span>

|<span data-ttu-id="7a23e-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-283">Requirement</span></span>| <span data-ttu-id="7a23e-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-286">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-286">1.0</span></span>|
|[<span data-ttu-id="7a23e-287">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-288">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-289">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-290">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-291">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="7a23e-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="7a23e-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="7a23e-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-295">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-295">Type</span></span>

*   <span data-ttu-id="7a23e-296">Date</span><span class="sxs-lookup"><span data-stu-id="7a23e-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-297">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-297">Requirements</span></span>

|<span data-ttu-id="7a23e-298">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-298">Requirement</span></span>| <span data-ttu-id="7a23e-299">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-300">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-301">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-301">1.0</span></span>|
|[<span data-ttu-id="7a23e-302">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-303">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-304">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-305">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-306">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="7a23e-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="7a23e-307">dateTimeModified: Date</span></span>

<span data-ttu-id="7a23e-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-310">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-311">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-311">Type</span></span>

*   <span data-ttu-id="7a23e-312">Date</span><span class="sxs-lookup"><span data-stu-id="7a23e-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-313">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-313">Requirements</span></span>

|<span data-ttu-id="7a23e-314">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-314">Requirement</span></span>| <span data-ttu-id="7a23e-315">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-316">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-317">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-317">1.0</span></span>|
|[<span data-ttu-id="7a23e-318">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-319">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-320">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-321">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-322">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="7a23e-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-324">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="7a23e-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-327">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-327">Read mode</span></span>

<span data-ttu-id="7a23e-328">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-329">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-329">Compose mode</span></span>

<span data-ttu-id="7a23e-330">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="7a23e-331">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="7a23e-332">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7a23e-333">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-333">Type</span></span>

*   <span data-ttu-id="7a23e-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-335">Requirements</span></span>

|<span data-ttu-id="7a23e-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-336">Requirement</span></span>| <span data-ttu-id="7a23e-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-339">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-339">1.0</span></span>|
|[<span data-ttu-id="7a23e-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-341">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-343">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="7a23e-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-p114">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="7a23e-p115">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-349">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-350">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-350">Type</span></span>

*   [<span data-ttu-id="7a23e-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7a23e-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-352">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-352">Requirements</span></span>

|<span data-ttu-id="7a23e-353">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-353">Requirement</span></span>| <span data-ttu-id="7a23e-354">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-355">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-356">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-356">1.0</span></span>|
|[<span data-ttu-id="7a23e-357">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-358">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-359">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-360">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-361">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="7a23e-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="7a23e-362">internetMessageId: String</span></span>

<span data-ttu-id="7a23e-p116">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-365">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-365">Type</span></span>

*   <span data-ttu-id="7a23e-366">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-367">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-367">Requirements</span></span>

|<span data-ttu-id="7a23e-368">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-368">Requirement</span></span>| <span data-ttu-id="7a23e-369">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-370">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-371">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-371">1.0</span></span>|
|[<span data-ttu-id="7a23e-372">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-373">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-374">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-375">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-376">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="7a23e-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="7a23e-377">itemClass: String</span></span>

<span data-ttu-id="7a23e-p117">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="7a23e-p118">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="7a23e-382">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-382">Type</span></span> | <span data-ttu-id="7a23e-383">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-383">Description</span></span> | <span data-ttu-id="7a23e-384">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="7a23e-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="7a23e-385">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="7a23e-385">Appointment items</span></span> | <span data-ttu-id="7a23e-386">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="7a23e-387">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="7a23e-387">Message items</span></span> | <span data-ttu-id="7a23e-388">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="7a23e-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="7a23e-389">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-390">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-390">Type</span></span>

*   <span data-ttu-id="7a23e-391">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-392">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-392">Requirements</span></span>

|<span data-ttu-id="7a23e-393">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-393">Requirement</span></span>| <span data-ttu-id="7a23e-394">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-395">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-396">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-396">1.0</span></span>|
|[<span data-ttu-id="7a23e-397">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-398">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-400">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-401">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="7a23e-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="7a23e-402">(nullable) itemId: String</span></span>

<span data-ttu-id="7a23e-p119">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-405">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="7a23e-405">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="7a23e-406">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="7a23e-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="7a23e-407">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="7a23e-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="7a23e-408">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="7a23e-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="7a23e-p121">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-411">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-411">Type</span></span>

*   <span data-ttu-id="7a23e-412">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-413">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-413">Requirements</span></span>

|<span data-ttu-id="7a23e-414">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-414">Requirement</span></span>| <span data-ttu-id="7a23e-415">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-416">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-417">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-417">1.0</span></span>|
|[<span data-ttu-id="7a23e-418">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-419">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-420">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-421">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-422">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-422">Example</span></span>

<span data-ttu-id="7a23e-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="7a23e-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-426">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="7a23e-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="7a23e-427">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-428">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-428">Type</span></span>

*   [<span data-ttu-id="7a23e-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="7a23e-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-430">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-430">Requirements</span></span>

|<span data-ttu-id="7a23e-431">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-431">Requirement</span></span>| <span data-ttu-id="7a23e-432">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-433">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-434">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-434">1.0</span></span>|
|[<span data-ttu-id="7a23e-435">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-436">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-437">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-438">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-439">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-439">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="7a23e-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-441">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-442">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-442">Read mode</span></span>

<span data-ttu-id="7a23e-443">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-444">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-444">Compose mode</span></span>

<span data-ttu-id="7a23e-445">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7a23e-446">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-446">Type</span></span>

*   <span data-ttu-id="7a23e-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-448">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-448">Requirements</span></span>

|<span data-ttu-id="7a23e-449">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-449">Requirement</span></span>| <span data-ttu-id="7a23e-450">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-451">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-452">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-452">1.0</span></span>|
|[<span data-ttu-id="7a23e-453">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-454">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-455">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-456">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="7a23e-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="7a23e-457">normalizedSubject: String</span></span>

<span data-ttu-id="7a23e-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="7a23e-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="7a23e-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-462">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-462">Type</span></span>

*   <span data-ttu-id="7a23e-463">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-464">Requirements</span></span>

|<span data-ttu-id="7a23e-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-465">Requirement</span></span>| <span data-ttu-id="7a23e-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-468">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-468">1.0</span></span>|
|[<span data-ttu-id="7a23e-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-470">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-472">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="7a23e-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-475">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-476">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-476">Type</span></span>

*   [<span data-ttu-id="7a23e-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="7a23e-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-478">Requirements</span></span>

|<span data-ttu-id="7a23e-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-479">Requirement</span></span>| <span data-ttu-id="7a23e-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-482">1.3</span><span class="sxs-lookup"><span data-stu-id="7a23e-482">1.3</span></span>|
|[<span data-ttu-id="7a23e-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-484">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-486">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-487">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="7a23e-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-489">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="7a23e-490">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-491">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-491">Read mode</span></span>

<span data-ttu-id="7a23e-492">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="7a23e-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="7a23e-493">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-493">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-494">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-495">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-495">Compose mode</span></span>

<span data-ttu-id="7a23e-496">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="7a23e-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="7a23e-497">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-497">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-498">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="7a23e-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7a23e-499">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="7a23e-500">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="7a23e-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7a23e-501">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-501">Type</span></span>

*   <span data-ttu-id="7a23e-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-503">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-503">Requirements</span></span>

|<span data-ttu-id="7a23e-504">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-504">Requirement</span></span>| <span data-ttu-id="7a23e-505">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-506">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-507">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-507">1.0</span></span>|
|[<span data-ttu-id="7a23e-508">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-509">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-510">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-511">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="7a23e-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-p128">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-515">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-515">Type</span></span>

*   [<span data-ttu-id="7a23e-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7a23e-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-517">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-517">Requirements</span></span>

|<span data-ttu-id="7a23e-518">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-518">Requirement</span></span>| <span data-ttu-id="7a23e-519">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-520">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-521">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-521">1.0</span></span>|
|[<span data-ttu-id="7a23e-522">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-523">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-524">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-525">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-526">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="7a23e-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-528">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="7a23e-529">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-530">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-530">Read mode</span></span>

<span data-ttu-id="7a23e-531">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="7a23e-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="7a23e-532">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-532">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-533">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-534">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-534">Compose mode</span></span>

<span data-ttu-id="7a23e-535">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="7a23e-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="7a23e-536">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-536">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-537">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="7a23e-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7a23e-538">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="7a23e-539">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="7a23e-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="7a23e-540">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-540">Type</span></span>

*   <span data-ttu-id="7a23e-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-542">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-542">Requirements</span></span>

|<span data-ttu-id="7a23e-543">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-543">Requirement</span></span>| <span data-ttu-id="7a23e-544">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-545">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-546">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-546">1.0</span></span>|
|[<span data-ttu-id="7a23e-547">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-548">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-549">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-550">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="7a23e-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-p132">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="7a23e-p133">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-556">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7a23e-557">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-557">Type</span></span>

*   [<span data-ttu-id="7a23e-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7a23e-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="7a23e-559">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-559">Requirements</span></span>

|<span data-ttu-id="7a23e-560">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-560">Requirement</span></span>| <span data-ttu-id="7a23e-561">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-562">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-563">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-563">1.0</span></span>|
|[<span data-ttu-id="7a23e-564">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-565">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-566">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-567">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-568">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="7a23e-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-570">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="7a23e-p134">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-573">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-573">Read mode</span></span>

<span data-ttu-id="7a23e-574">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-575">Compose mode</span></span>

<span data-ttu-id="7a23e-576">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="7a23e-577">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="7a23e-578">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7a23e-579">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-579">Type</span></span>

*   <span data-ttu-id="7a23e-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-581">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-581">Requirements</span></span>

|<span data-ttu-id="7a23e-582">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-582">Requirement</span></span>| <span data-ttu-id="7a23e-583">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-584">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-585">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-585">1.0</span></span>|
|[<span data-ttu-id="7a23e-586">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-587">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-588">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-589">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="7a23e-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-591">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="7a23e-592">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="7a23e-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-593">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-593">Read mode</span></span>

<span data-ttu-id="7a23e-p135">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-596">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-596">Compose mode</span></span>

<span data-ttu-id="7a23e-597">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="7a23e-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="7a23e-598">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-598">Type</span></span>

*   <span data-ttu-id="7a23e-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-600">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-600">Requirements</span></span>

|<span data-ttu-id="7a23e-601">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-601">Requirement</span></span>| <span data-ttu-id="7a23e-602">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-603">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-604">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-604">1.0</span></span>|
|[<span data-ttu-id="7a23e-605">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-606">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-607">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-608">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="7a23e-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="7a23e-610">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="7a23e-611">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7a23e-612">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-612">Read mode</span></span>

<span data-ttu-id="7a23e-613">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="7a23e-614">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-614">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-615">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="7a23e-616">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-616">Compose mode</span></span>

<span data-ttu-id="7a23e-617">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="7a23e-618">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7a23e-618">The collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7a23e-619">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="7a23e-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7a23e-620">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="7a23e-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="7a23e-621">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="7a23e-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7a23e-622">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-622">Type</span></span>

*   <span data-ttu-id="7a23e-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-624">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-624">Requirements</span></span>

|<span data-ttu-id="7a23e-625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-625">Requirement</span></span>| <span data-ttu-id="7a23e-626">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-627">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-628">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-628">1.0</span></span>|
|[<span data-ttu-id="7a23e-629">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-630">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-631">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-632">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="7a23e-633">Méthodes</span><span class="sxs-lookup"><span data-stu-id="7a23e-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="7a23e-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7a23e-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7a23e-635">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7a23e-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="7a23e-636">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="7a23e-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="7a23e-637">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7a23e-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-638">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-638">Parameters</span></span>

|<span data-ttu-id="7a23e-639">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-639">Name</span></span>| <span data-ttu-id="7a23e-640">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-640">Type</span></span>| <span data-ttu-id="7a23e-641">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-641">Attributes</span></span>| <span data-ttu-id="7a23e-642">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="7a23e-643">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-643">String</span></span>||<span data-ttu-id="7a23e-p139">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7a23e-646">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-646">String</span></span>||<span data-ttu-id="7a23e-p140">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7a23e-649">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-649">Object</span></span>| <span data-ttu-id="7a23e-650">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-650">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-651">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-651">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="7a23e-652">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-652">Object</span></span> | <span data-ttu-id="7a23e-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-653">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-654">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-654">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="7a23e-655">Boolean</span><span class="sxs-lookup"><span data-stu-id="7a23e-655">Boolean</span></span> | <span data-ttu-id="7a23e-656">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-656">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-657">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-657">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="7a23e-658">fonction</span><span class="sxs-lookup"><span data-stu-id="7a23e-658">function</span></span>| <span data-ttu-id="7a23e-659">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-659">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-660">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-660">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7a23e-661">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-661">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7a23e-662">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-662">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7a23e-663">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7a23e-663">Errors</span></span>

| <span data-ttu-id="7a23e-664">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="7a23e-664">Error code</span></span> | <span data-ttu-id="7a23e-665">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-665">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="7a23e-666">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="7a23e-666">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="7a23e-667">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="7a23e-667">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7a23e-668">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-668">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a23e-669">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-669">Requirements</span></span>

|<span data-ttu-id="7a23e-670">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-670">Requirement</span></span>| <span data-ttu-id="7a23e-671">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-671">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-672">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-672">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-673">1.1</span><span class="sxs-lookup"><span data-stu-id="7a23e-673">1.1</span></span>|
|[<span data-ttu-id="7a23e-674">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-674">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-675">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-675">ReadWriteItem</span></span>|
|[<span data-ttu-id="7a23e-676">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-676">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-677">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-677">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="7a23e-678">Exemples</span><span class="sxs-lookup"><span data-stu-id="7a23e-678">Examples</span></span>

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

<span data-ttu-id="7a23e-679">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-679">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="7a23e-680">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7a23e-680">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7a23e-681">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-681">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="7a23e-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="7a23e-685">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7a23e-685">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="7a23e-686">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="7a23e-686">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-687">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-687">Parameters</span></span>

|<span data-ttu-id="7a23e-688">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-688">Name</span></span>| <span data-ttu-id="7a23e-689">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-689">Type</span></span>| <span data-ttu-id="7a23e-690">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-690">Attributes</span></span>| <span data-ttu-id="7a23e-691">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-691">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="7a23e-692">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-692">String</span></span>||<span data-ttu-id="7a23e-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7a23e-695">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-695">String</span></span>||<span data-ttu-id="7a23e-696">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="7a23e-696">The subject of the item to be attached.</span></span> <span data-ttu-id="7a23e-697">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-697">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7a23e-698">Object</span><span class="sxs-lookup"><span data-stu-id="7a23e-698">Object</span></span>| <span data-ttu-id="7a23e-699">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-699">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-700">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-700">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7a23e-701">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-701">Object</span></span>| <span data-ttu-id="7a23e-702">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-702">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-703">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-703">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7a23e-704">fonction</span><span class="sxs-lookup"><span data-stu-id="7a23e-704">function</span></span>| <span data-ttu-id="7a23e-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-705">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-706">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-706">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7a23e-707">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-707">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7a23e-708">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-708">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7a23e-709">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7a23e-709">Errors</span></span>

| <span data-ttu-id="7a23e-710">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="7a23e-710">Error code</span></span> | <span data-ttu-id="7a23e-711">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-711">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7a23e-712">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-712">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a23e-713">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-713">Requirements</span></span>

|<span data-ttu-id="7a23e-714">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-714">Requirement</span></span>| <span data-ttu-id="7a23e-715">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-716">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-717">1.1</span><span class="sxs-lookup"><span data-stu-id="7a23e-717">1.1</span></span>|
|[<span data-ttu-id="7a23e-718">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-719">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-719">ReadWriteItem</span></span>|
|[<span data-ttu-id="7a23e-720">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-721">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-721">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-722">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-722">Example</span></span>

<span data-ttu-id="7a23e-723">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-723">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="7a23e-724">close()</span><span class="sxs-lookup"><span data-stu-id="7a23e-724">close()</span></span>

<span data-ttu-id="7a23e-725">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="7a23e-725">Closes the current item that is being composed.</span></span>

<span data-ttu-id="7a23e-p144">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-728">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-728">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="7a23e-729">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="7a23e-729">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-730">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-730">Requirements</span></span>

|<span data-ttu-id="7a23e-731">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-731">Requirement</span></span>| <span data-ttu-id="7a23e-732">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-732">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-733">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-733">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-734">1.3</span><span class="sxs-lookup"><span data-stu-id="7a23e-734">1.3</span></span>|
|[<span data-ttu-id="7a23e-735">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-735">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-736">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7a23e-736">Restricted</span></span>|
|[<span data-ttu-id="7a23e-737">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-737">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-738">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-738">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="7a23e-739">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7a23e-739">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="7a23e-740">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7a23e-740">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-741">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-741">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7a23e-742">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-742">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7a23e-743">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="7a23e-743">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="7a23e-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-747">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-747">Parameters</span></span>

| <span data-ttu-id="7a23e-748">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-748">Name</span></span> | <span data-ttu-id="7a23e-749">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-749">Type</span></span> | <span data-ttu-id="7a23e-750">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-750">Attributes</span></span> | <span data-ttu-id="7a23e-751">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-751">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="7a23e-752">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7a23e-752">String &#124; Object</span></span>| |<span data-ttu-id="7a23e-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7a23e-755">**OU**</span><span class="sxs-lookup"><span data-stu-id="7a23e-755">**OR**</span></span><br/><span data-ttu-id="7a23e-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="7a23e-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7a23e-758">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-758">String</span></span> | <span data-ttu-id="7a23e-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-759">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7a23e-762">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-762">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7a23e-763">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-763">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-764">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-764">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7a23e-765">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7a23e-765">String</span></span> | | <span data-ttu-id="7a23e-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7a23e-768">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-768">String</span></span> | | <span data-ttu-id="7a23e-769">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-769">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7a23e-770">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7a23e-770">String</span></span> | | <span data-ttu-id="7a23e-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="7a23e-773">Booléen</span><span class="sxs-lookup"><span data-stu-id="7a23e-773">Boolean</span></span> | | <span data-ttu-id="7a23e-p151">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7a23e-776">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-776">String</span></span> | | <span data-ttu-id="7a23e-p152">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7a23e-780">function</span><span class="sxs-lookup"><span data-stu-id="7a23e-780">function</span></span> | <span data-ttu-id="7a23e-781">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-781">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-782">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-782">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a23e-783">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-783">Requirements</span></span>

|<span data-ttu-id="7a23e-784">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-784">Requirement</span></span>| <span data-ttu-id="7a23e-785">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-786">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-787">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-787">1.0</span></span>|
|[<span data-ttu-id="7a23e-788">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-788">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-789">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-790">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-790">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-791">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-791">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7a23e-792">Exemples</span><span class="sxs-lookup"><span data-stu-id="7a23e-792">Examples</span></span>

<span data-ttu-id="7a23e-793">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-793">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="7a23e-794">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="7a23e-794">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="7a23e-795">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="7a23e-795">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7a23e-796">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="7a23e-796">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7a23e-797">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-797">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7a23e-798">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-798">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="7a23e-799">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7a23e-799">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="7a23e-800">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7a23e-800">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-801">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-801">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7a23e-802">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-802">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7a23e-803">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="7a23e-803">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="7a23e-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-807">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-807">Parameters</span></span>

| <span data-ttu-id="7a23e-808">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-808">Name</span></span> | <span data-ttu-id="7a23e-809">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-809">Type</span></span> | <span data-ttu-id="7a23e-810">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-810">Attributes</span></span> | <span data-ttu-id="7a23e-811">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-811">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="7a23e-812">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7a23e-812">String &#124; Object</span></span>| | <span data-ttu-id="7a23e-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7a23e-815">**OU**</span><span class="sxs-lookup"><span data-stu-id="7a23e-815">**OR**</span></span><br/><span data-ttu-id="7a23e-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="7a23e-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7a23e-818">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-818">String</span></span> | <span data-ttu-id="7a23e-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-819">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7a23e-822">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-822">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7a23e-823">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-823">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-824">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-824">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7a23e-825">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7a23e-825">String</span></span> | | <span data-ttu-id="7a23e-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7a23e-828">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-828">String</span></span> | | <span data-ttu-id="7a23e-829">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-829">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7a23e-830">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7a23e-830">String</span></span> | | <span data-ttu-id="7a23e-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="7a23e-833">Booléen</span><span class="sxs-lookup"><span data-stu-id="7a23e-833">Boolean</span></span> | | <span data-ttu-id="7a23e-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7a23e-836">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-836">String</span></span> | | <span data-ttu-id="7a23e-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7a23e-840">function</span><span class="sxs-lookup"><span data-stu-id="7a23e-840">function</span></span> | <span data-ttu-id="7a23e-841">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-841">&lt;optional&gt;</span></span> | <span data-ttu-id="7a23e-842">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a23e-843">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-843">Requirements</span></span>

|<span data-ttu-id="7a23e-844">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-844">Requirement</span></span>| <span data-ttu-id="7a23e-845">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-845">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-846">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-846">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-847">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-847">1.0</span></span>|
|[<span data-ttu-id="7a23e-848">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-848">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-849">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-849">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-850">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-850">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-851">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-851">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7a23e-852">Exemples</span><span class="sxs-lookup"><span data-stu-id="7a23e-852">Examples</span></span>

<span data-ttu-id="7a23e-853">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-853">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="7a23e-854">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="7a23e-854">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="7a23e-855">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="7a23e-855">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7a23e-856">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="7a23e-856">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7a23e-857">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-857">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7a23e-858">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-858">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="7a23e-859">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="7a23e-859">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="7a23e-860">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7a23e-860">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-861">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-861">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-862">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-862">Requirements</span></span>

|<span data-ttu-id="7a23e-863">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-863">Requirement</span></span>| <span data-ttu-id="7a23e-864">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-865">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-866">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-866">1.0</span></span>|
|[<span data-ttu-id="7a23e-867">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-868">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-868">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-869">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-870">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a23e-871">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7a23e-871">Returns:</span></span>

<span data-ttu-id="7a23e-872">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="7a23e-872">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="7a23e-873">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-873">Example</span></span>

<span data-ttu-id="7a23e-874">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-874">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="7a23e-875">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="7a23e-875">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="7a23e-876">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7a23e-876">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-877">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-877">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-878">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-878">Parameters</span></span>

|<span data-ttu-id="7a23e-879">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-879">Name</span></span>| <span data-ttu-id="7a23e-880">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-880">Type</span></span>| <span data-ttu-id="7a23e-881">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-881">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="7a23e-882">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="7a23e-882">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="7a23e-883">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="7a23e-883">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a23e-884">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-884">Requirements</span></span>

|<span data-ttu-id="7a23e-885">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-885">Requirement</span></span>| <span data-ttu-id="7a23e-886">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-887">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-888">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-888">1.0</span></span>|
|[<span data-ttu-id="7a23e-889">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-890">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7a23e-890">Restricted</span></span>|
|[<span data-ttu-id="7a23e-891">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-892">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-892">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a23e-893">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7a23e-893">Returns:</span></span>

<span data-ttu-id="7a23e-894">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="7a23e-894">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="7a23e-895">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="7a23e-895">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="7a23e-896">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-896">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="7a23e-897">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="7a23e-897">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="7a23e-898">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="7a23e-898">Value of `entityType`</span></span> | <span data-ttu-id="7a23e-899">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="7a23e-899">Type of objects in returned array</span></span> | <span data-ttu-id="7a23e-900">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="7a23e-900">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="7a23e-901">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-901">String</span></span> | <span data-ttu-id="7a23e-902">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7a23e-902">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="7a23e-903">Contact</span><span class="sxs-lookup"><span data-stu-id="7a23e-903">Contact</span></span> | <span data-ttu-id="7a23e-904">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7a23e-904">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="7a23e-905">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-905">String</span></span> | <span data-ttu-id="7a23e-906">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7a23e-906">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="7a23e-907">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="7a23e-907">MeetingSuggestion</span></span> | <span data-ttu-id="7a23e-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7a23e-908">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="7a23e-909">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="7a23e-909">PhoneNumber</span></span> | <span data-ttu-id="7a23e-910">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7a23e-910">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="7a23e-911">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="7a23e-911">TaskSuggestion</span></span> | <span data-ttu-id="7a23e-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7a23e-912">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="7a23e-913">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-913">String</span></span> | <span data-ttu-id="7a23e-914">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7a23e-914">**Restricted**</span></span> |

<span data-ttu-id="7a23e-915">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="7a23e-915">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="7a23e-916">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-916">Example</span></span>

<span data-ttu-id="7a23e-917">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7a23e-917">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="7a23e-918">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="7a23e-918">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="7a23e-919">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7a23e-919">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-920">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7a23e-921">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="7a23e-921">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-922">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-922">Parameters</span></span>

|<span data-ttu-id="7a23e-923">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-923">Name</span></span>| <span data-ttu-id="7a23e-924">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-924">Type</span></span>| <span data-ttu-id="7a23e-925">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-925">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7a23e-926">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-926">String</span></span>|<span data-ttu-id="7a23e-927">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="7a23e-927">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a23e-928">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-928">Requirements</span></span>

|<span data-ttu-id="7a23e-929">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-929">Requirement</span></span>| <span data-ttu-id="7a23e-930">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-930">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-931">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-931">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-932">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-932">1.0</span></span>|
|[<span data-ttu-id="7a23e-933">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-933">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-934">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-934">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-935">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-935">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-936">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-936">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a23e-937">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7a23e-937">Returns:</span></span>

<span data-ttu-id="7a23e-p162">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="7a23e-940">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="7a23e-940">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="7a23e-941">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="7a23e-941">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="7a23e-942">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7a23e-942">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-943">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-943">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7a23e-p163">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="7a23e-947">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="7a23e-947">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="7a23e-948">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-948">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="7a23e-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a23e-952">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-952">Requirements</span></span>

|<span data-ttu-id="7a23e-953">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-953">Requirement</span></span>| <span data-ttu-id="7a23e-954">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-954">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-955">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-955">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-956">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-956">1.0</span></span>|
|[<span data-ttu-id="7a23e-957">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-957">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-958">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-958">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-959">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-959">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-960">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-960">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a23e-961">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7a23e-961">Returns:</span></span>

<span data-ttu-id="7a23e-p165">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="7a23e-964">Type : Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-964">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="7a23e-965">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-965">Example</span></span>

<span data-ttu-id="7a23e-966">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="7a23e-966">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="7a23e-967">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="7a23e-967">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="7a23e-968">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7a23e-968">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-969">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7a23e-969">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7a23e-970">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="7a23e-970">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="7a23e-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-973">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-973">Parameters</span></span>

|<span data-ttu-id="7a23e-974">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-974">Name</span></span>| <span data-ttu-id="7a23e-975">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-975">Type</span></span>| <span data-ttu-id="7a23e-976">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-976">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7a23e-977">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-977">String</span></span>|<span data-ttu-id="7a23e-978">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="7a23e-978">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a23e-979">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-979">Requirements</span></span>

|<span data-ttu-id="7a23e-980">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-980">Requirement</span></span>| <span data-ttu-id="7a23e-981">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-981">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-982">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-982">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-983">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-983">1.0</span></span>|
|[<span data-ttu-id="7a23e-984">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-984">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-985">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-985">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-986">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-986">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-987">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-987">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a23e-988">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7a23e-988">Returns:</span></span>

<span data-ttu-id="7a23e-989">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7a23e-989">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="7a23e-990">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="7a23e-990">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="7a23e-991">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-991">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="7a23e-992">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="7a23e-992">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="7a23e-993">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="7a23e-993">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="7a23e-p167">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-996">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-996">Parameters</span></span>

|<span data-ttu-id="7a23e-997">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-997">Name</span></span>| <span data-ttu-id="7a23e-998">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-998">Type</span></span>| <span data-ttu-id="7a23e-999">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-999">Attributes</span></span>| <span data-ttu-id="7a23e-1000">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-1000">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="7a23e-1001">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7a23e-1001">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="7a23e-p168">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="7a23e-1005">Object</span><span class="sxs-lookup"><span data-stu-id="7a23e-1005">Object</span></span>| <span data-ttu-id="7a23e-1006">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1007">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1007">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7a23e-1008">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1008">Object</span></span>| <span data-ttu-id="7a23e-1009">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1009">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1010">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1010">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7a23e-1011">fonction</span><span class="sxs-lookup"><span data-stu-id="7a23e-1011">function</span></span>||<span data-ttu-id="7a23e-1012">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-1012">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7a23e-1013">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1013">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="7a23e-1014">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1014">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a23e-1015">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-1015">Requirements</span></span>

|<span data-ttu-id="7a23e-1016">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-1016">Requirement</span></span>| <span data-ttu-id="7a23e-1017">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-1017">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-1018">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1018">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-1019">1.2</span><span class="sxs-lookup"><span data-stu-id="7a23e-1019">1.2</span></span>|
|[<span data-ttu-id="7a23e-1020">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-1020">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-1021">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-1021">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-1022">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-1022">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-1023">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-1023">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a23e-1024">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7a23e-1024">Returns:</span></span>

<span data-ttu-id="7a23e-1025">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1025">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="7a23e-1026">Type : String</span><span class="sxs-lookup"><span data-stu-id="7a23e-1026">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7a23e-1027">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-1027">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="7a23e-1028">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7a23e-1028">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="7a23e-1029">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1029">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="7a23e-p170">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p170">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-1033">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1033">Parameters</span></span>

|<span data-ttu-id="7a23e-1034">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-1034">Name</span></span>| <span data-ttu-id="7a23e-1035">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-1035">Type</span></span>| <span data-ttu-id="7a23e-1036">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-1036">Attributes</span></span>| <span data-ttu-id="7a23e-1037">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-1037">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7a23e-1038">function</span><span class="sxs-lookup"><span data-stu-id="7a23e-1038">function</span></span>||<span data-ttu-id="7a23e-1039">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-1039">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7a23e-1040">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1040">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="7a23e-1041">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1041">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="7a23e-1042">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1042">Object</span></span>| <span data-ttu-id="7a23e-1043">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1044">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1044">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="7a23e-1045">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1045">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a23e-1046">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-1046">Requirements</span></span>

|<span data-ttu-id="7a23e-1047">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-1047">Requirement</span></span>| <span data-ttu-id="7a23e-1048">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-1048">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-1049">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1049">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-1050">1.0</span><span class="sxs-lookup"><span data-stu-id="7a23e-1050">1.0</span></span>|
|[<span data-ttu-id="7a23e-1051">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-1051">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-1052">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-1052">ReadItem</span></span>|
|[<span data-ttu-id="7a23e-1053">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-1053">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-1054">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a23e-1054">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-1055">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-1055">Example</span></span>

<span data-ttu-id="7a23e-p173">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p173">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="7a23e-1059">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7a23e-1059">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="7a23e-1060">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1060">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="7a23e-1061">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1061">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="7a23e-1062">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1062">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="7a23e-1063">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1063">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="7a23e-1064">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1064">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-1065">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1065">Parameters</span></span>

|<span data-ttu-id="7a23e-1066">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-1066">Name</span></span>| <span data-ttu-id="7a23e-1067">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-1067">Type</span></span>| <span data-ttu-id="7a23e-1068">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-1068">Attributes</span></span>| <span data-ttu-id="7a23e-1069">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-1069">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="7a23e-1070">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-1070">String</span></span>||<span data-ttu-id="7a23e-1071">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1071">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="7a23e-1072">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1072">Object</span></span>| <span data-ttu-id="7a23e-1073">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1074">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1074">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7a23e-1075">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1075">Object</span></span>| <span data-ttu-id="7a23e-1076">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1077">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1077">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7a23e-1078">fonction</span><span class="sxs-lookup"><span data-stu-id="7a23e-1078">function</span></span>| <span data-ttu-id="7a23e-1079">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1080">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-1080">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7a23e-1081">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1081">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7a23e-1082">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7a23e-1082">Errors</span></span>

| <span data-ttu-id="7a23e-1083">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="7a23e-1083">Error code</span></span> | <span data-ttu-id="7a23e-1084">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-1084">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="7a23e-1085">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1085">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a23e-1086">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-1086">Requirements</span></span>

|<span data-ttu-id="7a23e-1087">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-1087">Requirement</span></span>| <span data-ttu-id="7a23e-1088">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-1088">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-1089">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1089">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-1090">1.1</span><span class="sxs-lookup"><span data-stu-id="7a23e-1090">1.1</span></span>|
|[<span data-ttu-id="7a23e-1091">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-1091">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-1092">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-1092">ReadWriteItem</span></span>|
|[<span data-ttu-id="7a23e-1093">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-1093">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-1094">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-1094">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-1095">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-1095">Example</span></span>

<span data-ttu-id="7a23e-1096">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="7a23e-1096">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="7a23e-1097">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="7a23e-1097">saveAsync([options], callback)</span></span>

<span data-ttu-id="7a23e-1098">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1098">Asynchronously saves an item.</span></span>

<span data-ttu-id="7a23e-1099">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1099">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="7a23e-1100">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1100">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="7a23e-1101">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1101">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-1102">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1102">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="7a23e-1103">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1103">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="7a23e-p177">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p177">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="7a23e-1107">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="7a23e-1107">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="7a23e-1108">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1108">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="7a23e-1109">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1109">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="7a23e-1110">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="7a23e-1110">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="7a23e-1111">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1111">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-1112">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1112">Parameters</span></span>

|<span data-ttu-id="7a23e-1113">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-1113">Name</span></span>| <span data-ttu-id="7a23e-1114">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-1114">Type</span></span>| <span data-ttu-id="7a23e-1115">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-1115">Attributes</span></span>| <span data-ttu-id="7a23e-1116">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-1116">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="7a23e-1117">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1117">Object</span></span>| <span data-ttu-id="7a23e-1118">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1119">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1119">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7a23e-1120">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1120">Object</span></span>| <span data-ttu-id="7a23e-1121">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1122">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1122">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7a23e-1123">fonction</span><span class="sxs-lookup"><span data-stu-id="7a23e-1123">function</span></span>||<span data-ttu-id="7a23e-1124">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-1124">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7a23e-1125">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1125">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a23e-1126">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-1126">Requirements</span></span>

|<span data-ttu-id="7a23e-1127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-1127">Requirement</span></span>| <span data-ttu-id="7a23e-1128">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-1128">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-1129">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-1130">1.3</span><span class="sxs-lookup"><span data-stu-id="7a23e-1130">1.3</span></span>|
|[<span data-ttu-id="7a23e-1131">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-1131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-1132">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-1132">ReadWriteItem</span></span>|
|[<span data-ttu-id="7a23e-1133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-1133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-1134">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-1134">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="7a23e-1135">範例</span><span class="sxs-lookup"><span data-stu-id="7a23e-1135">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="7a23e-p179">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p179">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="7a23e-1138">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="7a23e-1138">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="7a23e-1139">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1139">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="7a23e-p180">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p180">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a23e-1143">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1143">Parameters</span></span>

|<span data-ttu-id="7a23e-1144">Nom</span><span class="sxs-lookup"><span data-stu-id="7a23e-1144">Name</span></span>| <span data-ttu-id="7a23e-1145">Type</span><span class="sxs-lookup"><span data-stu-id="7a23e-1145">Type</span></span>| <span data-ttu-id="7a23e-1146">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a23e-1146">Attributes</span></span>| <span data-ttu-id="7a23e-1147">Description</span><span class="sxs-lookup"><span data-stu-id="7a23e-1147">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7a23e-1148">String</span><span class="sxs-lookup"><span data-stu-id="7a23e-1148">String</span></span>||<span data-ttu-id="7a23e-p181">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="7a23e-p181">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="7a23e-1152">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1152">Object</span></span>| <span data-ttu-id="7a23e-1153">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1153">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1154">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1154">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7a23e-1155">Objet</span><span class="sxs-lookup"><span data-stu-id="7a23e-1155">Object</span></span>| <span data-ttu-id="7a23e-1156">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1157">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1157">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="7a23e-1158">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7a23e-1158">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="7a23e-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a23e-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="7a23e-1160">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1160">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="7a23e-1161">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1161">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="7a23e-1162">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1162">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="7a23e-1163">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1163">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="7a23e-1164">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="7a23e-1164">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="7a23e-1165">fonction</span><span class="sxs-lookup"><span data-stu-id="7a23e-1165">function</span></span>||<span data-ttu-id="7a23e-1166">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a23e-1166">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a23e-1167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a23e-1167">Requirements</span></span>

|<span data-ttu-id="7a23e-1168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a23e-1168">Requirement</span></span>| <span data-ttu-id="7a23e-1169">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a23e-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a23e-1170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a23e-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a23e-1171">1.2</span><span class="sxs-lookup"><span data-stu-id="7a23e-1171">1.2</span></span>|
|[<span data-ttu-id="7a23e-1172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a23e-1172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a23e-1173">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7a23e-1173">ReadWriteItem</span></span>|
|[<span data-ttu-id="7a23e-1174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a23e-1174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a23e-1175">Composition</span><span class="sxs-lookup"><span data-stu-id="7a23e-1175">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7a23e-1176">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a23e-1176">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
