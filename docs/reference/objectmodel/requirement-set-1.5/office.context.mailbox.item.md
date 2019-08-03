---
title: Office.context.mailbox.item - ensemble de conditions requises 1.5
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: cb841795769e19b76bda49c4a9b704734130eadc
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064635"
---
# <a name="item"></a><span data-ttu-id="43f38-102">élément</span><span class="sxs-lookup"><span data-stu-id="43f38-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="43f38-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="43f38-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="43f38-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="43f38-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-106">Requirements</span></span>

|<span data-ttu-id="43f38-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-107">Requirement</span></span>| <span data-ttu-id="43f38-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-110">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-110">1.0</span></span>|
|[<span data-ttu-id="43f38-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="43f38-112">Restricted</span></span>|
|[<span data-ttu-id="43f38-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="43f38-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="43f38-115">Members and methods</span></span>

| <span data-ttu-id="43f38-116">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-116">Member</span></span> | <span data-ttu-id="43f38-117">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="43f38-118">attachments</span><span class="sxs-lookup"><span data-stu-id="43f38-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="43f38-119">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-119">Member</span></span> |
| [<span data-ttu-id="43f38-120">bcc</span><span class="sxs-lookup"><span data-stu-id="43f38-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="43f38-121">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-121">Member</span></span> |
| [<span data-ttu-id="43f38-122">body</span><span class="sxs-lookup"><span data-stu-id="43f38-122">body</span></span>](#body-body) | <span data-ttu-id="43f38-123">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-123">Member</span></span> |
| [<span data-ttu-id="43f38-124">cc</span><span class="sxs-lookup"><span data-stu-id="43f38-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="43f38-125">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-125">Member</span></span> |
| [<span data-ttu-id="43f38-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="43f38-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="43f38-127">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-127">Member</span></span> |
| [<span data-ttu-id="43f38-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="43f38-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="43f38-129">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-129">Member</span></span> |
| [<span data-ttu-id="43f38-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="43f38-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="43f38-131">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-131">Member</span></span> |
| [<span data-ttu-id="43f38-132">end</span><span class="sxs-lookup"><span data-stu-id="43f38-132">end</span></span>](#end-datetime) | <span data-ttu-id="43f38-133">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-133">Member</span></span> |
| [<span data-ttu-id="43f38-134">from</span><span class="sxs-lookup"><span data-stu-id="43f38-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="43f38-135">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-135">Member</span></span> |
| [<span data-ttu-id="43f38-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="43f38-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="43f38-137">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-137">Member</span></span> |
| [<span data-ttu-id="43f38-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="43f38-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="43f38-139">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-139">Member</span></span> |
| [<span data-ttu-id="43f38-140">itemId</span><span class="sxs-lookup"><span data-stu-id="43f38-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="43f38-141">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-141">Member</span></span> |
| [<span data-ttu-id="43f38-142">itemType</span><span class="sxs-lookup"><span data-stu-id="43f38-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="43f38-143">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-143">Member</span></span> |
| [<span data-ttu-id="43f38-144">location</span><span class="sxs-lookup"><span data-stu-id="43f38-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="43f38-145">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-145">Member</span></span> |
| [<span data-ttu-id="43f38-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="43f38-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="43f38-147">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-147">Member</span></span> |
| [<span data-ttu-id="43f38-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="43f38-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="43f38-149">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-149">Member</span></span> |
| [<span data-ttu-id="43f38-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="43f38-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="43f38-151">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-151">Member</span></span> |
| [<span data-ttu-id="43f38-152">organizer</span><span class="sxs-lookup"><span data-stu-id="43f38-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="43f38-153">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-153">Member</span></span> |
| [<span data-ttu-id="43f38-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="43f38-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="43f38-155">Member</span><span class="sxs-lookup"><span data-stu-id="43f38-155">Member</span></span> |
| [<span data-ttu-id="43f38-156">sender</span><span class="sxs-lookup"><span data-stu-id="43f38-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="43f38-157">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-157">Member</span></span> |
| [<span data-ttu-id="43f38-158">start</span><span class="sxs-lookup"><span data-stu-id="43f38-158">start</span></span>](#start-datetime) | <span data-ttu-id="43f38-159">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-159">Member</span></span> |
| [<span data-ttu-id="43f38-160">subject</span><span class="sxs-lookup"><span data-stu-id="43f38-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="43f38-161">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-161">Member</span></span> |
| [<span data-ttu-id="43f38-162">to</span><span class="sxs-lookup"><span data-stu-id="43f38-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="43f38-163">Membre</span><span class="sxs-lookup"><span data-stu-id="43f38-163">Member</span></span> |
| [<span data-ttu-id="43f38-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="43f38-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-165">Method</span></span> |
| [<span data-ttu-id="43f38-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="43f38-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-167">Method</span></span> |
| [<span data-ttu-id="43f38-168">close</span><span class="sxs-lookup"><span data-stu-id="43f38-168">close</span></span>](#close) | <span data-ttu-id="43f38-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-169">Method</span></span> |
| [<span data-ttu-id="43f38-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="43f38-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="43f38-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-171">Method</span></span> |
| [<span data-ttu-id="43f38-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="43f38-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="43f38-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-173">Method</span></span> |
| [<span data-ttu-id="43f38-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="43f38-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="43f38-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-175">Method</span></span> |
| [<span data-ttu-id="43f38-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="43f38-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="43f38-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-177">Method</span></span> |
| [<span data-ttu-id="43f38-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="43f38-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="43f38-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-179">Method</span></span> |
| [<span data-ttu-id="43f38-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="43f38-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="43f38-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-181">Method</span></span> |
| [<span data-ttu-id="43f38-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="43f38-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="43f38-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-183">Method</span></span> |
| [<span data-ttu-id="43f38-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="43f38-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-185">Method</span></span> |
| [<span data-ttu-id="43f38-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="43f38-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-187">Method</span></span> |
| [<span data-ttu-id="43f38-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="43f38-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-189">Method</span></span> |
| [<span data-ttu-id="43f38-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="43f38-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-191">Method</span></span> |
| [<span data-ttu-id="43f38-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="43f38-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="43f38-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="43f38-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="43f38-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-194">Example</span></span>

<span data-ttu-id="43f38-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="43f38-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
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

### <a name="members"></a><span data-ttu-id="43f38-196">Members</span><span class="sxs-lookup"><span data-stu-id="43f38-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="43f38-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="43f38-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="43f38-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="43f38-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="43f38-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="43f38-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-202">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-202">Type</span></span>

*   <span data-ttu-id="43f38-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="43f38-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-204">Requirements</span></span>

|<span data-ttu-id="43f38-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-205">Requirement</span></span>| <span data-ttu-id="43f38-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-208">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-208">1.0</span></span>|
|[<span data-ttu-id="43f38-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-210">ReadItem</span></span>|
|[<span data-ttu-id="43f38-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-213">Example</span></span>

<span data-ttu-id="43f38-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="43f38-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-215">bcc :[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="43f38-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="43f38-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-218">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-218">Type</span></span>

*   [<span data-ttu-id="43f38-219">Destinataires</span><span class="sxs-lookup"><span data-stu-id="43f38-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-220">Requirements</span></span>

|<span data-ttu-id="43f38-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-221">Requirement</span></span>| <span data-ttu-id="43f38-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-224">1.1</span><span class="sxs-lookup"><span data-stu-id="43f38-224">1.1</span></span>|
|[<span data-ttu-id="43f38-225">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-226">ReadItem</span></span>|
|[<span data-ttu-id="43f38-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-228">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="43f38-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-230">body :[Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-231">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-232">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-232">Type</span></span>

*   [<span data-ttu-id="43f38-233">Body</span><span class="sxs-lookup"><span data-stu-id="43f38-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-234">Requirements</span></span>

|<span data-ttu-id="43f38-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-235">Requirement</span></span>| <span data-ttu-id="43f38-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-238">1.1</span><span class="sxs-lookup"><span data-stu-id="43f38-238">1.1</span></span>|
|[<span data-ttu-id="43f38-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-240">ReadItem</span></span>|
|[<span data-ttu-id="43f38-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-242">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-243">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-243">Example</span></span>

<span data-ttu-id="43f38-244">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="43f38-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="43f38-245">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="43f38-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-247">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="43f38-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="43f38-248">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-249">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-249">Read mode</span></span>

<span data-ttu-id="43f38-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="43f38-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-252">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-252">Compose mode</span></span>

<span data-ttu-id="43f38-253">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="43f38-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43f38-254">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-254">Type</span></span>

*   <span data-ttu-id="43f38-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-256">Requirements</span></span>

|<span data-ttu-id="43f38-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-257">Requirement</span></span>| <span data-ttu-id="43f38-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-260">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-260">1.0</span></span>|
|[<span data-ttu-id="43f38-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-262">ReadItem</span></span>|
|[<span data-ttu-id="43f38-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-264">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="43f38-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="43f38-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="43f38-266">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="43f38-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="43f38-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="43f38-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="43f38-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="43f38-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-271">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-271">Type</span></span>

*   <span data-ttu-id="43f38-272">String</span><span class="sxs-lookup"><span data-stu-id="43f38-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-273">Requirements</span></span>

|<span data-ttu-id="43f38-274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-274">Requirement</span></span>| <span data-ttu-id="43f38-275">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-277">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-277">1.0</span></span>|
|[<span data-ttu-id="43f38-278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-279">ReadItem</span></span>|
|[<span data-ttu-id="43f38-280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-281">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-282">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="43f38-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="43f38-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="43f38-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-286">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-286">Type</span></span>

*   <span data-ttu-id="43f38-287">Date</span><span class="sxs-lookup"><span data-stu-id="43f38-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-288">Requirements</span></span>

|<span data-ttu-id="43f38-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-289">Requirement</span></span>| <span data-ttu-id="43f38-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-292">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-292">1.0</span></span>|
|[<span data-ttu-id="43f38-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-294">ReadItem</span></span>|
|[<span data-ttu-id="43f38-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-296">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="43f38-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="43f38-298">dateTimeModified :Date</span></span>

<span data-ttu-id="43f38-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-301">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-302">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-302">Type</span></span>

*   <span data-ttu-id="43f38-303">Date</span><span class="sxs-lookup"><span data-stu-id="43f38-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-304">Requirements</span></span>

|<span data-ttu-id="43f38-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-305">Requirement</span></span>| <span data-ttu-id="43f38-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-308">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-308">1.0</span></span>|
|[<span data-ttu-id="43f38-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-310">ReadItem</span></span>|
|[<span data-ttu-id="43f38-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="43f38-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-314">end :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-315">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="43f38-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="43f38-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-318">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-318">Read mode</span></span>

<span data-ttu-id="43f38-319">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="43f38-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-320">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-320">Compose mode</span></span>

<span data-ttu-id="43f38-321">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="43f38-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="43f38-322">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="43f38-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="43f38-323">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="43f38-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="43f38-324">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-324">Type</span></span>

*   <span data-ttu-id="43f38-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-326">Requirements</span></span>

|<span data-ttu-id="43f38-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-327">Requirement</span></span>| <span data-ttu-id="43f38-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-330">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-330">1.0</span></span>|
|[<span data-ttu-id="43f38-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-332">ReadItem</span></span>|
|[<span data-ttu-id="43f38-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-334">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="43f38-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-335">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="43f38-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="43f38-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-340">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="43f38-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-341">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-341">Type</span></span>

*   [<span data-ttu-id="43f38-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="43f38-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-343">Requirements</span></span>

|<span data-ttu-id="43f38-344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-344">Requirement</span></span>| <span data-ttu-id="43f38-345">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-347">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-347">1.0</span></span>|
|[<span data-ttu-id="43f38-348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-349">ReadItem</span></span>|
|[<span data-ttu-id="43f38-350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-351">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-352">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="43f38-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="43f38-353">internetMessageId :String</span></span>

<span data-ttu-id="43f38-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-356">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-356">Type</span></span>

*   <span data-ttu-id="43f38-357">String</span><span class="sxs-lookup"><span data-stu-id="43f38-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-358">Requirements</span></span>

|<span data-ttu-id="43f38-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-359">Requirement</span></span>| <span data-ttu-id="43f38-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-362">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-362">1.0</span></span>|
|[<span data-ttu-id="43f38-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-364">ReadItem</span></span>|
|[<span data-ttu-id="43f38-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="43f38-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="43f38-368">itemClass :String</span></span>

<span data-ttu-id="43f38-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="43f38-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="43f38-373">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-373">Type</span></span> | <span data-ttu-id="43f38-374">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-374">Description</span></span> | <span data-ttu-id="43f38-375">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="43f38-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="43f38-376">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="43f38-376">Appointment items</span></span> | <span data-ttu-id="43f38-377">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="43f38-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="43f38-378">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="43f38-378">Message items</span></span> | <span data-ttu-id="43f38-379">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="43f38-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="43f38-380">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="43f38-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-381">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-381">Type</span></span>

*   <span data-ttu-id="43f38-382">String</span><span class="sxs-lookup"><span data-stu-id="43f38-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-383">Requirements</span></span>

|<span data-ttu-id="43f38-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-384">Requirement</span></span>| <span data-ttu-id="43f38-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-387">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-387">1.0</span></span>|
|[<span data-ttu-id="43f38-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-389">ReadItem</span></span>|
|[<span data-ttu-id="43f38-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="43f38-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="43f38-393">(nullable) itemId :String</span></span>

<span data-ttu-id="43f38-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-396">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="43f38-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="43f38-397">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="43f38-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="43f38-398">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="43f38-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="43f38-399">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="43f38-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="43f38-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-402">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-402">Type</span></span>

*   <span data-ttu-id="43f38-403">String</span><span class="sxs-lookup"><span data-stu-id="43f38-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-404">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-404">Requirements</span></span>

|<span data-ttu-id="43f38-405">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-405">Requirement</span></span>| <span data-ttu-id="43f38-406">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-407">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-408">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-408">1.0</span></span>|
|[<span data-ttu-id="43f38-409">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-410">ReadItem</span></span>|
|[<span data-ttu-id="43f38-411">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-412">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-413">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-413">Example</span></span>

<span data-ttu-id="43f38-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="43f38-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="43f38-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-417">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="43f38-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="43f38-418">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-419">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-419">Type</span></span>

*   [<span data-ttu-id="43f38-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="43f38-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-421">Requirements</span></span>

|<span data-ttu-id="43f38-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-422">Requirement</span></span>| <span data-ttu-id="43f38-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-425">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-425">1.0</span></span>|
|[<span data-ttu-id="43f38-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-427">ReadItem</span></span>|
|[<span data-ttu-id="43f38-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-429">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="43f38-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-431">location :String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-432">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-433">Read mode</span></span>

<span data-ttu-id="43f38-434">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-435">Compose mode</span></span>

<span data-ttu-id="43f38-436">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43f38-437">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-437">Type</span></span>

*   <span data-ttu-id="43f38-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-439">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-439">Requirements</span></span>

|<span data-ttu-id="43f38-440">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-440">Requirement</span></span>| <span data-ttu-id="43f38-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-442">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-443">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-443">1.0</span></span>|
|[<span data-ttu-id="43f38-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-445">ReadItem</span></span>|
|[<span data-ttu-id="43f38-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-447">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="43f38-448">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="43f38-448">normalizedSubject :String</span></span>

<span data-ttu-id="43f38-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="43f38-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="43f38-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-453">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-453">Type</span></span>

*   <span data-ttu-id="43f38-454">String</span><span class="sxs-lookup"><span data-stu-id="43f38-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-455">Requirements</span></span>

|<span data-ttu-id="43f38-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-456">Requirement</span></span>| <span data-ttu-id="43f38-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-459">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-459">1.0</span></span>|
|[<span data-ttu-id="43f38-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-461">ReadItem</span></span>|
|[<span data-ttu-id="43f38-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-463">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-464">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="43f38-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-465">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-466">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-467">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-467">Type</span></span>

*   [<span data-ttu-id="43f38-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="43f38-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-469">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-469">Requirements</span></span>

|<span data-ttu-id="43f38-470">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-470">Requirement</span></span>| <span data-ttu-id="43f38-471">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-472">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-473">1.3</span><span class="sxs-lookup"><span data-stu-id="43f38-473">1.3</span></span>|
|[<span data-ttu-id="43f38-474">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-475">ReadItem</span></span>|
|[<span data-ttu-id="43f38-476">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-477">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-478">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="43f38-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-480">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="43f38-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="43f38-481">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-482">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-482">Read mode</span></span>

<span data-ttu-id="43f38-483">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="43f38-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-484">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-484">Compose mode</span></span>

<span data-ttu-id="43f38-485">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="43f38-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43f38-486">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-486">Type</span></span>

*   <span data-ttu-id="43f38-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-488">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-488">Requirements</span></span>

|<span data-ttu-id="43f38-489">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-489">Requirement</span></span>| <span data-ttu-id="43f38-490">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-491">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-492">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-492">1.0</span></span>|
|[<span data-ttu-id="43f38-493">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-494">ReadItem</span></span>|
|[<span data-ttu-id="43f38-495">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-496">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="43f38-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-497">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-500">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-500">Type</span></span>

*   [<span data-ttu-id="43f38-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="43f38-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-502">Requirements</span></span>

|<span data-ttu-id="43f38-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-503">Requirement</span></span>| <span data-ttu-id="43f38-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-506">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-506">1.0</span></span>|
|[<span data-ttu-id="43f38-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-508">ReadItem</span></span>|
|[<span data-ttu-id="43f38-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-510">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-511">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="43f38-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-513">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="43f38-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="43f38-514">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-515">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-515">Read mode</span></span>

<span data-ttu-id="43f38-516">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="43f38-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-517">Compose mode</span></span>

<span data-ttu-id="43f38-518">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="43f38-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="43f38-519">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-519">Type</span></span>

*   <span data-ttu-id="43f38-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-521">Requirements</span></span>

|<span data-ttu-id="43f38-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-522">Requirement</span></span>| <span data-ttu-id="43f38-523">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-525">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-525">1.0</span></span>|
|[<span data-ttu-id="43f38-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-527">ReadItem</span></span>|
|[<span data-ttu-id="43f38-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-529">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="43f38-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-530">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="43f38-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="43f38-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="43f38-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-535">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="43f38-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="43f38-536">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-536">Type</span></span>

*   [<span data-ttu-id="43f38-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="43f38-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="43f38-538">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-538">Requirements</span></span>

|<span data-ttu-id="43f38-539">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-539">Requirement</span></span>| <span data-ttu-id="43f38-540">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-541">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-542">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-542">1.0</span></span>|
|[<span data-ttu-id="43f38-543">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-544">ReadItem</span></span>|
|[<span data-ttu-id="43f38-545">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-546">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-547">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="43f38-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-548">start :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-549">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="43f38-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="43f38-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-552">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-552">Read mode</span></span>

<span data-ttu-id="43f38-553">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="43f38-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-554">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-554">Compose mode</span></span>

<span data-ttu-id="43f38-555">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="43f38-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="43f38-556">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="43f38-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="43f38-557">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="43f38-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="43f38-558">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-558">Type</span></span>

*   <span data-ttu-id="43f38-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-560">Requirements</span></span>

|<span data-ttu-id="43f38-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-561">Requirement</span></span>| <span data-ttu-id="43f38-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-564">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-564">1.0</span></span>|
|[<span data-ttu-id="43f38-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-566">ReadItem</span></span>|
|[<span data-ttu-id="43f38-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-568">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="43f38-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-569">subject :String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="43f38-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="43f38-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-572">Read mode</span></span>

<span data-ttu-id="43f38-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="43f38-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-575">Compose mode</span></span>

<span data-ttu-id="43f38-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="43f38-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="43f38-577">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-577">Type</span></span>

*   <span data-ttu-id="43f38-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-579">Requirements</span></span>

|<span data-ttu-id="43f38-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-580">Requirement</span></span>| <span data-ttu-id="43f38-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-583">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-583">1.0</span></span>|
|[<span data-ttu-id="43f38-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-585">ReadItem</span></span>|
|[<span data-ttu-id="43f38-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-587">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-587">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="43f38-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="43f38-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="43f38-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="43f38-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43f38-591">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-591">Read mode</span></span>

<span data-ttu-id="43f38-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="43f38-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="43f38-594">Mode composition</span><span class="sxs-lookup"><span data-stu-id="43f38-594">Compose mode</span></span>

<span data-ttu-id="43f38-595">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="43f38-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43f38-596">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-596">Type</span></span>

*   <span data-ttu-id="43f38-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-598">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-598">Requirements</span></span>

|<span data-ttu-id="43f38-599">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-599">Requirement</span></span>| <span data-ttu-id="43f38-600">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-601">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-602">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-602">1.0</span></span>|
|[<span data-ttu-id="43f38-603">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-604">ReadItem</span></span>|
|[<span data-ttu-id="43f38-605">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-606">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="43f38-607">Méthodes</span><span class="sxs-lookup"><span data-stu-id="43f38-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="43f38-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43f38-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="43f38-609">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="43f38-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="43f38-610">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="43f38-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="43f38-611">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="43f38-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-612">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-612">Parameters</span></span>

|<span data-ttu-id="43f38-613">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-613">Name</span></span>| <span data-ttu-id="43f38-614">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-614">Type</span></span>| <span data-ttu-id="43f38-615">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-615">Attributes</span></span>| <span data-ttu-id="43f38-616">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="43f38-617">String</span><span class="sxs-lookup"><span data-stu-id="43f38-617">String</span></span>||<span data-ttu-id="43f38-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="43f38-620">String</span><span class="sxs-lookup"><span data-stu-id="43f38-620">String</span></span>||<span data-ttu-id="43f38-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="43f38-623">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-623">Object</span></span>| <span data-ttu-id="43f38-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-624">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="43f38-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="43f38-626">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-626">Object</span></span> | <span data-ttu-id="43f38-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-627">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="43f38-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="43f38-629">Boolean</span></span> | <span data-ttu-id="43f38-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-630">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-631">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="43f38-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="43f38-632">fonction</span><span class="sxs-lookup"><span data-stu-id="43f38-632">function</span></span>| <span data-ttu-id="43f38-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-633">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-634">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="43f38-635">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43f38-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="43f38-636">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="43f38-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43f38-637">Erreurs</span><span class="sxs-lookup"><span data-stu-id="43f38-637">Errors</span></span>

| <span data-ttu-id="43f38-638">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="43f38-638">Error code</span></span> | <span data-ttu-id="43f38-639">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="43f38-640">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="43f38-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="43f38-641">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="43f38-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="43f38-642">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="43f38-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f38-643">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-643">Requirements</span></span>

|<span data-ttu-id="43f38-644">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-644">Requirement</span></span>| <span data-ttu-id="43f38-645">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-646">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-647">1.1</span><span class="sxs-lookup"><span data-stu-id="43f38-647">1.1</span></span>|
|[<span data-ttu-id="43f38-648">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43f38-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="43f38-650">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-651">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="43f38-652">Exemples</span><span class="sxs-lookup"><span data-stu-id="43f38-652">Examples</span></span>

```javascript
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

<span data-ttu-id="43f38-653">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="43f38-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="43f38-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43f38-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="43f38-655">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="43f38-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="43f38-659">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="43f38-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="43f38-660">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="43f38-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-661">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-661">Parameters</span></span>

|<span data-ttu-id="43f38-662">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-662">Name</span></span>| <span data-ttu-id="43f38-663">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-663">Type</span></span>| <span data-ttu-id="43f38-664">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-664">Attributes</span></span>| <span data-ttu-id="43f38-665">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="43f38-666">String</span><span class="sxs-lookup"><span data-stu-id="43f38-666">String</span></span>||<span data-ttu-id="43f38-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="43f38-669">String</span><span class="sxs-lookup"><span data-stu-id="43f38-669">String</span></span>||<span data-ttu-id="43f38-670">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="43f38-670">The subject of the item to be attached.</span></span> <span data-ttu-id="43f38-671">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="43f38-672">Object</span><span class="sxs-lookup"><span data-stu-id="43f38-672">Object</span></span>| <span data-ttu-id="43f38-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-673">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-674">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="43f38-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43f38-675">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-675">Object</span></span>| <span data-ttu-id="43f38-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-676">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-677">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43f38-678">fonction</span><span class="sxs-lookup"><span data-stu-id="43f38-678">function</span></span>| <span data-ttu-id="43f38-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-679">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-680">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="43f38-681">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43f38-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="43f38-682">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="43f38-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43f38-683">Erreurs</span><span class="sxs-lookup"><span data-stu-id="43f38-683">Errors</span></span>

| <span data-ttu-id="43f38-684">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="43f38-684">Error code</span></span> | <span data-ttu-id="43f38-685">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="43f38-686">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="43f38-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f38-687">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-687">Requirements</span></span>

|<span data-ttu-id="43f38-688">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-688">Requirement</span></span>| <span data-ttu-id="43f38-689">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-690">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-691">1.1</span><span class="sxs-lookup"><span data-stu-id="43f38-691">1.1</span></span>|
|[<span data-ttu-id="43f38-692">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43f38-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="43f38-694">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-695">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-696">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-696">Example</span></span>

<span data-ttu-id="43f38-697">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="43f38-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

#### <a name="close"></a><span data-ttu-id="43f38-698">close()</span><span class="sxs-lookup"><span data-stu-id="43f38-698">close()</span></span>

<span data-ttu-id="43f38-699">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="43f38-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="43f38-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="43f38-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-702">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="43f38-703">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="43f38-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-704">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-704">Requirements</span></span>

|<span data-ttu-id="43f38-705">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-705">Requirement</span></span>| <span data-ttu-id="43f38-706">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-707">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-708">1.3</span><span class="sxs-lookup"><span data-stu-id="43f38-708">1.3</span></span>|
|[<span data-ttu-id="43f38-709">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-710">Restreinte</span><span class="sxs-lookup"><span data-stu-id="43f38-710">Restricted</span></span>|
|[<span data-ttu-id="43f38-711">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-712">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="43f38-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="43f38-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="43f38-714">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="43f38-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-715">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f38-716">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="43f38-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="43f38-717">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="43f38-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="43f38-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="43f38-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-721">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-721">Parameters</span></span>

| <span data-ttu-id="43f38-722">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-722">Name</span></span> | <span data-ttu-id="43f38-723">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-723">Type</span></span> | <span data-ttu-id="43f38-724">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-724">Attributes</span></span> | <span data-ttu-id="43f38-725">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="43f38-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="43f38-726">String &#124; Object</span></span>| |<span data-ttu-id="43f38-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="43f38-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="43f38-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="43f38-729">**OR**</span></span><br/><span data-ttu-id="43f38-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="43f38-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="43f38-732">String</span><span class="sxs-lookup"><span data-stu-id="43f38-732">String</span></span> | <span data-ttu-id="43f38-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-733">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="43f38-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="43f38-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="43f38-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-737">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-738">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="43f38-739">Chaîne</span><span class="sxs-lookup"><span data-stu-id="43f38-739">String</span></span> | | <span data-ttu-id="43f38-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="43f38-742">String</span><span class="sxs-lookup"><span data-stu-id="43f38-742">String</span></span> | | <span data-ttu-id="43f38-743">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="43f38-744">Chaîne</span><span class="sxs-lookup"><span data-stu-id="43f38-744">String</span></span> | | <span data-ttu-id="43f38-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="43f38-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="43f38-747">Booléen</span><span class="sxs-lookup"><span data-stu-id="43f38-747">Boolean</span></span> | | <span data-ttu-id="43f38-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="43f38-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="43f38-750">String</span><span class="sxs-lookup"><span data-stu-id="43f38-750">String</span></span> | | <span data-ttu-id="43f38-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="43f38-754">function</span><span class="sxs-lookup"><span data-stu-id="43f38-754">function</span></span> | <span data-ttu-id="43f38-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-755">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-756">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f38-757">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-757">Requirements</span></span>

|<span data-ttu-id="43f38-758">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-758">Requirement</span></span>| <span data-ttu-id="43f38-759">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-760">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-761">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-761">1.0</span></span>|
|[<span data-ttu-id="43f38-762">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-763">ReadItem</span></span>|
|[<span data-ttu-id="43f38-764">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-765">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="43f38-766">Exemples</span><span class="sxs-lookup"><span data-stu-id="43f38-766">Examples</span></span>

<span data-ttu-id="43f38-767">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="43f38-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="43f38-768">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="43f38-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="43f38-769">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="43f38-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="43f38-770">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="43f38-770">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="43f38-771">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-771">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="43f38-772">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="43f38-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="43f38-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="43f38-774">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="43f38-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-775">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f38-776">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="43f38-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="43f38-777">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="43f38-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="43f38-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="43f38-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-781">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-781">Parameters</span></span>

| <span data-ttu-id="43f38-782">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-782">Name</span></span> | <span data-ttu-id="43f38-783">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-783">Type</span></span> | <span data-ttu-id="43f38-784">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-784">Attributes</span></span> | <span data-ttu-id="43f38-785">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="43f38-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="43f38-786">String &#124; Object</span></span>| | <span data-ttu-id="43f38-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="43f38-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="43f38-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="43f38-789">**OR**</span></span><br/><span data-ttu-id="43f38-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="43f38-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="43f38-792">String</span><span class="sxs-lookup"><span data-stu-id="43f38-792">String</span></span> | <span data-ttu-id="43f38-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-793">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="43f38-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="43f38-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="43f38-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-797">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-798">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="43f38-799">Chaîne</span><span class="sxs-lookup"><span data-stu-id="43f38-799">String</span></span> | | <span data-ttu-id="43f38-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="43f38-802">String</span><span class="sxs-lookup"><span data-stu-id="43f38-802">String</span></span> | | <span data-ttu-id="43f38-803">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="43f38-804">Chaîne</span><span class="sxs-lookup"><span data-stu-id="43f38-804">String</span></span> | | <span data-ttu-id="43f38-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="43f38-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="43f38-807">Booléen</span><span class="sxs-lookup"><span data-stu-id="43f38-807">Boolean</span></span> | | <span data-ttu-id="43f38-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="43f38-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="43f38-810">String</span><span class="sxs-lookup"><span data-stu-id="43f38-810">String</span></span> | | <span data-ttu-id="43f38-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="43f38-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="43f38-814">function</span><span class="sxs-lookup"><span data-stu-id="43f38-814">function</span></span> | <span data-ttu-id="43f38-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-815">&lt;optional&gt;</span></span> | <span data-ttu-id="43f38-816">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f38-817">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-817">Requirements</span></span>

|<span data-ttu-id="43f38-818">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-818">Requirement</span></span>| <span data-ttu-id="43f38-819">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-820">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-821">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-821">1.0</span></span>|
|[<span data-ttu-id="43f38-822">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-823">ReadItem</span></span>|
|[<span data-ttu-id="43f38-824">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-825">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="43f38-826">Exemples</span><span class="sxs-lookup"><span data-stu-id="43f38-826">Examples</span></span>

<span data-ttu-id="43f38-827">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="43f38-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="43f38-828">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="43f38-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="43f38-829">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="43f38-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="43f38-830">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="43f38-830">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="43f38-831">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-831">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="43f38-832">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="43f38-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="43f38-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="43f38-834">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="43f38-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-835">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-836">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-836">Requirements</span></span>

|<span data-ttu-id="43f38-837">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-837">Requirement</span></span>| <span data-ttu-id="43f38-838">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-839">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-840">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-840">1.0</span></span>|
|[<span data-ttu-id="43f38-841">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-842">ReadItem</span></span>|
|[<span data-ttu-id="43f38-843">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-844">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f38-845">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="43f38-845">Returns:</span></span>

<span data-ttu-id="43f38-846">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="43f38-846">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="43f38-847">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-847">Example</span></span>

<span data-ttu-id="43f38-848">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="43f38-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="43f38-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="43f38-850">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="43f38-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-851">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-852">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-852">Parameters</span></span>

|<span data-ttu-id="43f38-853">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-853">Name</span></span>| <span data-ttu-id="43f38-854">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-854">Type</span></span>| <span data-ttu-id="43f38-855">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="43f38-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="43f38-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="43f38-857">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="43f38-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f38-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-858">Requirements</span></span>

|<span data-ttu-id="43f38-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-859">Requirement</span></span>| <span data-ttu-id="43f38-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-862">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-862">1.0</span></span>|
|[<span data-ttu-id="43f38-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-864">Restreinte</span><span class="sxs-lookup"><span data-stu-id="43f38-864">Restricted</span></span>|
|[<span data-ttu-id="43f38-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f38-867">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="43f38-867">Returns:</span></span>

<span data-ttu-id="43f38-868">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="43f38-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="43f38-869">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="43f38-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="43f38-870">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="43f38-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="43f38-871">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="43f38-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="43f38-872">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="43f38-872">Value of `entityType`</span></span> | <span data-ttu-id="43f38-873">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="43f38-873">Type of objects in returned array</span></span> | <span data-ttu-id="43f38-874">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="43f38-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="43f38-875">String</span><span class="sxs-lookup"><span data-stu-id="43f38-875">String</span></span> | <span data-ttu-id="43f38-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="43f38-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="43f38-877">Contact</span><span class="sxs-lookup"><span data-stu-id="43f38-877">Contact</span></span> | <span data-ttu-id="43f38-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43f38-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="43f38-879">String</span><span class="sxs-lookup"><span data-stu-id="43f38-879">String</span></span> | <span data-ttu-id="43f38-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43f38-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="43f38-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="43f38-881">MeetingSuggestion</span></span> | <span data-ttu-id="43f38-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43f38-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="43f38-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="43f38-883">PhoneNumber</span></span> | <span data-ttu-id="43f38-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="43f38-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="43f38-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="43f38-885">TaskSuggestion</span></span> | <span data-ttu-id="43f38-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43f38-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="43f38-887">String</span><span class="sxs-lookup"><span data-stu-id="43f38-887">String</span></span> | <span data-ttu-id="43f38-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="43f38-888">**Restricted**</span></span> |

<span data-ttu-id="43f38-889">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="43f38-889">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="43f38-890">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-890">Example</span></span>

<span data-ttu-id="43f38-891">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="43f38-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficeviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="43f38-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="43f38-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="43f38-893">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="43f38-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-894">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f38-895">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="43f38-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-896">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-896">Parameters</span></span>

|<span data-ttu-id="43f38-897">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-897">Name</span></span>| <span data-ttu-id="43f38-898">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-898">Type</span></span>| <span data-ttu-id="43f38-899">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="43f38-900">String</span><span class="sxs-lookup"><span data-stu-id="43f38-900">String</span></span>|<span data-ttu-id="43f38-901">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="43f38-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f38-902">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-902">Requirements</span></span>

|<span data-ttu-id="43f38-903">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-903">Requirement</span></span>| <span data-ttu-id="43f38-904">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-905">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-906">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-906">1.0</span></span>|
|[<span data-ttu-id="43f38-907">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-908">ReadItem</span></span>|
|[<span data-ttu-id="43f38-909">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-910">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f38-911">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="43f38-911">Returns:</span></span>

<span data-ttu-id="43f38-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="43f38-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="43f38-914">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="43f38-914">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="43f38-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="43f38-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="43f38-916">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="43f38-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-917">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f38-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="43f38-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="43f38-921">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="43f38-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="43f38-922">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="43f38-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="43f38-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f38-926">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-926">Requirements</span></span>

|<span data-ttu-id="43f38-927">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-927">Requirement</span></span>| <span data-ttu-id="43f38-928">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-929">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-930">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-930">1.0</span></span>|
|[<span data-ttu-id="43f38-931">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-932">ReadItem</span></span>|
|[<span data-ttu-id="43f38-933">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-934">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f38-935">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="43f38-935">Returns:</span></span>

<span data-ttu-id="43f38-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="43f38-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="43f38-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="43f38-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43f38-939">Object</span><span class="sxs-lookup"><span data-stu-id="43f38-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="43f38-940">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-940">Example</span></span>

<span data-ttu-id="43f38-941">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="43f38-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="43f38-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="43f38-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="43f38-943">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="43f38-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-944">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="43f38-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f38-945">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="43f38-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="43f38-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="43f38-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-948">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-948">Parameters</span></span>

|<span data-ttu-id="43f38-949">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-949">Name</span></span>| <span data-ttu-id="43f38-950">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-950">Type</span></span>| <span data-ttu-id="43f38-951">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="43f38-952">String</span><span class="sxs-lookup"><span data-stu-id="43f38-952">String</span></span>|<span data-ttu-id="43f38-953">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="43f38-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f38-954">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-954">Requirements</span></span>

|<span data-ttu-id="43f38-955">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-955">Requirement</span></span>| <span data-ttu-id="43f38-956">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-957">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-958">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-958">1.0</span></span>|
|[<span data-ttu-id="43f38-959">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-959">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-960">ReadItem</span></span>|
|[<span data-ttu-id="43f38-961">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-961">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-962">Lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f38-963">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="43f38-963">Returns:</span></span>

<span data-ttu-id="43f38-964">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="43f38-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="43f38-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="43f38-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43f38-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="43f38-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="43f38-967">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="43f38-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="43f38-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="43f38-969">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="43f38-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="43f38-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="43f38-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-972">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-972">Parameters</span></span>

|<span data-ttu-id="43f38-973">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-973">Name</span></span>| <span data-ttu-id="43f38-974">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-974">Type</span></span>| <span data-ttu-id="43f38-975">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-975">Attributes</span></span>| <span data-ttu-id="43f38-976">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="43f38-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="43f38-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="43f38-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="43f38-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="43f38-981">Object</span><span class="sxs-lookup"><span data-stu-id="43f38-981">Object</span></span>| <span data-ttu-id="43f38-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-982">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-983">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="43f38-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43f38-984">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-984">Object</span></span>| <span data-ttu-id="43f38-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-985">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-986">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43f38-987">fonction</span><span class="sxs-lookup"><span data-stu-id="43f38-987">function</span></span>||<span data-ttu-id="43f38-988">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43f38-989">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="43f38-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="43f38-990">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="43f38-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f38-991">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-991">Requirements</span></span>

|<span data-ttu-id="43f38-992">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-992">Requirement</span></span>| <span data-ttu-id="43f38-993">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-994">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-995">1.2</span><span class="sxs-lookup"><span data-stu-id="43f38-995">1.2</span></span>|
|[<span data-ttu-id="43f38-996">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-996">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43f38-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="43f38-998">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-998">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-999">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f38-1000">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="43f38-1000">Returns:</span></span>

<span data-ttu-id="43f38-1001">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="43f38-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="43f38-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="43f38-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43f38-1003">String</span><span class="sxs-lookup"><span data-stu-id="43f38-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="43f38-1004">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-1004">Example</span></span>

```javascript
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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="43f38-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43f38-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="43f38-1006">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="43f38-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="43f38-p163">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="43f38-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-1010">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-1010">Parameters</span></span>

|<span data-ttu-id="43f38-1011">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-1011">Name</span></span>| <span data-ttu-id="43f38-1012">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-1012">Type</span></span>| <span data-ttu-id="43f38-1013">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-1013">Attributes</span></span>| <span data-ttu-id="43f38-1014">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="43f38-1015">function</span><span class="sxs-lookup"><span data-stu-id="43f38-1015">function</span></span>||<span data-ttu-id="43f38-1016">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43f38-1017">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43f38-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="43f38-1018">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="43f38-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="43f38-1019">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1019">Object</span></span>| <span data-ttu-id="43f38-1020">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1021">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="43f38-1022">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f38-1023">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-1023">Requirements</span></span>

|<span data-ttu-id="43f38-1024">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-1024">Requirement</span></span>| <span data-ttu-id="43f38-1025">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-1026">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="43f38-1027">1.0</span></span>|
|[<span data-ttu-id="43f38-1028">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-1028">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f38-1029">ReadItem</span></span>|
|[<span data-ttu-id="43f38-1030">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-1030">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-1031">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="43f38-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-1032">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-1032">Example</span></span>

<span data-ttu-id="43f38-p166">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="43f38-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="43f38-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43f38-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="43f38-1037">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43f38-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="43f38-p167">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook sur le web et appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="43f38-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-1042">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-1042">Parameters</span></span>

|<span data-ttu-id="43f38-1043">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-1043">Name</span></span>| <span data-ttu-id="43f38-1044">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-1044">Type</span></span>| <span data-ttu-id="43f38-1045">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-1045">Attributes</span></span>| <span data-ttu-id="43f38-1046">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="43f38-1047">String</span><span class="sxs-lookup"><span data-stu-id="43f38-1047">String</span></span>||<span data-ttu-id="43f38-1048">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="43f38-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="43f38-1049">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1049">Object</span></span>| <span data-ttu-id="43f38-1050">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1051">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="43f38-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43f38-1052">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1052">Object</span></span>| <span data-ttu-id="43f38-1053">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1054">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43f38-1055">fonction</span><span class="sxs-lookup"><span data-stu-id="43f38-1055">function</span></span>| <span data-ttu-id="43f38-1056">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1057">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="43f38-1058">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="43f38-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43f38-1059">Erreurs</span><span class="sxs-lookup"><span data-stu-id="43f38-1059">Errors</span></span>

| <span data-ttu-id="43f38-1060">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="43f38-1060">Error code</span></span> | <span data-ttu-id="43f38-1061">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="43f38-1062">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="43f38-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f38-1063">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-1063">Requirements</span></span>

|<span data-ttu-id="43f38-1064">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-1064">Requirement</span></span>| <span data-ttu-id="43f38-1065">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-1066">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="43f38-1067">1.1</span></span>|
|[<span data-ttu-id="43f38-1068">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43f38-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="43f38-1070">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-1071">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-1072">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-1072">Example</span></span>

<span data-ttu-id="43f38-1073">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="43f38-1073">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="43f38-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="43f38-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="43f38-1075">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="43f38-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="43f38-p168">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="43f38-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-1079">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="43f38-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="43f38-1080">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="43f38-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="43f38-p170">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="43f38-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="43f38-1084">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="43f38-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="43f38-1085">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="43f38-1085">Note: Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="43f38-1086">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="43f38-1086">The `saveAsync` method will fail when called from a meeting in compose mode.</span></span> <span data-ttu-id="43f38-1087">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="43f38-1087">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="43f38-1088">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="43f38-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-1089">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-1089">Parameters</span></span>

|<span data-ttu-id="43f38-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-1090">Name</span></span>| <span data-ttu-id="43f38-1091">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-1091">Type</span></span>| <span data-ttu-id="43f38-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-1092">Attributes</span></span>| <span data-ttu-id="43f38-1093">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="43f38-1094">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1094">Object</span></span>| <span data-ttu-id="43f38-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1096">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="43f38-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43f38-1097">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1097">Object</span></span>| <span data-ttu-id="43f38-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1099">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43f38-1100">fonction</span><span class="sxs-lookup"><span data-stu-id="43f38-1100">function</span></span>||<span data-ttu-id="43f38-1101">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43f38-1102">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43f38-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f38-1103">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-1103">Requirements</span></span>

|<span data-ttu-id="43f38-1104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-1104">Requirement</span></span>| <span data-ttu-id="43f38-1105">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-1106">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="43f38-1107">1.3</span></span>|
|[<span data-ttu-id="43f38-1108">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-1108">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43f38-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="43f38-1110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-1110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-1111">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="43f38-1112">範例</span><span class="sxs-lookup"><span data-stu-id="43f38-1112">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="43f38-p172">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="43f38-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="43f38-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="43f38-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="43f38-1116">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="43f38-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="43f38-p173">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="43f38-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f38-1120">Paramètres</span><span class="sxs-lookup"><span data-stu-id="43f38-1120">Parameters</span></span>

|<span data-ttu-id="43f38-1121">Nom</span><span class="sxs-lookup"><span data-stu-id="43f38-1121">Name</span></span>| <span data-ttu-id="43f38-1122">Type</span><span class="sxs-lookup"><span data-stu-id="43f38-1122">Type</span></span>| <span data-ttu-id="43f38-1123">Attributs</span><span class="sxs-lookup"><span data-stu-id="43f38-1123">Attributes</span></span>| <span data-ttu-id="43f38-1124">Description</span><span class="sxs-lookup"><span data-stu-id="43f38-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="43f38-1125">String</span><span class="sxs-lookup"><span data-stu-id="43f38-1125">String</span></span>||<span data-ttu-id="43f38-p174">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="43f38-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="43f38-1129">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1129">Object</span></span>| <span data-ttu-id="43f38-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1131">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="43f38-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43f38-1132">Objet</span><span class="sxs-lookup"><span data-stu-id="43f38-1132">Object</span></span>| <span data-ttu-id="43f38-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1134">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="43f38-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="43f38-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="43f38-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="43f38-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43f38-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="43f38-1137">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="43f38-1137">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="43f38-1138">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="43f38-1138">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="43f38-1139">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="43f38-1139">If html and the field supports HTML (the subject doesn't), the current style is applied in olwebshort and the default style is applied in Outlook.</span></span> <span data-ttu-id="43f38-1140">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="43f38-1140">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="43f38-1141">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="43f38-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="43f38-1142">fonction</span><span class="sxs-lookup"><span data-stu-id="43f38-1142">function</span></span>||<span data-ttu-id="43f38-1143">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43f38-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f38-1144">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="43f38-1144">Requirements</span></span>

|<span data-ttu-id="43f38-1145">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="43f38-1145">Requirement</span></span>| <span data-ttu-id="43f38-1146">Valeur</span><span class="sxs-lookup"><span data-stu-id="43f38-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f38-1147">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="43f38-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f38-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="43f38-1148">1.2</span></span>|
|[<span data-ttu-id="43f38-1149">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="43f38-1149">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f38-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43f38-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="43f38-1151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="43f38-1151">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43f38-1152">Composition</span><span class="sxs-lookup"><span data-stu-id="43f38-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43f38-1153">Exemple</span><span class="sxs-lookup"><span data-stu-id="43f38-1153">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
