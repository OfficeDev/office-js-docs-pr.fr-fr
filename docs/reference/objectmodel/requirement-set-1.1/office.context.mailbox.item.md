---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 20d3aaecc5e0c62f86a46ae29010a6462446bf1d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696441"
---
# <a name="item"></a><span data-ttu-id="1dafa-102">élément</span><span class="sxs-lookup"><span data-stu-id="1dafa-102">item</span></span>

### <span data-ttu-id="1dafa-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="1dafa-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="1dafa-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="1dafa-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-107">Requirements</span></span>

|<span data-ttu-id="1dafa-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-108">Requirement</span></span>| <span data-ttu-id="1dafa-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-111">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-111">1.0</span></span>|
|[<span data-ttu-id="1dafa-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="1dafa-113">Restricted</span></span>|
|[<span data-ttu-id="1dafa-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-115">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1dafa-116">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="1dafa-116">Members and methods</span></span>

| <span data-ttu-id="1dafa-117">Membre	</span><span class="sxs-lookup"><span data-stu-id="1dafa-117">Member</span></span> | <span data-ttu-id="1dafa-118">Type	</span><span class="sxs-lookup"><span data-stu-id="1dafa-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1dafa-119">attachments</span><span class="sxs-lookup"><span data-stu-id="1dafa-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="1dafa-120">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-120">Member</span></span> |
| [<span data-ttu-id="1dafa-121">bcc</span><span class="sxs-lookup"><span data-stu-id="1dafa-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="1dafa-122">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-122">Member</span></span> |
| [<span data-ttu-id="1dafa-123">body</span><span class="sxs-lookup"><span data-stu-id="1dafa-123">body</span></span>](#body-body) | <span data-ttu-id="1dafa-124">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-124">Member</span></span> |
| [<span data-ttu-id="1dafa-125">cc</span><span class="sxs-lookup"><span data-stu-id="1dafa-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1dafa-126">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-126">Member</span></span> |
| [<span data-ttu-id="1dafa-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="1dafa-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1dafa-128">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-128">Member</span></span> |
| [<span data-ttu-id="1dafa-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1dafa-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1dafa-130">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-130">Member</span></span> |
| [<span data-ttu-id="1dafa-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1dafa-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1dafa-132">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-132">Member</span></span> |
| [<span data-ttu-id="1dafa-133">end</span><span class="sxs-lookup"><span data-stu-id="1dafa-133">end</span></span>](#end-datetime) | <span data-ttu-id="1dafa-134">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-134">Member</span></span> |
| [<span data-ttu-id="1dafa-135">from</span><span class="sxs-lookup"><span data-stu-id="1dafa-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="1dafa-136">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-136">Member</span></span> |
| [<span data-ttu-id="1dafa-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1dafa-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1dafa-138">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-138">Member</span></span> |
| [<span data-ttu-id="1dafa-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="1dafa-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1dafa-140">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-140">Member</span></span> |
| [<span data-ttu-id="1dafa-141">itemId</span><span class="sxs-lookup"><span data-stu-id="1dafa-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1dafa-142">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-142">Member</span></span> |
| [<span data-ttu-id="1dafa-143">itemType</span><span class="sxs-lookup"><span data-stu-id="1dafa-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="1dafa-144">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-144">Member</span></span> |
| [<span data-ttu-id="1dafa-145">location</span><span class="sxs-lookup"><span data-stu-id="1dafa-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="1dafa-146">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-146">Member</span></span> |
| [<span data-ttu-id="1dafa-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1dafa-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1dafa-148">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-148">Member</span></span> |
| [<span data-ttu-id="1dafa-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1dafa-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1dafa-150">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-150">Member</span></span> |
| [<span data-ttu-id="1dafa-151">organizer</span><span class="sxs-lookup"><span data-stu-id="1dafa-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="1dafa-152">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-152">Member</span></span> |
| [<span data-ttu-id="1dafa-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1dafa-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1dafa-154">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-154">Member</span></span> |
| [<span data-ttu-id="1dafa-155">sender</span><span class="sxs-lookup"><span data-stu-id="1dafa-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="1dafa-156">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-156">Member</span></span> |
| [<span data-ttu-id="1dafa-157">start</span><span class="sxs-lookup"><span data-stu-id="1dafa-157">start</span></span>](#start-datetime) | <span data-ttu-id="1dafa-158">Member</span><span class="sxs-lookup"><span data-stu-id="1dafa-158">Member</span></span> |
| [<span data-ttu-id="1dafa-159">subject</span><span class="sxs-lookup"><span data-stu-id="1dafa-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="1dafa-160">Membre</span><span class="sxs-lookup"><span data-stu-id="1dafa-160">Member</span></span> |
| [<span data-ttu-id="1dafa-161">to</span><span class="sxs-lookup"><span data-stu-id="1dafa-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1dafa-162">Membre</span><span class="sxs-lookup"><span data-stu-id="1dafa-162">Member</span></span> |
| [<span data-ttu-id="1dafa-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1dafa-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1dafa-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-164">Method</span></span> |
| [<span data-ttu-id="1dafa-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1dafa-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1dafa-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-166">Method</span></span> |
| [<span data-ttu-id="1dafa-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1dafa-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="1dafa-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-168">Method</span></span> |
| [<span data-ttu-id="1dafa-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1dafa-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="1dafa-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-170">Method</span></span> |
| [<span data-ttu-id="1dafa-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="1dafa-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="1dafa-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-172">Method</span></span> |
| [<span data-ttu-id="1dafa-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1dafa-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1dafa-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-174">Method</span></span> |
| [<span data-ttu-id="1dafa-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1dafa-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1dafa-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-176">Method</span></span> |
| [<span data-ttu-id="1dafa-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1dafa-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1dafa-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-178">Method</span></span> |
| [<span data-ttu-id="1dafa-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1dafa-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1dafa-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-180">Method</span></span> |
| [<span data-ttu-id="1dafa-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1dafa-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1dafa-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-182">Method</span></span> |
| [<span data-ttu-id="1dafa-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1dafa-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1dafa-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="1dafa-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1dafa-185">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-185">Example</span></span>

<span data-ttu-id="1dafa-186">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="1dafa-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1dafa-187">Membres</span><span class="sxs-lookup"><span data-stu-id="1dafa-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="1dafa-188">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="1dafa-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="1dafa-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-191">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="1dafa-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1dafa-192">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="1dafa-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-193">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-193">Type</span></span>

*   <span data-ttu-id="1dafa-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="1dafa-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-195">Requirements</span></span>

|<span data-ttu-id="1dafa-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-196">Requirement</span></span>| <span data-ttu-id="1dafa-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-199">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-199">1.0</span></span>|
|[<span data-ttu-id="1dafa-200">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-201">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-202">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-203">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-204">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-204">Example</span></span>

<span data-ttu-id="1dafa-205">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="1dafa-206">CCI: [destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-207">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="1dafa-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1dafa-208">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-209">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-209">Type</span></span>

*   [<span data-ttu-id="1dafa-210">Destinataires</span><span class="sxs-lookup"><span data-stu-id="1dafa-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="1dafa-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-211">Requirements</span></span>

|<span data-ttu-id="1dafa-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-212">Requirement</span></span>| <span data-ttu-id="1dafa-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-215">1.1</span><span class="sxs-lookup"><span data-stu-id="1dafa-215">1.1</span></span>|
|[<span data-ttu-id="1dafa-216">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-217">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-218">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-219">Composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-220">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-220">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="1dafa-221">Body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-222">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="1dafa-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-223">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-223">Type</span></span>

*   [<span data-ttu-id="1dafa-224">Body</span><span class="sxs-lookup"><span data-stu-id="1dafa-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="1dafa-225">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-225">Requirements</span></span>

|<span data-ttu-id="1dafa-226">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-226">Requirement</span></span>| <span data-ttu-id="1dafa-227">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-228">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-229">1.1</span><span class="sxs-lookup"><span data-stu-id="1dafa-229">1.1</span></span>|
|[<span data-ttu-id="1dafa-230">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-231">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-232">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-233">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-234">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-234">Example</span></span>

<span data-ttu-id="1dafa-235">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="1dafa-235">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1dafa-236">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-236">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="1dafa-237">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-238">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="1dafa-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1dafa-239">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-240">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-240">Read mode</span></span>

<span data-ttu-id="1dafa-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-243">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-243">Compose mode</span></span>

<span data-ttu-id="1dafa-244">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="1dafa-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1dafa-245">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-245">Type</span></span>

*   <span data-ttu-id="1dafa-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-247">Requirements</span></span>

|<span data-ttu-id="1dafa-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-248">Requirement</span></span>| <span data-ttu-id="1dafa-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-251">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-251">1.0</span></span>|
|[<span data-ttu-id="1dafa-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-253">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="1dafa-256">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="1dafa-257">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="1dafa-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1dafa-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1dafa-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-262">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-262">Type</span></span>

*   <span data-ttu-id="1dafa-263">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-264">Requirements</span></span>

|<span data-ttu-id="1dafa-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-265">Requirement</span></span>| <span data-ttu-id="1dafa-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-268">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-268">1.0</span></span>|
|[<span data-ttu-id="1dafa-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-270">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-272">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-273">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-273">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="1dafa-274">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="1dafa-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="1dafa-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-277">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-277">Type</span></span>

*   <span data-ttu-id="1dafa-278">Date</span><span class="sxs-lookup"><span data-stu-id="1dafa-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-279">Requirements</span></span>

|<span data-ttu-id="1dafa-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-280">Requirement</span></span>| <span data-ttu-id="1dafa-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-283">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-283">1.0</span></span>|
|[<span data-ttu-id="1dafa-284">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-285">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-287">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-288">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="1dafa-289">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="1dafa-289">dateTimeModified: Date</span></span>

<span data-ttu-id="1dafa-290">Obtient la date et l’heure de la dernière modification d’un élément.</span><span class="sxs-lookup"><span data-stu-id="1dafa-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="1dafa-291">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-292">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-293">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-293">Type</span></span>

*   <span data-ttu-id="1dafa-294">Date</span><span class="sxs-lookup"><span data-stu-id="1dafa-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-295">Requirements</span></span>

|<span data-ttu-id="1dafa-296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-296">Requirement</span></span>| <span data-ttu-id="1dafa-297">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-299">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-299">1.0</span></span>|
|[<span data-ttu-id="1dafa-300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-301">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-303">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-304">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-304">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="1dafa-305">fin: date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-306">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1dafa-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-309">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-309">Read mode</span></span>

<span data-ttu-id="1dafa-310">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-310">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-311">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-311">Compose mode</span></span>

<span data-ttu-id="1dafa-312">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1dafa-313">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="1dafa-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1dafa-314">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1dafa-315">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-315">Type</span></span>

*   <span data-ttu-id="1dafa-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-317">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-317">Requirements</span></span>

|<span data-ttu-id="1dafa-318">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-318">Requirement</span></span>| <span data-ttu-id="1dafa-319">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-320">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-321">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-321">1.0</span></span>|
|[<span data-ttu-id="1dafa-322">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-323">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-324">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-325">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-325">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="1dafa-326">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1dafa-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-331">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-332">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-332">Type</span></span>

*   [<span data-ttu-id="1dafa-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1dafa-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="1dafa-334">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-334">Requirements</span></span>

|<span data-ttu-id="1dafa-335">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-335">Requirement</span></span>| <span data-ttu-id="1dafa-336">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-337">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-338">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-338">1.0</span></span>|
|[<span data-ttu-id="1dafa-339">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-340">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-341">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-342">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-343">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-343">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="1dafa-344">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-344">internetMessageId: String</span></span>

<span data-ttu-id="1dafa-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-347">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-347">Type</span></span>

*   <span data-ttu-id="1dafa-348">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-349">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-349">Requirements</span></span>

|<span data-ttu-id="1dafa-350">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-350">Requirement</span></span>| <span data-ttu-id="1dafa-351">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-352">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-353">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-353">1.0</span></span>|
|[<span data-ttu-id="1dafa-354">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-355">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-356">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-357">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-358">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-358">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="1dafa-359">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-359">itemClass: String</span></span>

<span data-ttu-id="1dafa-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1dafa-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1dafa-364">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-364">Type</span></span> | <span data-ttu-id="1dafa-365">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-365">Description</span></span> | <span data-ttu-id="1dafa-366">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="1dafa-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1dafa-367">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="1dafa-367">Appointment items</span></span> | <span data-ttu-id="1dafa-368">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="1dafa-369">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="1dafa-369">Message items</span></span> | <span data-ttu-id="1dafa-370">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="1dafa-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1dafa-371">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-372">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-372">Type</span></span>

*   <span data-ttu-id="1dafa-373">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-374">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-374">Requirements</span></span>

|<span data-ttu-id="1dafa-375">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-375">Requirement</span></span>| <span data-ttu-id="1dafa-376">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-377">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-378">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-378">1.0</span></span>|
|[<span data-ttu-id="1dafa-379">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-380">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-381">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-382">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-383">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-383">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1dafa-384">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="1dafa-384">(nullable) itemId: String</span></span>

<span data-ttu-id="1dafa-385">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="1dafa-386">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-387">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="1dafa-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1dafa-388">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="1dafa-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1dafa-389">Avant d’effectuer des appels d’API REST à l’aide de cette valeur `Office.context.mailbox.convertToRestId`, elle doit être convertie à l’aide de, qui est disponible à partir de l’ensemble de conditions requises 1,3.</span><span class="sxs-lookup"><span data-stu-id="1dafa-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="1dafa-390">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="1dafa-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-391">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-391">Type</span></span>

*   <span data-ttu-id="1dafa-392">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-393">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-393">Requirements</span></span>

|<span data-ttu-id="1dafa-394">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-394">Requirement</span></span>| <span data-ttu-id="1dafa-395">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-396">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-397">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-397">1.0</span></span>|
|[<span data-ttu-id="1dafa-398">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-399">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-400">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-401">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-402">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-402">Example</span></span>

<span data-ttu-id="1dafa-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="1dafa-405">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-406">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="1dafa-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1dafa-407">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-408">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-408">Type</span></span>

*   [<span data-ttu-id="1dafa-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1dafa-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="1dafa-410">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-410">Requirements</span></span>

|<span data-ttu-id="1dafa-411">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-411">Requirement</span></span>| <span data-ttu-id="1dafa-412">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-413">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-414">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-414">1.0</span></span>|
|[<span data-ttu-id="1dafa-415">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-416">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-417">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-418">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-419">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-419">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="1dafa-420">Location: String | [Emplacement](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-421">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-422">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-422">Read mode</span></span>

<span data-ttu-id="1dafa-423">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-424">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-424">Compose mode</span></span>

<span data-ttu-id="1dafa-425">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1dafa-426">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-426">Type</span></span>

*   <span data-ttu-id="1dafa-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-428">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-428">Requirements</span></span>

|<span data-ttu-id="1dafa-429">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-429">Requirement</span></span>| <span data-ttu-id="1dafa-430">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-431">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-432">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-432">1.0</span></span>|
|[<span data-ttu-id="1dafa-433">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-434">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-435">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-436">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-436">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1dafa-437">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-437">normalizedSubject: String</span></span>

<span data-ttu-id="1dafa-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1dafa-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="1dafa-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-442">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-442">Type</span></span>

*   <span data-ttu-id="1dafa-443">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-444">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-444">Requirements</span></span>

|<span data-ttu-id="1dafa-445">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-445">Requirement</span></span>| <span data-ttu-id="1dafa-446">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-447">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-448">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-448">1.0</span></span>|
|[<span data-ttu-id="1dafa-449">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-450">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-451">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-452">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-453">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-453">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="1dafa-454">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="1dafa-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-455">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1dafa-456">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-457">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-457">Read mode</span></span>

<span data-ttu-id="1dafa-458">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="1dafa-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-459">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-459">Compose mode</span></span>

<span data-ttu-id="1dafa-460">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="1dafa-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1dafa-461">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-461">Type</span></span>

*   <span data-ttu-id="1dafa-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-463">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-463">Requirements</span></span>

|<span data-ttu-id="1dafa-464">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-464">Requirement</span></span>| <span data-ttu-id="1dafa-465">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-466">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-467">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-467">1.0</span></span>|
|[<span data-ttu-id="1dafa-468">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-469">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-470">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-471">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-471">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="1dafa-472">Organisateur: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-475">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-475">Type</span></span>

*   [<span data-ttu-id="1dafa-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1dafa-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="1dafa-477">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-477">Requirements</span></span>

|<span data-ttu-id="1dafa-478">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-478">Requirement</span></span>| <span data-ttu-id="1dafa-479">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-480">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-481">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-481">1.0</span></span>|
|[<span data-ttu-id="1dafa-482">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-483">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-484">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-485">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-486">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-486">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="1dafa-487">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="1dafa-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-488">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1dafa-489">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-490">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-490">Read mode</span></span>

<span data-ttu-id="1dafa-491">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="1dafa-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-492">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-492">Compose mode</span></span>

<span data-ttu-id="1dafa-493">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="1dafa-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1dafa-494">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-494">Type</span></span>

*   <span data-ttu-id="1dafa-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-496">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-496">Requirements</span></span>

|<span data-ttu-id="1dafa-497">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-497">Requirement</span></span>| <span data-ttu-id="1dafa-498">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-499">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-500">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-500">1.0</span></span>|
|[<span data-ttu-id="1dafa-501">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-502">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-503">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-504">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-504">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="1dafa-505">expéditeur: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1dafa-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-510">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1dafa-511">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-511">Type</span></span>

*   [<span data-ttu-id="1dafa-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1dafa-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="1dafa-513">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-513">Requirements</span></span>

|<span data-ttu-id="1dafa-514">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-514">Requirement</span></span>| <span data-ttu-id="1dafa-515">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-516">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-517">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-517">1.0</span></span>|
|[<span data-ttu-id="1dafa-518">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-519">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-520">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-521">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-522">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-522">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="1dafa-523">début: date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-524">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1dafa-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-527">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-527">Read mode</span></span>

<span data-ttu-id="1dafa-528">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-528">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-529">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-529">Compose mode</span></span>

<span data-ttu-id="1dafa-530">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1dafa-531">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="1dafa-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1dafa-532">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1dafa-533">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-533">Type</span></span>

*   <span data-ttu-id="1dafa-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-535">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-535">Requirements</span></span>

|<span data-ttu-id="1dafa-536">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-536">Requirement</span></span>| <span data-ttu-id="1dafa-537">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-538">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-539">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-539">1.0</span></span>|
|[<span data-ttu-id="1dafa-540">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-541">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-542">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-543">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-543">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="1dafa-544">Subject: String | [Objet](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-545">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="1dafa-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1dafa-546">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="1dafa-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-547">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-547">Read mode</span></span>

<span data-ttu-id="1dafa-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-550">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-550">Compose mode</span></span>

<span data-ttu-id="1dafa-551">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="1dafa-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1dafa-552">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-552">Type</span></span>

*   <span data-ttu-id="1dafa-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-554">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-554">Requirements</span></span>

|<span data-ttu-id="1dafa-555">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-555">Requirement</span></span>| <span data-ttu-id="1dafa-556">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-557">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-558">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-558">1.0</span></span>|
|[<span data-ttu-id="1dafa-559">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-560">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-561">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-562">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-562">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="1dafa-563">to: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="1dafa-564">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="1dafa-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1dafa-565">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dafa-566">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-566">Read mode</span></span>

<span data-ttu-id="1dafa-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1dafa-569">Mode composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-569">Compose mode</span></span>

<span data-ttu-id="1dafa-570">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="1dafa-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1dafa-571">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-571">Type</span></span>

*   <span data-ttu-id="1dafa-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-573">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-573">Requirements</span></span>

|<span data-ttu-id="1dafa-574">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-574">Requirement</span></span>| <span data-ttu-id="1dafa-575">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-576">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-577">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-577">1.0</span></span>|
|[<span data-ttu-id="1dafa-578">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-579">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-580">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-581">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1dafa-582">Méthodes</span><span class="sxs-lookup"><span data-stu-id="1dafa-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1dafa-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1dafa-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1dafa-584">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="1dafa-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1dafa-585">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="1dafa-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1dafa-586">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="1dafa-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-587">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-587">Parameters</span></span>

|<span data-ttu-id="1dafa-588">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-588">Name</span></span>| <span data-ttu-id="1dafa-589">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-589">Type</span></span>| <span data-ttu-id="1dafa-590">Attributs</span><span class="sxs-lookup"><span data-stu-id="1dafa-590">Attributes</span></span>| <span data-ttu-id="1dafa-591">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1dafa-592">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-592">String</span></span>||<span data-ttu-id="1dafa-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1dafa-595">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-595">String</span></span>||<span data-ttu-id="1dafa-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1dafa-598">Objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-598">Object</span></span>| <span data-ttu-id="1dafa-599">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-599">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-600">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1dafa-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dafa-601">Objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-601">Object</span></span>| <span data-ttu-id="1dafa-602">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-602">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-603">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dafa-604">fonction</span><span class="sxs-lookup"><span data-stu-id="1dafa-604">function</span></span>| <span data-ttu-id="1dafa-605">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-605">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-606">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dafa-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1dafa-607">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1dafa-608">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="1dafa-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1dafa-609">Erreurs</span><span class="sxs-lookup"><span data-stu-id="1dafa-609">Errors</span></span>

| <span data-ttu-id="1dafa-610">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="1dafa-610">Error code</span></span> | <span data-ttu-id="1dafa-611">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1dafa-612">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="1dafa-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1dafa-613">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="1dafa-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1dafa-614">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="1dafa-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dafa-615">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-615">Requirements</span></span>

|<span data-ttu-id="1dafa-616">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-616">Requirement</span></span>| <span data-ttu-id="1dafa-617">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-618">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-619">1.1</span><span class="sxs-lookup"><span data-stu-id="1dafa-619">1.1</span></span>|
|[<span data-ttu-id="1dafa-620">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dafa-622">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-623">Composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-624">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-624">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1dafa-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1dafa-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1dafa-626">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1dafa-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1dafa-630">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="1dafa-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1dafa-631">Si votre complément Office est en cours d’exécution dans Outlook sur le Web, `addItemAttachmentAsync` la méthode peut joindre des éléments à des éléments autres que l’élément que vous modifiez; Toutefois, cette option n’est pas prise en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="1dafa-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-632">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-632">Parameters</span></span>

|<span data-ttu-id="1dafa-633">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-633">Name</span></span>| <span data-ttu-id="1dafa-634">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-634">Type</span></span>| <span data-ttu-id="1dafa-635">Attributs</span><span class="sxs-lookup"><span data-stu-id="1dafa-635">Attributes</span></span>| <span data-ttu-id="1dafa-636">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1dafa-637">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-637">String</span></span>||<span data-ttu-id="1dafa-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1dafa-640">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-640">String</span></span>||<span data-ttu-id="1dafa-641">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="1dafa-641">The subject of the item to be attached.</span></span> <span data-ttu-id="1dafa-642">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1dafa-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1dafa-643">Object</span><span class="sxs-lookup"><span data-stu-id="1dafa-643">Object</span></span>| <span data-ttu-id="1dafa-644">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-644">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-645">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1dafa-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dafa-646">Objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-646">Object</span></span>| <span data-ttu-id="1dafa-647">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-647">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-648">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dafa-649">fonction</span><span class="sxs-lookup"><span data-stu-id="1dafa-649">function</span></span>| <span data-ttu-id="1dafa-650">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-650">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-651">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dafa-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1dafa-652">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1dafa-653">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="1dafa-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1dafa-654">Erreurs</span><span class="sxs-lookup"><span data-stu-id="1dafa-654">Errors</span></span>

| <span data-ttu-id="1dafa-655">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="1dafa-655">Error code</span></span> | <span data-ttu-id="1dafa-656">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1dafa-657">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="1dafa-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dafa-658">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-658">Requirements</span></span>

|<span data-ttu-id="1dafa-659">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-659">Requirement</span></span>| <span data-ttu-id="1dafa-660">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-661">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-662">1.1</span><span class="sxs-lookup"><span data-stu-id="1dafa-662">1.1</span></span>|
|[<span data-ttu-id="1dafa-663">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dafa-665">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-666">Composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-667">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-667">Example</span></span>

<span data-ttu-id="1dafa-668">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1dafa-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1dafa-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1dafa-670">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1dafa-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-671">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1dafa-672">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="1dafa-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1dafa-673">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="1dafa-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-674">La possibilité d’inclure des pièces jointes dans `displayReplyAllForm` l’appel à n’est pas prise en charge dans l’ensemble de conditions requises 1,1.</span><span class="sxs-lookup"><span data-stu-id="1dafa-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="1dafa-675">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="1dafa-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-676">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-676">Parameters</span></span>

|<span data-ttu-id="1dafa-677">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-677">Name</span></span>| <span data-ttu-id="1dafa-678">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-678">Type</span></span>| <span data-ttu-id="1dafa-679">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1dafa-680">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1dafa-680">String &#124; Object</span></span>| |<span data-ttu-id="1dafa-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1dafa-683">**OU**</span><span class="sxs-lookup"><span data-stu-id="1dafa-683">**OR**</span></span><br/><span data-ttu-id="1dafa-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="1dafa-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1dafa-686">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-686">String</span></span> | <span data-ttu-id="1dafa-687">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-687">&lt;optional&gt;</span></span> | <span data-ttu-id="1dafa-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="1dafa-690">fonction</span><span class="sxs-lookup"><span data-stu-id="1dafa-690">function</span></span> | <span data-ttu-id="1dafa-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-691">&lt;optional&gt;</span></span> | <span data-ttu-id="1dafa-692">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dafa-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dafa-693">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-693">Requirements</span></span>

|<span data-ttu-id="1dafa-694">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-694">Requirement</span></span>| <span data-ttu-id="1dafa-695">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-696">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-697">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-697">1.0</span></span>|
|[<span data-ttu-id="1dafa-698">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-699">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-700">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-701">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1dafa-702">Exemples</span><span class="sxs-lookup"><span data-stu-id="1dafa-702">Examples</span></span>

<span data-ttu-id="1dafa-703">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1dafa-704">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="1dafa-704">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1dafa-705">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="1dafa-705">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1dafa-706">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-706">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1dafa-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1dafa-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1dafa-708">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1dafa-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-709">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1dafa-710">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="1dafa-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1dafa-711">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="1dafa-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-712">La possibilité d’inclure des pièces jointes dans `displayReplyForm` l’appel à n’est pas prise en charge dans l’ensemble de conditions requises 1,1.</span><span class="sxs-lookup"><span data-stu-id="1dafa-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="1dafa-713">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="1dafa-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-714">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-714">Parameters</span></span>

|<span data-ttu-id="1dafa-715">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-715">Name</span></span>| <span data-ttu-id="1dafa-716">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-716">Type</span></span>| <span data-ttu-id="1dafa-717">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1dafa-718">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1dafa-718">String &#124; Object</span></span>| | <span data-ttu-id="1dafa-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1dafa-721">**OU**</span><span class="sxs-lookup"><span data-stu-id="1dafa-721">**OR**</span></span><br/><span data-ttu-id="1dafa-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="1dafa-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1dafa-724">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-724">String</span></span> | <span data-ttu-id="1dafa-725">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-725">&lt;optional&gt;</span></span> | <span data-ttu-id="1dafa-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="1dafa-728">fonction</span><span class="sxs-lookup"><span data-stu-id="1dafa-728">function</span></span> | <span data-ttu-id="1dafa-729">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-729">&lt;optional&gt;</span></span> | <span data-ttu-id="1dafa-730">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dafa-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dafa-731">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-731">Requirements</span></span>

|<span data-ttu-id="1dafa-732">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-732">Requirement</span></span>| <span data-ttu-id="1dafa-733">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-734">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-735">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-735">1.0</span></span>|
|[<span data-ttu-id="1dafa-736">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-737">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-738">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-739">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1dafa-740">Exemples</span><span class="sxs-lookup"><span data-stu-id="1dafa-740">Examples</span></span>

<span data-ttu-id="1dafa-741">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1dafa-742">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="1dafa-742">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1dafa-743">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="1dafa-743">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1dafa-744">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-744">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="1dafa-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="1dafa-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="1dafa-746">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1dafa-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-747">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-748">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-748">Requirements</span></span>

|<span data-ttu-id="1dafa-749">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-749">Requirement</span></span>| <span data-ttu-id="1dafa-750">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-751">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-752">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-752">1.0</span></span>|
|[<span data-ttu-id="1dafa-753">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-754">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-755">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-756">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dafa-757">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1dafa-757">Returns:</span></span>

<span data-ttu-id="1dafa-758">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="1dafa-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="1dafa-759">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-759">Example</span></span>

<span data-ttu-id="1dafa-760">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-760">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="1dafa-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="1dafa-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="1dafa-762">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1dafa-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-763">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-764">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-764">Parameters</span></span>

|<span data-ttu-id="1dafa-765">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-765">Name</span></span>| <span data-ttu-id="1dafa-766">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-766">Type</span></span>| <span data-ttu-id="1dafa-767">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1dafa-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1dafa-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="1dafa-769">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="1dafa-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dafa-770">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-770">Requirements</span></span>

|<span data-ttu-id="1dafa-771">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-771">Requirement</span></span>| <span data-ttu-id="1dafa-772">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-773">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-774">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-774">1.0</span></span>|
|[<span data-ttu-id="1dafa-775">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-776">Restreinte</span><span class="sxs-lookup"><span data-stu-id="1dafa-776">Restricted</span></span>|
|[<span data-ttu-id="1dafa-777">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-778">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dafa-779">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1dafa-779">Returns:</span></span>

<span data-ttu-id="1dafa-780">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="1dafa-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1dafa-781">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="1dafa-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1dafa-782">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1dafa-783">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="1dafa-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1dafa-784">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="1dafa-784">Value of `entityType`</span></span> | <span data-ttu-id="1dafa-785">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="1dafa-785">Type of objects in returned array</span></span> | <span data-ttu-id="1dafa-786">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="1dafa-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1dafa-787">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-787">String</span></span> | <span data-ttu-id="1dafa-788">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1dafa-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1dafa-789">Contact</span><span class="sxs-lookup"><span data-stu-id="1dafa-789">Contact</span></span> | <span data-ttu-id="1dafa-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dafa-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1dafa-791">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-791">String</span></span> | <span data-ttu-id="1dafa-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dafa-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1dafa-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1dafa-793">MeetingSuggestion</span></span> | <span data-ttu-id="1dafa-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dafa-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1dafa-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1dafa-795">PhoneNumber</span></span> | <span data-ttu-id="1dafa-796">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1dafa-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1dafa-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1dafa-797">TaskSuggestion</span></span> | <span data-ttu-id="1dafa-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dafa-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1dafa-799">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-799">String</span></span> | <span data-ttu-id="1dafa-800">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1dafa-800">**Restricted**</span></span> |

<span data-ttu-id="1dafa-801">Type :  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="1dafa-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="1dafa-802">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-802">Example</span></span>

<span data-ttu-id="1dafa-803">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1dafa-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="1dafa-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="1dafa-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="1dafa-805">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="1dafa-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-806">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1dafa-807">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="1dafa-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-808">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-808">Parameters</span></span>

|<span data-ttu-id="1dafa-809">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-809">Name</span></span>| <span data-ttu-id="1dafa-810">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-810">Type</span></span>| <span data-ttu-id="1dafa-811">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1dafa-812">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-812">String</span></span>|<span data-ttu-id="1dafa-813">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="1dafa-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dafa-814">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-814">Requirements</span></span>

|<span data-ttu-id="1dafa-815">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-815">Requirement</span></span>| <span data-ttu-id="1dafa-816">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-817">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-818">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-818">1.0</span></span>|
|[<span data-ttu-id="1dafa-819">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-820">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-821">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-822">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dafa-823">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1dafa-823">Returns:</span></span>

<span data-ttu-id="1dafa-p146">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="1dafa-826">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="1dafa-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="1dafa-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1dafa-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1dafa-828">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="1dafa-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-829">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1dafa-p147">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1dafa-833">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="1dafa-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1dafa-834">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="1dafa-p148">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dafa-837">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-837">Requirements</span></span>

|<span data-ttu-id="1dafa-838">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-838">Requirement</span></span>| <span data-ttu-id="1dafa-839">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-840">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-841">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-841">1.0</span></span>|
|[<span data-ttu-id="1dafa-842">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-843">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-844">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-845">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dafa-846">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1dafa-846">Returns:</span></span>

<span data-ttu-id="1dafa-p149">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="1dafa-849">Type: objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-849">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="1dafa-850">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-850">Example</span></span>

<span data-ttu-id="1dafa-851">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="1dafa-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1dafa-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1dafa-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1dafa-853">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="1dafa-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1dafa-854">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="1dafa-854">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1dafa-855">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="1dafa-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1dafa-p150">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-858">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-858">Parameters</span></span>

|<span data-ttu-id="1dafa-859">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-859">Name</span></span>| <span data-ttu-id="1dafa-860">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-860">Type</span></span>| <span data-ttu-id="1dafa-861">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1dafa-862">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1dafa-862">String</span></span>|<span data-ttu-id="1dafa-863">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="1dafa-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dafa-864">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-864">Requirements</span></span>

|<span data-ttu-id="1dafa-865">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-865">Requirement</span></span>| <span data-ttu-id="1dafa-866">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-867">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-868">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-868">1.0</span></span>|
|[<span data-ttu-id="1dafa-869">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-869">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-870">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-871">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-871">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-872">Lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dafa-873">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1dafa-873">Returns:</span></span>

<span data-ttu-id="1dafa-874">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="1dafa-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="1dafa-875">Type: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="1dafa-875">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="1dafa-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1dafa-877">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1dafa-877">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1dafa-878">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1dafa-878">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1dafa-p151">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-882">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-882">Parameters</span></span>

|<span data-ttu-id="1dafa-883">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-883">Name</span></span>| <span data-ttu-id="1dafa-884">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-884">Type</span></span>| <span data-ttu-id="1dafa-885">Attributs</span><span class="sxs-lookup"><span data-stu-id="1dafa-885">Attributes</span></span>| <span data-ttu-id="1dafa-886">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-886">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1dafa-887">function</span><span class="sxs-lookup"><span data-stu-id="1dafa-887">function</span></span>||<span data-ttu-id="1dafa-888">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dafa-888">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1dafa-889">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1dafa-889">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1dafa-890">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="1dafa-890">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1dafa-891">Objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-891">Object</span></span>| <span data-ttu-id="1dafa-892">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-892">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-893">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-893">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1dafa-894">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-894">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dafa-895">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-895">Requirements</span></span>

|<span data-ttu-id="1dafa-896">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-896">Requirement</span></span>| <span data-ttu-id="1dafa-897">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-897">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-898">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-898">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-899">1.0</span><span class="sxs-lookup"><span data-stu-id="1dafa-899">1.0</span></span>|
|[<span data-ttu-id="1dafa-900">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-900">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-901">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-901">ReadItem</span></span>|
|[<span data-ttu-id="1dafa-902">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-902">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-903">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1dafa-903">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-904">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-904">Example</span></span>

<span data-ttu-id="1dafa-p154">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1dafa-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1dafa-908">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1dafa-908">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1dafa-909">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1dafa-909">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1dafa-910">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="1dafa-910">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="1dafa-911">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="1dafa-911">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="1dafa-912">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="1dafa-912">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="1dafa-913">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="1dafa-913">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dafa-914">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1dafa-914">Parameters</span></span>

|<span data-ttu-id="1dafa-915">Nom</span><span class="sxs-lookup"><span data-stu-id="1dafa-915">Name</span></span>| <span data-ttu-id="1dafa-916">Type</span><span class="sxs-lookup"><span data-stu-id="1dafa-916">Type</span></span>| <span data-ttu-id="1dafa-917">Attributs</span><span class="sxs-lookup"><span data-stu-id="1dafa-917">Attributes</span></span>| <span data-ttu-id="1dafa-918">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-918">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1dafa-919">String</span><span class="sxs-lookup"><span data-stu-id="1dafa-919">String</span></span>||<span data-ttu-id="1dafa-920">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="1dafa-920">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1dafa-921">Objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-921">Object</span></span>| <span data-ttu-id="1dafa-922">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-922">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-923">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1dafa-923">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dafa-924">Objet</span><span class="sxs-lookup"><span data-stu-id="1dafa-924">Object</span></span>| <span data-ttu-id="1dafa-925">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-925">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-926">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dafa-926">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dafa-927">fonction</span><span class="sxs-lookup"><span data-stu-id="1dafa-927">function</span></span>| <span data-ttu-id="1dafa-928">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1dafa-928">&lt;optional&gt;</span></span>|<span data-ttu-id="1dafa-929">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dafa-929">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1dafa-930">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="1dafa-930">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1dafa-931">Erreurs</span><span class="sxs-lookup"><span data-stu-id="1dafa-931">Errors</span></span>

| <span data-ttu-id="1dafa-932">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="1dafa-932">Error code</span></span> | <span data-ttu-id="1dafa-933">Description</span><span class="sxs-lookup"><span data-stu-id="1dafa-933">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1dafa-934">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="1dafa-934">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dafa-935">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1dafa-935">Requirements</span></span>

|<span data-ttu-id="1dafa-936">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1dafa-936">Requirement</span></span>| <span data-ttu-id="1dafa-937">Valeur</span><span class="sxs-lookup"><span data-stu-id="1dafa-937">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dafa-938">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1dafa-938">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dafa-939">1.1</span><span class="sxs-lookup"><span data-stu-id="1dafa-939">1.1</span></span>|
|[<span data-ttu-id="1dafa-940">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1dafa-940">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dafa-941">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dafa-941">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dafa-942">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1dafa-942">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1dafa-943">Composition</span><span class="sxs-lookup"><span data-stu-id="1dafa-943">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dafa-944">Exemple</span><span class="sxs-lookup"><span data-stu-id="1dafa-944">Example</span></span>

<span data-ttu-id="1dafa-945">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="1dafa-945">The following code removes an attachment with an identifier of '0'.</span></span>

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
