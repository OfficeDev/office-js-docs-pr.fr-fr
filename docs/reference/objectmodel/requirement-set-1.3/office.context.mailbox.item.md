---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: dfea6e86372c2a4310b5fa458ff3dcede7c6a184
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696280"
---
# <a name="item"></a><span data-ttu-id="da42a-102">élément</span><span class="sxs-lookup"><span data-stu-id="da42a-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="da42a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="da42a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="da42a-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="da42a-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-106">Requirements</span></span>

|<span data-ttu-id="da42a-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-107">Requirement</span></span>| <span data-ttu-id="da42a-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-110">1.0</span></span>|
|[<span data-ttu-id="da42a-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="da42a-112">Restricted</span></span>|
|[<span data-ttu-id="da42a-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="da42a-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="da42a-115">Members and methods</span></span>

| <span data-ttu-id="da42a-116">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-116">Member</span></span> | <span data-ttu-id="da42a-117">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="da42a-118">attachments</span><span class="sxs-lookup"><span data-stu-id="da42a-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="da42a-119">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-119">Member</span></span> |
| [<span data-ttu-id="da42a-120">bcc</span><span class="sxs-lookup"><span data-stu-id="da42a-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="da42a-121">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-121">Member</span></span> |
| [<span data-ttu-id="da42a-122">body</span><span class="sxs-lookup"><span data-stu-id="da42a-122">body</span></span>](#body-body) | <span data-ttu-id="da42a-123">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-123">Member</span></span> |
| [<span data-ttu-id="da42a-124">cc</span><span class="sxs-lookup"><span data-stu-id="da42a-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="da42a-125">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-125">Member</span></span> |
| [<span data-ttu-id="da42a-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="da42a-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="da42a-127">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-127">Member</span></span> |
| [<span data-ttu-id="da42a-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="da42a-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="da42a-129">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-129">Member</span></span> |
| [<span data-ttu-id="da42a-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="da42a-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="da42a-131">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-131">Member</span></span> |
| [<span data-ttu-id="da42a-132">end</span><span class="sxs-lookup"><span data-stu-id="da42a-132">end</span></span>](#end-datetime) | <span data-ttu-id="da42a-133">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-133">Member</span></span> |
| [<span data-ttu-id="da42a-134">from</span><span class="sxs-lookup"><span data-stu-id="da42a-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="da42a-135">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-135">Member</span></span> |
| [<span data-ttu-id="da42a-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="da42a-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="da42a-137">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-137">Member</span></span> |
| [<span data-ttu-id="da42a-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="da42a-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="da42a-139">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-139">Member</span></span> |
| [<span data-ttu-id="da42a-140">itemId</span><span class="sxs-lookup"><span data-stu-id="da42a-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="da42a-141">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-141">Member</span></span> |
| [<span data-ttu-id="da42a-142">itemType</span><span class="sxs-lookup"><span data-stu-id="da42a-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="da42a-143">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-143">Member</span></span> |
| [<span data-ttu-id="da42a-144">location</span><span class="sxs-lookup"><span data-stu-id="da42a-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="da42a-145">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-145">Member</span></span> |
| [<span data-ttu-id="da42a-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="da42a-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="da42a-147">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-147">Member</span></span> |
| [<span data-ttu-id="da42a-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="da42a-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="da42a-149">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-149">Member</span></span> |
| [<span data-ttu-id="da42a-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="da42a-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="da42a-151">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-151">Member</span></span> |
| [<span data-ttu-id="da42a-152">organizer</span><span class="sxs-lookup"><span data-stu-id="da42a-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="da42a-153">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-153">Member</span></span> |
| [<span data-ttu-id="da42a-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="da42a-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="da42a-155">Member</span><span class="sxs-lookup"><span data-stu-id="da42a-155">Member</span></span> |
| [<span data-ttu-id="da42a-156">sender</span><span class="sxs-lookup"><span data-stu-id="da42a-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="da42a-157">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-157">Member</span></span> |
| [<span data-ttu-id="da42a-158">start</span><span class="sxs-lookup"><span data-stu-id="da42a-158">start</span></span>](#start-datetime) | <span data-ttu-id="da42a-159">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-159">Member</span></span> |
| [<span data-ttu-id="da42a-160">subject</span><span class="sxs-lookup"><span data-stu-id="da42a-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="da42a-161">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-161">Member</span></span> |
| [<span data-ttu-id="da42a-162">to</span><span class="sxs-lookup"><span data-stu-id="da42a-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="da42a-163">Membre</span><span class="sxs-lookup"><span data-stu-id="da42a-163">Member</span></span> |
| [<span data-ttu-id="da42a-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="da42a-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-165">Method</span></span> |
| [<span data-ttu-id="da42a-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="da42a-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-167">Method</span></span> |
| [<span data-ttu-id="da42a-168">close</span><span class="sxs-lookup"><span data-stu-id="da42a-168">close</span></span>](#close) | <span data-ttu-id="da42a-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-169">Method</span></span> |
| [<span data-ttu-id="da42a-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="da42a-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="da42a-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-171">Method</span></span> |
| [<span data-ttu-id="da42a-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="da42a-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="da42a-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-173">Method</span></span> |
| [<span data-ttu-id="da42a-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="da42a-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="da42a-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-175">Method</span></span> |
| [<span data-ttu-id="da42a-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="da42a-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="da42a-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-177">Method</span></span> |
| [<span data-ttu-id="da42a-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="da42a-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="da42a-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-179">Method</span></span> |
| [<span data-ttu-id="da42a-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="da42a-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="da42a-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-181">Method</span></span> |
| [<span data-ttu-id="da42a-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="da42a-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="da42a-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-183">Method</span></span> |
| [<span data-ttu-id="da42a-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="da42a-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-185">Method</span></span> |
| [<span data-ttu-id="da42a-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="da42a-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-187">Method</span></span> |
| [<span data-ttu-id="da42a-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="da42a-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-189">Method</span></span> |
| [<span data-ttu-id="da42a-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="da42a-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-191">Method</span></span> |
| [<span data-ttu-id="da42a-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="da42a-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="da42a-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="da42a-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="da42a-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-194">Example</span></span>

<span data-ttu-id="da42a-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="da42a-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="da42a-196">Membres</span><span class="sxs-lookup"><span data-stu-id="da42a-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-13"></a><span data-ttu-id="da42a-197">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="da42a-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

<span data-ttu-id="da42a-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="da42a-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="da42a-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="da42a-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-202">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-202">Type</span></span>

*   <span data-ttu-id="da42a-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="da42a-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-204">Requirements</span></span>

|<span data-ttu-id="da42a-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-205">Requirement</span></span>| <span data-ttu-id="da42a-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-208">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-208">1.0</span></span>|
|[<span data-ttu-id="da42a-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-210">ReadItem</span></span>|
|[<span data-ttu-id="da42a-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-213">Example</span></span>

<span data-ttu-id="da42a-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="da42a-215">CCI: [destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="da42a-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="da42a-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-218">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-218">Type</span></span>

*   [<span data-ttu-id="da42a-219">Destinataires</span><span class="sxs-lookup"><span data-stu-id="da42a-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-220">Requirements</span></span>

|<span data-ttu-id="da42a-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-221">Requirement</span></span>| <span data-ttu-id="da42a-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-224">1.1</span><span class="sxs-lookup"><span data-stu-id="da42a-224">1.1</span></span>|
|[<span data-ttu-id="da42a-225">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-226">ReadItem</span></span>|
|[<span data-ttu-id="da42a-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-228">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-13"></a><span data-ttu-id="da42a-230">Body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-231">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-232">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-232">Type</span></span>

*   [<span data-ttu-id="da42a-233">Body</span><span class="sxs-lookup"><span data-stu-id="da42a-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-234">Requirements</span></span>

|<span data-ttu-id="da42a-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-235">Requirement</span></span>| <span data-ttu-id="da42a-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-238">1.1</span><span class="sxs-lookup"><span data-stu-id="da42a-238">1.1</span></span>|
|[<span data-ttu-id="da42a-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-240">ReadItem</span></span>|
|[<span data-ttu-id="da42a-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-242">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-243">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-243">Example</span></span>

<span data-ttu-id="da42a-244">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="da42a-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="da42a-245">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="da42a-246">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-247">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="da42a-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="da42a-248">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-249">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-249">Read mode</span></span>

<span data-ttu-id="da42a-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="da42a-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-252">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-252">Compose mode</span></span>

<span data-ttu-id="da42a-253">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="da42a-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

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

##### <a name="type"></a><span data-ttu-id="da42a-254">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-254">Type</span></span>

*   <span data-ttu-id="da42a-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-256">Requirements</span></span>

|<span data-ttu-id="da42a-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-257">Requirement</span></span>| <span data-ttu-id="da42a-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-260">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-260">1.0</span></span>|
|[<span data-ttu-id="da42a-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-262">ReadItem</span></span>|
|[<span data-ttu-id="da42a-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="da42a-265">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="da42a-266">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="da42a-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="da42a-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="da42a-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="da42a-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="da42a-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-271">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-271">Type</span></span>

*   <span data-ttu-id="da42a-272">String</span><span class="sxs-lookup"><span data-stu-id="da42a-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-273">Requirements</span></span>

|<span data-ttu-id="da42a-274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-274">Requirement</span></span>| <span data-ttu-id="da42a-275">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-277">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-277">1.0</span></span>|
|[<span data-ttu-id="da42a-278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-279">ReadItem</span></span>|
|[<span data-ttu-id="da42a-280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-281">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-282">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="da42a-283">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="da42a-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="da42a-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-286">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-286">Type</span></span>

*   <span data-ttu-id="da42a-287">Date</span><span class="sxs-lookup"><span data-stu-id="da42a-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-288">Requirements</span></span>

|<span data-ttu-id="da42a-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-289">Requirement</span></span>| <span data-ttu-id="da42a-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-292">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-292">1.0</span></span>|
|[<span data-ttu-id="da42a-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-294">ReadItem</span></span>|
|[<span data-ttu-id="da42a-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-296">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="da42a-298">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="da42a-298">dateTimeModified: Date</span></span>

<span data-ttu-id="da42a-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-301">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-302">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-302">Type</span></span>

*   <span data-ttu-id="da42a-303">Date</span><span class="sxs-lookup"><span data-stu-id="da42a-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-304">Requirements</span></span>

|<span data-ttu-id="da42a-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-305">Requirement</span></span>| <span data-ttu-id="da42a-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-308">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-308">1.0</span></span>|
|[<span data-ttu-id="da42a-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-310">ReadItem</span></span>|
|[<span data-ttu-id="da42a-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="da42a-314">fin: date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-315">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="da42a-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="da42a-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-318">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-318">Read mode</span></span>

<span data-ttu-id="da42a-319">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="da42a-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-320">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-320">Compose mode</span></span>

<span data-ttu-id="da42a-321">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="da42a-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="da42a-322">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="da42a-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="da42a-323">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="da42a-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="da42a-324">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-324">Type</span></span>

*   <span data-ttu-id="da42a-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-326">Requirements</span></span>

|<span data-ttu-id="da42a-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-327">Requirement</span></span>| <span data-ttu-id="da42a-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-330">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-330">1.0</span></span>|
|[<span data-ttu-id="da42a-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-332">ReadItem</span></span>|
|[<span data-ttu-id="da42a-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-334">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="da42a-335">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="da42a-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="da42a-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-340">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="da42a-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-341">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-341">Type</span></span>

*   [<span data-ttu-id="da42a-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="da42a-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-343">Requirements</span></span>

|<span data-ttu-id="da42a-344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-344">Requirement</span></span>| <span data-ttu-id="da42a-345">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-347">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-347">1.0</span></span>|
|[<span data-ttu-id="da42a-348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-349">ReadItem</span></span>|
|[<span data-ttu-id="da42a-350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-351">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-352">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="da42a-353">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-353">internetMessageId: String</span></span>

<span data-ttu-id="da42a-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-356">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-356">Type</span></span>

*   <span data-ttu-id="da42a-357">String</span><span class="sxs-lookup"><span data-stu-id="da42a-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-358">Requirements</span></span>

|<span data-ttu-id="da42a-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-359">Requirement</span></span>| <span data-ttu-id="da42a-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-362">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-362">1.0</span></span>|
|[<span data-ttu-id="da42a-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-364">ReadItem</span></span>|
|[<span data-ttu-id="da42a-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="da42a-368">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-368">itemClass: String</span></span>

<span data-ttu-id="da42a-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="da42a-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="da42a-373">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-373">Type</span></span> | <span data-ttu-id="da42a-374">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-374">Description</span></span> | <span data-ttu-id="da42a-375">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="da42a-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="da42a-376">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="da42a-376">Appointment items</span></span> | <span data-ttu-id="da42a-377">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="da42a-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="da42a-378">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="da42a-378">Message items</span></span> | <span data-ttu-id="da42a-379">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="da42a-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="da42a-380">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="da42a-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-381">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-381">Type</span></span>

*   <span data-ttu-id="da42a-382">String</span><span class="sxs-lookup"><span data-stu-id="da42a-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-383">Requirements</span></span>

|<span data-ttu-id="da42a-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-384">Requirement</span></span>| <span data-ttu-id="da42a-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-387">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-387">1.0</span></span>|
|[<span data-ttu-id="da42a-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-389">ReadItem</span></span>|
|[<span data-ttu-id="da42a-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="da42a-393">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="da42a-393">(nullable) itemId: String</span></span>

<span data-ttu-id="da42a-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-396">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="da42a-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="da42a-397">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="da42a-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="da42a-398">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="da42a-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="da42a-399">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="da42a-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="da42a-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-402">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-402">Type</span></span>

*   <span data-ttu-id="da42a-403">String</span><span class="sxs-lookup"><span data-stu-id="da42a-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-404">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-404">Requirements</span></span>

|<span data-ttu-id="da42a-405">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-405">Requirement</span></span>| <span data-ttu-id="da42a-406">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-407">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-408">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-408">1.0</span></span>|
|[<span data-ttu-id="da42a-409">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-410">ReadItem</span></span>|
|[<span data-ttu-id="da42a-411">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-412">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-413">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-413">Example</span></span>

<span data-ttu-id="da42a-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="da42a-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-13"></a><span data-ttu-id="da42a-416">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-417">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="da42a-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="da42a-418">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-419">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-419">Type</span></span>

*   [<span data-ttu-id="da42a-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="da42a-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-421">Requirements</span></span>

|<span data-ttu-id="da42a-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-422">Requirement</span></span>| <span data-ttu-id="da42a-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-425">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-425">1.0</span></span>|
|[<span data-ttu-id="da42a-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-427">ReadItem</span></span>|
|[<span data-ttu-id="da42a-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-429">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-13"></a><span data-ttu-id="da42a-431">Location: String | [Emplacement](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-432">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-433">Read mode</span></span>

<span data-ttu-id="da42a-434">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-435">Compose mode</span></span>

<span data-ttu-id="da42a-436">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da42a-437">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-437">Type</span></span>

*   <span data-ttu-id="da42a-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-439">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-439">Requirements</span></span>

|<span data-ttu-id="da42a-440">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-440">Requirement</span></span>| <span data-ttu-id="da42a-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-442">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-443">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-443">1.0</span></span>|
|[<span data-ttu-id="da42a-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-445">ReadItem</span></span>|
|[<span data-ttu-id="da42a-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-447">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="da42a-448">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-448">normalizedSubject: String</span></span>

<span data-ttu-id="da42a-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="da42a-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="da42a-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-453">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-453">Type</span></span>

*   <span data-ttu-id="da42a-454">String</span><span class="sxs-lookup"><span data-stu-id="da42a-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-455">Requirements</span></span>

|<span data-ttu-id="da42a-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-456">Requirement</span></span>| <span data-ttu-id="da42a-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-459">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-459">1.0</span></span>|
|[<span data-ttu-id="da42a-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-461">ReadItem</span></span>|
|[<span data-ttu-id="da42a-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-463">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-464">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-13"></a><span data-ttu-id="da42a-465">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-466">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-467">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-467">Type</span></span>

*   [<span data-ttu-id="da42a-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="da42a-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-469">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-469">Requirements</span></span>

|<span data-ttu-id="da42a-470">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-470">Requirement</span></span>| <span data-ttu-id="da42a-471">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-472">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-473">1.3</span><span class="sxs-lookup"><span data-stu-id="da42a-473">1.3</span></span>|
|[<span data-ttu-id="da42a-474">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-475">ReadItem</span></span>|
|[<span data-ttu-id="da42a-476">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-477">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-478">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="da42a-479">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.3) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="da42a-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-480">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="da42a-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="da42a-481">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-482">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-482">Read mode</span></span>

<span data-ttu-id="da42a-483">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="da42a-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-484">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-484">Compose mode</span></span>

<span data-ttu-id="da42a-485">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="da42a-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da42a-486">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-486">Type</span></span>

*   <span data-ttu-id="da42a-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-488">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-488">Requirements</span></span>

|<span data-ttu-id="da42a-489">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-489">Requirement</span></span>| <span data-ttu-id="da42a-490">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-491">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-492">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-492">1.0</span></span>|
|[<span data-ttu-id="da42a-493">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-494">ReadItem</span></span>|
|[<span data-ttu-id="da42a-495">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-496">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="da42a-497">Organisateur: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-500">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-500">Type</span></span>

*   [<span data-ttu-id="da42a-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="da42a-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-502">Requirements</span></span>

|<span data-ttu-id="da42a-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-503">Requirement</span></span>| <span data-ttu-id="da42a-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-506">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-506">1.0</span></span>|
|[<span data-ttu-id="da42a-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-508">ReadItem</span></span>|
|[<span data-ttu-id="da42a-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-510">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-511">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="da42a-512">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.3) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="da42a-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-513">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="da42a-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="da42a-514">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-515">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-515">Read mode</span></span>

<span data-ttu-id="da42a-516">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="da42a-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-517">Compose mode</span></span>

<span data-ttu-id="da42a-518">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="da42a-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="da42a-519">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-519">Type</span></span>

*   <span data-ttu-id="da42a-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-521">Requirements</span></span>

|<span data-ttu-id="da42a-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-522">Requirement</span></span>| <span data-ttu-id="da42a-523">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-525">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-525">1.0</span></span>|
|[<span data-ttu-id="da42a-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-527">ReadItem</span></span>|
|[<span data-ttu-id="da42a-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-529">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="da42a-530">expéditeur: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="da42a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="da42a-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="da42a-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-535">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="da42a-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="da42a-536">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-536">Type</span></span>

*   [<span data-ttu-id="da42a-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="da42a-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="da42a-538">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-538">Requirements</span></span>

|<span data-ttu-id="da42a-539">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-539">Requirement</span></span>| <span data-ttu-id="da42a-540">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-541">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-542">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-542">1.0</span></span>|
|[<span data-ttu-id="da42a-543">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-544">ReadItem</span></span>|
|[<span data-ttu-id="da42a-545">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-546">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-547">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="da42a-548">début: date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-549">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="da42a-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="da42a-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-552">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-552">Read mode</span></span>

<span data-ttu-id="da42a-553">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="da42a-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-554">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-554">Compose mode</span></span>

<span data-ttu-id="da42a-555">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="da42a-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="da42a-556">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="da42a-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="da42a-557">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="da42a-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="da42a-558">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-558">Type</span></span>

*   <span data-ttu-id="da42a-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-560">Requirements</span></span>

|<span data-ttu-id="da42a-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-561">Requirement</span></span>| <span data-ttu-id="da42a-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-564">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-564">1.0</span></span>|
|[<span data-ttu-id="da42a-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-566">ReadItem</span></span>|
|[<span data-ttu-id="da42a-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-13"></a><span data-ttu-id="da42a-569">Subject: String | [Objet](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="da42a-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="da42a-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-572">Read mode</span></span>

<span data-ttu-id="da42a-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="da42a-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-575">Compose mode</span></span>

<span data-ttu-id="da42a-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="da42a-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="da42a-577">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-577">Type</span></span>

*   <span data-ttu-id="da42a-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-579">Requirements</span></span>

|<span data-ttu-id="da42a-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-580">Requirement</span></span>| <span data-ttu-id="da42a-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-583">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-583">1.0</span></span>|
|[<span data-ttu-id="da42a-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-585">ReadItem</span></span>|
|[<span data-ttu-id="da42a-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-587">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="da42a-588">to: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="da42a-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="da42a-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="da42a-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da42a-591">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-591">Read mode</span></span>

<span data-ttu-id="da42a-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="da42a-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="da42a-594">Mode composition</span><span class="sxs-lookup"><span data-stu-id="da42a-594">Compose mode</span></span>

<span data-ttu-id="da42a-595">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="da42a-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da42a-596">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-596">Type</span></span>

*   <span data-ttu-id="da42a-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-598">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-598">Requirements</span></span>

|<span data-ttu-id="da42a-599">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-599">Requirement</span></span>| <span data-ttu-id="da42a-600">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-601">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-602">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-602">1.0</span></span>|
|[<span data-ttu-id="da42a-603">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-604">ReadItem</span></span>|
|[<span data-ttu-id="da42a-605">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-606">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="da42a-607">Méthodes</span><span class="sxs-lookup"><span data-stu-id="da42a-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="da42a-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da42a-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="da42a-609">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="da42a-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="da42a-610">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="da42a-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="da42a-611">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="da42a-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-612">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-612">Parameters</span></span>

|<span data-ttu-id="da42a-613">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-613">Name</span></span>| <span data-ttu-id="da42a-614">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-614">Type</span></span>| <span data-ttu-id="da42a-615">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-615">Attributes</span></span>| <span data-ttu-id="da42a-616">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="da42a-617">String</span><span class="sxs-lookup"><span data-stu-id="da42a-617">String</span></span>||<span data-ttu-id="da42a-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="da42a-620">String</span><span class="sxs-lookup"><span data-stu-id="da42a-620">String</span></span>||<span data-ttu-id="da42a-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="da42a-623">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-623">Object</span></span>| <span data-ttu-id="da42a-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-624">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="da42a-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="da42a-626">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-626">Object</span></span>| <span data-ttu-id="da42a-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-627">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="da42a-629">fonction</span><span class="sxs-lookup"><span data-stu-id="da42a-629">function</span></span>| <span data-ttu-id="da42a-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-630">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-631">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="da42a-632">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da42a-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="da42a-633">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="da42a-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="da42a-634">Erreurs</span><span class="sxs-lookup"><span data-stu-id="da42a-634">Errors</span></span>

| <span data-ttu-id="da42a-635">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="da42a-635">Error code</span></span> | <span data-ttu-id="da42a-636">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="da42a-637">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="da42a-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="da42a-638">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="da42a-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="da42a-639">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="da42a-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="da42a-640">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-640">Requirements</span></span>

|<span data-ttu-id="da42a-641">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-641">Requirement</span></span>| <span data-ttu-id="da42a-642">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-643">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-644">1.1</span><span class="sxs-lookup"><span data-stu-id="da42a-644">1.1</span></span>|
|[<span data-ttu-id="da42a-645">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da42a-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="da42a-647">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-648">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-649">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-649">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="da42a-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da42a-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="da42a-651">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="da42a-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="da42a-655">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="da42a-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="da42a-656">Si votre complément Office est en cours d’exécution dans Outlook sur le Web, `addItemAttachmentAsync` la méthode peut joindre des éléments à des éléments autres que l’élément que vous modifiez; Toutefois, cette option n’est pas prise en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="da42a-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-657">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-657">Parameters</span></span>

|<span data-ttu-id="da42a-658">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-658">Name</span></span>| <span data-ttu-id="da42a-659">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-659">Type</span></span>| <span data-ttu-id="da42a-660">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-660">Attributes</span></span>| <span data-ttu-id="da42a-661">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="da42a-662">Chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-662">String</span></span>||<span data-ttu-id="da42a-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="da42a-665">String</span><span class="sxs-lookup"><span data-stu-id="da42a-665">String</span></span>||<span data-ttu-id="da42a-666">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="da42a-666">The subject of the item to be attached.</span></span> <span data-ttu-id="da42a-667">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="da42a-668">Object</span><span class="sxs-lookup"><span data-stu-id="da42a-668">Object</span></span>| <span data-ttu-id="da42a-669">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-669">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-670">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="da42a-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="da42a-671">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-671">Object</span></span>| <span data-ttu-id="da42a-672">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-672">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-673">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="da42a-674">fonction</span><span class="sxs-lookup"><span data-stu-id="da42a-674">function</span></span>| <span data-ttu-id="da42a-675">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-675">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-676">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="da42a-677">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da42a-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="da42a-678">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="da42a-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="da42a-679">Erreurs</span><span class="sxs-lookup"><span data-stu-id="da42a-679">Errors</span></span>

| <span data-ttu-id="da42a-680">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="da42a-680">Error code</span></span> | <span data-ttu-id="da42a-681">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="da42a-682">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="da42a-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="da42a-683">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-683">Requirements</span></span>

|<span data-ttu-id="da42a-684">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-684">Requirement</span></span>| <span data-ttu-id="da42a-685">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-686">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-687">1.1</span><span class="sxs-lookup"><span data-stu-id="da42a-687">1.1</span></span>|
|[<span data-ttu-id="da42a-688">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da42a-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="da42a-690">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-691">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-692">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-692">Example</span></span>

<span data-ttu-id="da42a-693">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="da42a-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="da42a-694">close()</span><span class="sxs-lookup"><span data-stu-id="da42a-694">close()</span></span>

<span data-ttu-id="da42a-695">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="da42a-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="da42a-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="da42a-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-698">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="da42a-699">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="da42a-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-700">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-700">Requirements</span></span>

|<span data-ttu-id="da42a-701">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-701">Requirement</span></span>| <span data-ttu-id="da42a-702">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-703">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-704">1.3</span><span class="sxs-lookup"><span data-stu-id="da42a-704">1.3</span></span>|
|[<span data-ttu-id="da42a-705">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-706">Restreinte</span><span class="sxs-lookup"><span data-stu-id="da42a-706">Restricted</span></span>|
|[<span data-ttu-id="da42a-707">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-708">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-708">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="da42a-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="da42a-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="da42a-710">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="da42a-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-711">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="da42a-712">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="da42a-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="da42a-713">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="da42a-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="da42a-714">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="da42a-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="da42a-715">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="da42a-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="da42a-716">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="da42a-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-717">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-717">Parameters</span></span>

|<span data-ttu-id="da42a-718">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-718">Name</span></span>| <span data-ttu-id="da42a-719">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-719">Type</span></span>| <span data-ttu-id="da42a-720">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="da42a-721">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="da42a-721">String &#124; Object</span></span>| |<span data-ttu-id="da42a-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="da42a-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="da42a-724">**OU**</span><span class="sxs-lookup"><span data-stu-id="da42a-724">**OR**</span></span><br/><span data-ttu-id="da42a-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="da42a-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="da42a-727">String</span><span class="sxs-lookup"><span data-stu-id="da42a-727">String</span></span> | <span data-ttu-id="da42a-728">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-728">&lt;optional&gt;</span></span> | <span data-ttu-id="da42a-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="da42a-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="da42a-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="da42a-732">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-732">&lt;optional&gt;</span></span> | <span data-ttu-id="da42a-733">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="da42a-734">String</span><span class="sxs-lookup"><span data-stu-id="da42a-734">String</span></span> | | <span data-ttu-id="da42a-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="da42a-737">String</span><span class="sxs-lookup"><span data-stu-id="da42a-737">String</span></span> | | <span data-ttu-id="da42a-738">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="da42a-739">Chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-739">String</span></span> | | <span data-ttu-id="da42a-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="da42a-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="da42a-742">String</span><span class="sxs-lookup"><span data-stu-id="da42a-742">String</span></span> | | <span data-ttu-id="da42a-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="da42a-746">function</span><span class="sxs-lookup"><span data-stu-id="da42a-746">function</span></span> | <span data-ttu-id="da42a-747">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-747">&lt;optional&gt;</span></span> | <span data-ttu-id="da42a-748">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="da42a-749">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-749">Requirements</span></span>

|<span data-ttu-id="da42a-750">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-750">Requirement</span></span>| <span data-ttu-id="da42a-751">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-752">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-753">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-753">1.0</span></span>|
|[<span data-ttu-id="da42a-754">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-755">ReadItem</span></span>|
|[<span data-ttu-id="da42a-756">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-757">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="da42a-758">Exemples</span><span class="sxs-lookup"><span data-stu-id="da42a-758">Examples</span></span>

<span data-ttu-id="da42a-759">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="da42a-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="da42a-760">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="da42a-760">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="da42a-761">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="da42a-761">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="da42a-762">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="da42a-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="da42a-763">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="da42a-764">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="da42a-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="da42a-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="da42a-766">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="da42a-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-767">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="da42a-768">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="da42a-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="da42a-769">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="da42a-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="da42a-770">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="da42a-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="da42a-771">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="da42a-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="da42a-772">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="da42a-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-773">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-773">Parameters</span></span>

|<span data-ttu-id="da42a-774">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-774">Name</span></span>| <span data-ttu-id="da42a-775">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-775">Type</span></span>| <span data-ttu-id="da42a-776">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="da42a-777">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="da42a-777">String &#124; Object</span></span>| | <span data-ttu-id="da42a-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="da42a-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="da42a-780">**OU**</span><span class="sxs-lookup"><span data-stu-id="da42a-780">**OR**</span></span><br/><span data-ttu-id="da42a-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="da42a-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="da42a-783">String</span><span class="sxs-lookup"><span data-stu-id="da42a-783">String</span></span> | <span data-ttu-id="da42a-784">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-784">&lt;optional&gt;</span></span> | <span data-ttu-id="da42a-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="da42a-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="da42a-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="da42a-788">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-788">&lt;optional&gt;</span></span> | <span data-ttu-id="da42a-789">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="da42a-790">String</span><span class="sxs-lookup"><span data-stu-id="da42a-790">String</span></span> | | <span data-ttu-id="da42a-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="da42a-793">String</span><span class="sxs-lookup"><span data-stu-id="da42a-793">String</span></span> | | <span data-ttu-id="da42a-794">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="da42a-795">Chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-795">String</span></span> | | <span data-ttu-id="da42a-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="da42a-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="da42a-798">String</span><span class="sxs-lookup"><span data-stu-id="da42a-798">String</span></span> | | <span data-ttu-id="da42a-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="da42a-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="da42a-802">function</span><span class="sxs-lookup"><span data-stu-id="da42a-802">function</span></span> | <span data-ttu-id="da42a-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-803">&lt;optional&gt;</span></span> | <span data-ttu-id="da42a-804">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="da42a-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-805">Requirements</span></span>

|<span data-ttu-id="da42a-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-806">Requirement</span></span>| <span data-ttu-id="da42a-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-809">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-809">1.0</span></span>|
|[<span data-ttu-id="da42a-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-811">ReadItem</span></span>|
|[<span data-ttu-id="da42a-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-813">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="da42a-814">Exemples</span><span class="sxs-lookup"><span data-stu-id="da42a-814">Examples</span></span>

<span data-ttu-id="da42a-815">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="da42a-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="da42a-816">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="da42a-816">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="da42a-817">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="da42a-817">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="da42a-818">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="da42a-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="da42a-819">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="da42a-820">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-13"></a><span data-ttu-id="da42a-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="da42a-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="da42a-822">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="da42a-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-823">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-824">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-824">Requirements</span></span>

|<span data-ttu-id="da42a-825">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-825">Requirement</span></span>| <span data-ttu-id="da42a-826">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-827">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-828">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-828">1.0</span></span>|
|[<span data-ttu-id="da42a-829">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-830">ReadItem</span></span>|
|[<span data-ttu-id="da42a-831">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-832">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da42a-833">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="da42a-833">Returns:</span></span>

<span data-ttu-id="da42a-834">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="da42a-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span></span>

##### <a name="example"></a><span data-ttu-id="da42a-835">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-835">Example</span></span>

<span data-ttu-id="da42a-836">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-836">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="da42a-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="da42a-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="da42a-838">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="da42a-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-839">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-840">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-840">Parameters</span></span>

|<span data-ttu-id="da42a-841">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-841">Name</span></span>| <span data-ttu-id="da42a-842">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-842">Type</span></span>| <span data-ttu-id="da42a-843">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="da42a-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="da42a-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.3)|<span data-ttu-id="da42a-845">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="da42a-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da42a-846">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-846">Requirements</span></span>

|<span data-ttu-id="da42a-847">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-847">Requirement</span></span>| <span data-ttu-id="da42a-848">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-849">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-850">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-850">1.0</span></span>|
|[<span data-ttu-id="da42a-851">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-852">Restreinte</span><span class="sxs-lookup"><span data-stu-id="da42a-852">Restricted</span></span>|
|[<span data-ttu-id="da42a-853">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-854">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da42a-855">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="da42a-855">Returns:</span></span>

<span data-ttu-id="da42a-856">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="da42a-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="da42a-857">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="da42a-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="da42a-858">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="da42a-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="da42a-859">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="da42a-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="da42a-860">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="da42a-860">Value of `entityType`</span></span> | <span data-ttu-id="da42a-861">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="da42a-861">Type of objects in returned array</span></span> | <span data-ttu-id="da42a-862">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="da42a-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="da42a-863">String</span><span class="sxs-lookup"><span data-stu-id="da42a-863">String</span></span> | <span data-ttu-id="da42a-864">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="da42a-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="da42a-865">Contact</span><span class="sxs-lookup"><span data-stu-id="da42a-865">Contact</span></span> | <span data-ttu-id="da42a-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da42a-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="da42a-867">String</span><span class="sxs-lookup"><span data-stu-id="da42a-867">String</span></span> | <span data-ttu-id="da42a-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da42a-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="da42a-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="da42a-869">MeetingSuggestion</span></span> | <span data-ttu-id="da42a-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da42a-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="da42a-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="da42a-871">PhoneNumber</span></span> | <span data-ttu-id="da42a-872">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="da42a-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="da42a-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="da42a-873">TaskSuggestion</span></span> | <span data-ttu-id="da42a-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da42a-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="da42a-875">String</span><span class="sxs-lookup"><span data-stu-id="da42a-875">String</span></span> | <span data-ttu-id="da42a-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="da42a-876">**Restricted**</span></span> |

<span data-ttu-id="da42a-877">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="da42a-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

##### <a name="example"></a><span data-ttu-id="da42a-878">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-878">Example</span></span>

<span data-ttu-id="da42a-879">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="da42a-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="da42a-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="da42a-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="da42a-881">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="da42a-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-882">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="da42a-883">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="da42a-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-884">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-884">Parameters</span></span>

|<span data-ttu-id="da42a-885">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-885">Name</span></span>| <span data-ttu-id="da42a-886">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-886">Type</span></span>| <span data-ttu-id="da42a-887">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="da42a-888">Chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-888">String</span></span>|<span data-ttu-id="da42a-889">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="da42a-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da42a-890">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-890">Requirements</span></span>

|<span data-ttu-id="da42a-891">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-891">Requirement</span></span>| <span data-ttu-id="da42a-892">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-893">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-894">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-894">1.0</span></span>|
|[<span data-ttu-id="da42a-895">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-896">ReadItem</span></span>|
|[<span data-ttu-id="da42a-897">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-898">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da42a-899">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="da42a-899">Returns:</span></span>

<span data-ttu-id="da42a-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="da42a-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="da42a-902">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="da42a-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="da42a-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="da42a-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="da42a-904">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="da42a-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-905">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="da42a-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="da42a-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="da42a-909">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="da42a-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="da42a-910">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="da42a-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="da42a-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da42a-914">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-914">Requirements</span></span>

|<span data-ttu-id="da42a-915">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-915">Requirement</span></span>| <span data-ttu-id="da42a-916">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-917">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-918">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-918">1.0</span></span>|
|[<span data-ttu-id="da42a-919">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-920">ReadItem</span></span>|
|[<span data-ttu-id="da42a-921">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-922">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da42a-923">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="da42a-923">Returns:</span></span>

<span data-ttu-id="da42a-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="da42a-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="da42a-926">Type: objet</span><span class="sxs-lookup"><span data-stu-id="da42a-926">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="da42a-927">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-927">Example</span></span>

<span data-ttu-id="da42a-928">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="da42a-928">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="da42a-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="da42a-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="da42a-930">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="da42a-930">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-931">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="da42a-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="da42a-932">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="da42a-932">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="da42a-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="da42a-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-935">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-935">Parameters</span></span>

|<span data-ttu-id="da42a-936">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-936">Name</span></span>| <span data-ttu-id="da42a-937">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-937">Type</span></span>| <span data-ttu-id="da42a-938">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-938">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="da42a-939">Chaîne</span><span class="sxs-lookup"><span data-stu-id="da42a-939">String</span></span>|<span data-ttu-id="da42a-940">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="da42a-940">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da42a-941">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-941">Requirements</span></span>

|<span data-ttu-id="da42a-942">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-942">Requirement</span></span>| <span data-ttu-id="da42a-943">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-944">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-945">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-945">1.0</span></span>|
|[<span data-ttu-id="da42a-946">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-947">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-947">ReadItem</span></span>|
|[<span data-ttu-id="da42a-948">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-949">Lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-949">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da42a-950">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="da42a-950">Returns:</span></span>

<span data-ttu-id="da42a-951">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="da42a-951">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="da42a-952">Type: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="da42a-952">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="da42a-953">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-953">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="da42a-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="da42a-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="da42a-955">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="da42a-955">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="da42a-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="da42a-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-958">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-958">Parameters</span></span>

|<span data-ttu-id="da42a-959">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-959">Name</span></span>| <span data-ttu-id="da42a-960">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-960">Type</span></span>| <span data-ttu-id="da42a-961">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-961">Attributes</span></span>| <span data-ttu-id="da42a-962">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-962">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="da42a-963">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="da42a-963">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="da42a-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="da42a-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="da42a-967">Object</span><span class="sxs-lookup"><span data-stu-id="da42a-967">Object</span></span>| <span data-ttu-id="da42a-968">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-968">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-969">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="da42a-969">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="da42a-970">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-970">Object</span></span>| <span data-ttu-id="da42a-971">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-971">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-972">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-972">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="da42a-973">fonction</span><span class="sxs-lookup"><span data-stu-id="da42a-973">function</span></span>||<span data-ttu-id="da42a-974">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-974">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="da42a-975">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="da42a-975">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="da42a-976">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="da42a-976">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da42a-977">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-977">Requirements</span></span>

|<span data-ttu-id="da42a-978">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-978">Requirement</span></span>| <span data-ttu-id="da42a-979">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-979">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-980">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-980">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-981">1.2</span><span class="sxs-lookup"><span data-stu-id="da42a-981">1.2</span></span>|
|[<span data-ttu-id="da42a-982">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-982">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-983">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da42a-983">ReadWriteItem</span></span>|
|[<span data-ttu-id="da42a-984">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-984">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-985">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-985">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="da42a-986">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="da42a-986">Returns:</span></span>

<span data-ttu-id="da42a-987">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="da42a-987">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="da42a-988">Type : String</span><span class="sxs-lookup"><span data-stu-id="da42a-988">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="da42a-989">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-989">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="da42a-990">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="da42a-990">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="da42a-991">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="da42a-991">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="da42a-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="da42a-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-995">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-995">Parameters</span></span>

|<span data-ttu-id="da42a-996">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-996">Name</span></span>| <span data-ttu-id="da42a-997">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-997">Type</span></span>| <span data-ttu-id="da42a-998">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-998">Attributes</span></span>| <span data-ttu-id="da42a-999">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-999">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="da42a-1000">function</span><span class="sxs-lookup"><span data-stu-id="da42a-1000">function</span></span>||<span data-ttu-id="da42a-1001">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-1001">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="da42a-1002">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da42a-1002">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="da42a-1003">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="da42a-1003">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="da42a-1004">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-1004">Object</span></span>| <span data-ttu-id="da42a-1005">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1005">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1006">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-1006">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="da42a-1007">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-1007">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da42a-1008">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-1008">Requirements</span></span>

|<span data-ttu-id="da42a-1009">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-1009">Requirement</span></span>| <span data-ttu-id="da42a-1010">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-1010">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-1011">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-1011">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-1012">1.0</span><span class="sxs-lookup"><span data-stu-id="da42a-1012">1.0</span></span>|
|[<span data-ttu-id="da42a-1013">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-1013">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-1014">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da42a-1014">ReadItem</span></span>|
|[<span data-ttu-id="da42a-1015">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-1015">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-1016">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="da42a-1016">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-1017">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-1017">Example</span></span>

<span data-ttu-id="da42a-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="da42a-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="da42a-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da42a-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="da42a-1022">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="da42a-1022">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="da42a-1023">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-1023">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="da42a-1024">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="da42a-1024">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="da42a-1025">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="da42a-1025">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="da42a-1026">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="da42a-1026">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-1027">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-1027">Parameters</span></span>

|<span data-ttu-id="da42a-1028">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-1028">Name</span></span>| <span data-ttu-id="da42a-1029">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-1029">Type</span></span>| <span data-ttu-id="da42a-1030">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-1030">Attributes</span></span>| <span data-ttu-id="da42a-1031">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-1031">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="da42a-1032">String</span><span class="sxs-lookup"><span data-stu-id="da42a-1032">String</span></span>||<span data-ttu-id="da42a-1033">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="da42a-1033">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="da42a-1034">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-1034">Object</span></span>| <span data-ttu-id="da42a-1035">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1036">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="da42a-1036">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="da42a-1037">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-1037">Object</span></span>| <span data-ttu-id="da42a-1038">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1039">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-1039">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="da42a-1040">fonction</span><span class="sxs-lookup"><span data-stu-id="da42a-1040">function</span></span>| <span data-ttu-id="da42a-1041">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1042">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-1042">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="da42a-1043">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="da42a-1043">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="da42a-1044">Erreurs</span><span class="sxs-lookup"><span data-stu-id="da42a-1044">Errors</span></span>

| <span data-ttu-id="da42a-1045">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="da42a-1045">Error code</span></span> | <span data-ttu-id="da42a-1046">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-1046">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="da42a-1047">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="da42a-1047">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="da42a-1048">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-1048">Requirements</span></span>

|<span data-ttu-id="da42a-1049">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-1049">Requirement</span></span>| <span data-ttu-id="da42a-1050">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-1051">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-1052">1.1</span><span class="sxs-lookup"><span data-stu-id="da42a-1052">1.1</span></span>|
|[<span data-ttu-id="da42a-1053">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-1054">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da42a-1054">ReadWriteItem</span></span>|
|[<span data-ttu-id="da42a-1055">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-1056">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-1056">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-1057">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-1057">Example</span></span>

<span data-ttu-id="da42a-1058">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="da42a-1058">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="da42a-1059">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="da42a-1059">saveAsync([options], callback)</span></span>

<span data-ttu-id="da42a-1060">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="da42a-1060">Asynchronously saves an item.</span></span>

<span data-ttu-id="da42a-1061">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-1061">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="da42a-1062">Dans Outlook sur le Web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="da42a-1062">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="da42a-1063">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="da42a-1063">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-1064">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="da42a-1064">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="da42a-1065">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="da42a-1065">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="da42a-p168">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="da42a-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="da42a-1069">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="da42a-1069">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="da42a-1070">Outlook sur Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="da42a-1070">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="da42a-1071">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="da42a-1071">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="da42a-1072">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="da42a-1072">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="da42a-1073">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="da42a-1073">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-1074">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-1074">Parameters</span></span>

|<span data-ttu-id="da42a-1075">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-1075">Name</span></span>| <span data-ttu-id="da42a-1076">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-1076">Type</span></span>| <span data-ttu-id="da42a-1077">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-1077">Attributes</span></span>| <span data-ttu-id="da42a-1078">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-1078">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="da42a-1079">Object</span><span class="sxs-lookup"><span data-stu-id="da42a-1079">Object</span></span>| <span data-ttu-id="da42a-1080">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1080">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1081">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="da42a-1081">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="da42a-1082">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-1082">Object</span></span>| <span data-ttu-id="da42a-1083">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1084">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-1084">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="da42a-1085">fonction</span><span class="sxs-lookup"><span data-stu-id="da42a-1085">function</span></span>||<span data-ttu-id="da42a-1086">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="da42a-1087">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da42a-1087">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da42a-1088">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-1088">Requirements</span></span>

|<span data-ttu-id="da42a-1089">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-1089">Requirement</span></span>| <span data-ttu-id="da42a-1090">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-1091">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-1092">1.3</span><span class="sxs-lookup"><span data-stu-id="da42a-1092">1.3</span></span>|
|[<span data-ttu-id="da42a-1093">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-1093">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da42a-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="da42a-1095">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-1095">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-1096">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-1096">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="da42a-1097">範例</span><span class="sxs-lookup"><span data-stu-id="da42a-1097">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="da42a-p170">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="da42a-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="da42a-1100">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="da42a-1100">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="da42a-1101">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="da42a-1101">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="da42a-p171">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="da42a-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da42a-1105">Paramètres</span><span class="sxs-lookup"><span data-stu-id="da42a-1105">Parameters</span></span>

|<span data-ttu-id="da42a-1106">Nom</span><span class="sxs-lookup"><span data-stu-id="da42a-1106">Name</span></span>| <span data-ttu-id="da42a-1107">Type</span><span class="sxs-lookup"><span data-stu-id="da42a-1107">Type</span></span>| <span data-ttu-id="da42a-1108">Attributs</span><span class="sxs-lookup"><span data-stu-id="da42a-1108">Attributes</span></span>| <span data-ttu-id="da42a-1109">Description</span><span class="sxs-lookup"><span data-stu-id="da42a-1109">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="da42a-1110">String</span><span class="sxs-lookup"><span data-stu-id="da42a-1110">String</span></span>||<span data-ttu-id="da42a-p172">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="da42a-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="da42a-1114">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-1114">Object</span></span>| <span data-ttu-id="da42a-1115">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1115">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1116">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="da42a-1116">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="da42a-1117">Objet</span><span class="sxs-lookup"><span data-stu-id="da42a-1117">Object</span></span>| <span data-ttu-id="da42a-1118">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1119">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="da42a-1119">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="da42a-1120">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="da42a-1120">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="da42a-1121">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="da42a-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="da42a-1122">Si `text`, le style actuel est appliqué dans Outlook sur le Web et les clients de bureau.</span><span class="sxs-lookup"><span data-stu-id="da42a-1122">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="da42a-1123">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="da42a-1123">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="da42a-1124">Si `html` et que le champ prend en charge le format html (l’objet ne l’est pas), le style actuel est appliqué dans Outlook sur le Web et le style par défaut est appliqué dans les clients de bureau Outlook.</span><span class="sxs-lookup"><span data-stu-id="da42a-1124">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="da42a-1125">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="da42a-1125">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="da42a-1126">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="da42a-1126">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="da42a-1127">fonction</span><span class="sxs-lookup"><span data-stu-id="da42a-1127">function</span></span>||<span data-ttu-id="da42a-1128">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da42a-1128">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="da42a-1129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da42a-1129">Requirements</span></span>

|<span data-ttu-id="da42a-1130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="da42a-1130">Requirement</span></span>| <span data-ttu-id="da42a-1131">Valeur</span><span class="sxs-lookup"><span data-stu-id="da42a-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="da42a-1132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="da42a-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da42a-1133">1.2</span><span class="sxs-lookup"><span data-stu-id="da42a-1133">1.2</span></span>|
|[<span data-ttu-id="da42a-1134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="da42a-1134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da42a-1135">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da42a-1135">ReadWriteItem</span></span>|
|[<span data-ttu-id="da42a-1136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="da42a-1136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da42a-1137">Composition</span><span class="sxs-lookup"><span data-stu-id="da42a-1137">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da42a-1138">Exemple</span><span class="sxs-lookup"><span data-stu-id="da42a-1138">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
