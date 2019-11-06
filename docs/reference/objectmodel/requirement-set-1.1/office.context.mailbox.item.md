---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,1
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5cbf942ea9b1351e0f945a9ca5534a9ba090b79b
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001613"
---
# <a name="item"></a><span data-ttu-id="14ff7-102">élément</span><span class="sxs-lookup"><span data-stu-id="14ff7-102">item</span></span>

### <span data-ttu-id="14ff7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="14ff7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="14ff7-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="14ff7-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-107">Requirements</span></span>

|<span data-ttu-id="14ff7-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-108">Requirement</span></span>| <span data-ttu-id="14ff7-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-111">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-111">1.0</span></span>|
|[<span data-ttu-id="14ff7-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="14ff7-113">Restricted</span></span>|
|[<span data-ttu-id="14ff7-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-115">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="14ff7-116">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="14ff7-116">Members and methods</span></span>

| <span data-ttu-id="14ff7-117">Membre	</span><span class="sxs-lookup"><span data-stu-id="14ff7-117">Member</span></span> | <span data-ttu-id="14ff7-118">Type	</span><span class="sxs-lookup"><span data-stu-id="14ff7-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="14ff7-119">attachments</span><span class="sxs-lookup"><span data-stu-id="14ff7-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="14ff7-120">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-120">Member</span></span> |
| [<span data-ttu-id="14ff7-121">bcc</span><span class="sxs-lookup"><span data-stu-id="14ff7-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="14ff7-122">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-122">Member</span></span> |
| [<span data-ttu-id="14ff7-123">body</span><span class="sxs-lookup"><span data-stu-id="14ff7-123">body</span></span>](#body-body) | <span data-ttu-id="14ff7-124">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-124">Member</span></span> |
| [<span data-ttu-id="14ff7-125">cc</span><span class="sxs-lookup"><span data-stu-id="14ff7-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14ff7-126">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-126">Member</span></span> |
| [<span data-ttu-id="14ff7-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="14ff7-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="14ff7-128">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-128">Member</span></span> |
| [<span data-ttu-id="14ff7-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="14ff7-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="14ff7-130">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-130">Member</span></span> |
| [<span data-ttu-id="14ff7-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="14ff7-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="14ff7-132">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-132">Member</span></span> |
| [<span data-ttu-id="14ff7-133">end</span><span class="sxs-lookup"><span data-stu-id="14ff7-133">end</span></span>](#end-datetime) | <span data-ttu-id="14ff7-134">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-134">Member</span></span> |
| [<span data-ttu-id="14ff7-135">from</span><span class="sxs-lookup"><span data-stu-id="14ff7-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="14ff7-136">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-136">Member</span></span> |
| [<span data-ttu-id="14ff7-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="14ff7-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="14ff7-138">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-138">Member</span></span> |
| [<span data-ttu-id="14ff7-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="14ff7-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="14ff7-140">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-140">Member</span></span> |
| [<span data-ttu-id="14ff7-141">itemId</span><span class="sxs-lookup"><span data-stu-id="14ff7-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="14ff7-142">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-142">Member</span></span> |
| [<span data-ttu-id="14ff7-143">itemType</span><span class="sxs-lookup"><span data-stu-id="14ff7-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="14ff7-144">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-144">Member</span></span> |
| [<span data-ttu-id="14ff7-145">location</span><span class="sxs-lookup"><span data-stu-id="14ff7-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="14ff7-146">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-146">Member</span></span> |
| [<span data-ttu-id="14ff7-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="14ff7-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="14ff7-148">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-148">Member</span></span> |
| [<span data-ttu-id="14ff7-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="14ff7-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14ff7-150">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-150">Member</span></span> |
| [<span data-ttu-id="14ff7-151">organizer</span><span class="sxs-lookup"><span data-stu-id="14ff7-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="14ff7-152">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-152">Member</span></span> |
| [<span data-ttu-id="14ff7-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="14ff7-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14ff7-154">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-154">Member</span></span> |
| [<span data-ttu-id="14ff7-155">sender</span><span class="sxs-lookup"><span data-stu-id="14ff7-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="14ff7-156">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-156">Member</span></span> |
| [<span data-ttu-id="14ff7-157">start</span><span class="sxs-lookup"><span data-stu-id="14ff7-157">start</span></span>](#start-datetime) | <span data-ttu-id="14ff7-158">Member</span><span class="sxs-lookup"><span data-stu-id="14ff7-158">Member</span></span> |
| [<span data-ttu-id="14ff7-159">subject</span><span class="sxs-lookup"><span data-stu-id="14ff7-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="14ff7-160">Membre</span><span class="sxs-lookup"><span data-stu-id="14ff7-160">Member</span></span> |
| [<span data-ttu-id="14ff7-161">to</span><span class="sxs-lookup"><span data-stu-id="14ff7-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14ff7-162">Membre</span><span class="sxs-lookup"><span data-stu-id="14ff7-162">Member</span></span> |
| [<span data-ttu-id="14ff7-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="14ff7-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="14ff7-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-164">Method</span></span> |
| [<span data-ttu-id="14ff7-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="14ff7-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="14ff7-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-166">Method</span></span> |
| [<span data-ttu-id="14ff7-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="14ff7-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="14ff7-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-168">Method</span></span> |
| [<span data-ttu-id="14ff7-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="14ff7-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="14ff7-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-170">Method</span></span> |
| [<span data-ttu-id="14ff7-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="14ff7-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="14ff7-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-172">Method</span></span> |
| [<span data-ttu-id="14ff7-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="14ff7-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="14ff7-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-174">Method</span></span> |
| [<span data-ttu-id="14ff7-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="14ff7-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="14ff7-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-176">Method</span></span> |
| [<span data-ttu-id="14ff7-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="14ff7-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="14ff7-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-178">Method</span></span> |
| [<span data-ttu-id="14ff7-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="14ff7-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="14ff7-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-180">Method</span></span> |
| [<span data-ttu-id="14ff7-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="14ff7-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="14ff7-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-182">Method</span></span> |
| [<span data-ttu-id="14ff7-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="14ff7-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="14ff7-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="14ff7-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="14ff7-185">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-185">Example</span></span>

<span data-ttu-id="14ff7-186">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="14ff7-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="14ff7-187">Members</span><span class="sxs-lookup"><span data-stu-id="14ff7-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="14ff7-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="14ff7-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="14ff7-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-191">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="14ff7-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="14ff7-192">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="14ff7-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-193">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-193">Type</span></span>

*   <span data-ttu-id="14ff7-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="14ff7-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-195">Requirements</span></span>

|<span data-ttu-id="14ff7-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-196">Requirement</span></span>| <span data-ttu-id="14ff7-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-199">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-199">1.0</span></span>|
|[<span data-ttu-id="14ff7-200">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-201">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-202">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-203">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-204">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-204">Example</span></span>

<span data-ttu-id="14ff7-205">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14ff7-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-207">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="14ff7-208">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-208">Compose mode only.</span></span>

<span data-ttu-id="14ff7-209">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-209">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-210">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="14ff7-210">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14ff7-211">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-211">Get 500 members maximum.</span></span>
- <span data-ttu-id="14ff7-212">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="14ff7-212">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-213">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-213">Type</span></span>

*   [<span data-ttu-id="14ff7-214">Destinataires</span><span class="sxs-lookup"><span data-stu-id="14ff7-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14ff7-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-215">Requirements</span></span>

|<span data-ttu-id="14ff7-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-216">Requirement</span></span>| <span data-ttu-id="14ff7-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-219">1.1</span><span class="sxs-lookup"><span data-stu-id="14ff7-219">1.1</span></span>|
|[<span data-ttu-id="14ff7-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-221">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-223">Composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="14ff7-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-226">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="14ff7-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-227">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-227">Type</span></span>

*   [<span data-ttu-id="14ff7-228">Body</span><span class="sxs-lookup"><span data-stu-id="14ff7-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14ff7-229">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-229">Requirements</span></span>

|<span data-ttu-id="14ff7-230">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-230">Requirement</span></span>| <span data-ttu-id="14ff7-231">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-232">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-233">1.1</span><span class="sxs-lookup"><span data-stu-id="14ff7-233">1.1</span></span>|
|[<span data-ttu-id="14ff7-234">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-235">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-238">Example</span></span>

<span data-ttu-id="14ff7-239">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="14ff7-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="14ff7-240">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14ff7-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-242">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="14ff7-243">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-244">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-244">Read mode</span></span>

<span data-ttu-id="14ff7-245">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-245">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="14ff7-246">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-246">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-247">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-247">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-248">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-248">Compose mode</span></span>

<span data-ttu-id="14ff7-249">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="14ff7-250">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-251">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="14ff7-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14ff7-252">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="14ff7-253">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="14ff7-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14ff7-254">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-254">Type</span></span>

*   <span data-ttu-id="14ff7-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-256">Requirements</span></span>

|<span data-ttu-id="14ff7-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-257">Requirement</span></span>| <span data-ttu-id="14ff7-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-260">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-260">1.0</span></span>|
|[<span data-ttu-id="14ff7-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-262">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="14ff7-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="14ff7-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="14ff7-266">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="14ff7-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="14ff7-p110">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="14ff7-p111">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-271">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-271">Type</span></span>

*   <span data-ttu-id="14ff7-272">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-273">Requirements</span></span>

|<span data-ttu-id="14ff7-274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-274">Requirement</span></span>| <span data-ttu-id="14ff7-275">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-277">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-277">1.0</span></span>|
|[<span data-ttu-id="14ff7-278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-279">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-281">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-282">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="14ff7-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="14ff7-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="14ff7-p112">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-286">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-286">Type</span></span>

*   <span data-ttu-id="14ff7-287">Date</span><span class="sxs-lookup"><span data-stu-id="14ff7-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-288">Requirements</span></span>

|<span data-ttu-id="14ff7-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-289">Requirement</span></span>| <span data-ttu-id="14ff7-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-292">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-292">1.0</span></span>|
|[<span data-ttu-id="14ff7-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-294">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-296">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="14ff7-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="14ff7-298">dateTimeModified: Date</span></span>

<span data-ttu-id="14ff7-p113">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-301">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-302">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-302">Type</span></span>

*   <span data-ttu-id="14ff7-303">Date</span><span class="sxs-lookup"><span data-stu-id="14ff7-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-304">Requirements</span></span>

|<span data-ttu-id="14ff7-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-305">Requirement</span></span>| <span data-ttu-id="14ff7-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-308">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-308">1.0</span></span>|
|[<span data-ttu-id="14ff7-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-310">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="14ff7-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-315">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="14ff7-p114">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-318">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-318">Read mode</span></span>

<span data-ttu-id="14ff7-319">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-320">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-320">Compose mode</span></span>

<span data-ttu-id="14ff7-321">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="14ff7-322">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="14ff7-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="14ff7-323">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="14ff7-324">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-324">Type</span></span>

*   <span data-ttu-id="14ff7-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-326">Requirements</span></span>

|<span data-ttu-id="14ff7-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-327">Requirement</span></span>| <span data-ttu-id="14ff7-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-330">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-330">1.0</span></span>|
|[<span data-ttu-id="14ff7-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-332">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-334">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="14ff7-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-p115">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="14ff7-p116">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-340">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-341">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-341">Type</span></span>

*   [<span data-ttu-id="14ff7-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14ff7-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14ff7-343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-343">Requirements</span></span>

|<span data-ttu-id="14ff7-344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-344">Requirement</span></span>| <span data-ttu-id="14ff7-345">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-347">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-347">1.0</span></span>|
|[<span data-ttu-id="14ff7-348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-349">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-351">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-352">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="14ff7-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="14ff7-353">internetMessageId: String</span></span>

<span data-ttu-id="14ff7-p117">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-356">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-356">Type</span></span>

*   <span data-ttu-id="14ff7-357">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-358">Requirements</span></span>

|<span data-ttu-id="14ff7-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-359">Requirement</span></span>| <span data-ttu-id="14ff7-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-362">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-362">1.0</span></span>|
|[<span data-ttu-id="14ff7-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-364">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="14ff7-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="14ff7-368">itemClass: String</span></span>

<span data-ttu-id="14ff7-p118">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="14ff7-p119">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="14ff7-373">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-373">Type</span></span> | <span data-ttu-id="14ff7-374">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-374">Description</span></span> | <span data-ttu-id="14ff7-375">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="14ff7-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="14ff7-376">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="14ff7-376">Appointment items</span></span> | <span data-ttu-id="14ff7-377">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="14ff7-378">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="14ff7-378">Message items</span></span> | <span data-ttu-id="14ff7-379">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="14ff7-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="14ff7-380">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-381">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-381">Type</span></span>

*   <span data-ttu-id="14ff7-382">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-383">Requirements</span></span>

|<span data-ttu-id="14ff7-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-384">Requirement</span></span>| <span data-ttu-id="14ff7-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-387">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-387">1.0</span></span>|
|[<span data-ttu-id="14ff7-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-389">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="14ff7-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="14ff7-393">(nullable) itemId: String</span></span>

<span data-ttu-id="14ff7-394">Obtient l' [identificateur d’élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-394">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="14ff7-395">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-395">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-396">L’identificateur renvoyé par la `itemId` propriété est identique à l’identificateur d' [élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="14ff7-396">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="14ff7-397">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="14ff7-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="14ff7-398">Avant d’effectuer des appels d’API REST à l’aide de cette valeur `Office.context.mailbox.convertToRestId`, elle doit être convertie à l’aide de, qui est disponible à partir de l’ensemble de conditions requises 1,3.</span><span class="sxs-lookup"><span data-stu-id="14ff7-398">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="14ff7-399">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="14ff7-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-400">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-400">Type</span></span>

*   <span data-ttu-id="14ff7-401">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-402">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-402">Requirements</span></span>

|<span data-ttu-id="14ff7-403">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-403">Requirement</span></span>| <span data-ttu-id="14ff7-404">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-405">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-406">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-406">1.0</span></span>|
|[<span data-ttu-id="14ff7-407">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-407">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-408">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-409">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-409">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-410">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-411">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-411">Example</span></span>

<span data-ttu-id="14ff7-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="14ff7-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-415">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="14ff7-415">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="14ff7-416">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-416">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-417">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-417">Type</span></span>

*   [<span data-ttu-id="14ff7-418">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="14ff7-418">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14ff7-419">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-419">Requirements</span></span>

|<span data-ttu-id="14ff7-420">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-420">Requirement</span></span>| <span data-ttu-id="14ff7-421">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-421">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-422">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-423">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-423">1.0</span></span>|
|[<span data-ttu-id="14ff7-424">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-425">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-425">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-426">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-427">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-427">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-428">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-428">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="14ff7-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-430">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-430">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-431">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-431">Read mode</span></span>

<span data-ttu-id="14ff7-432">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-432">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-433">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-433">Compose mode</span></span>

<span data-ttu-id="14ff7-434">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-434">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14ff7-435">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-435">Type</span></span>

*   <span data-ttu-id="14ff7-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-437">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-437">Requirements</span></span>

|<span data-ttu-id="14ff7-438">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-438">Requirement</span></span>| <span data-ttu-id="14ff7-439">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-440">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-441">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-441">1.0</span></span>|
|[<span data-ttu-id="14ff7-442">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-443">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-444">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-445">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-445">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="14ff7-446">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="14ff7-446">normalizedSubject: String</span></span>

<span data-ttu-id="14ff7-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="14ff7-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="14ff7-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-451">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-451">Type</span></span>

*   <span data-ttu-id="14ff7-452">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-453">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-453">Requirements</span></span>

|<span data-ttu-id="14ff7-454">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-454">Requirement</span></span>| <span data-ttu-id="14ff7-455">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-456">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-457">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-457">1.0</span></span>|
|[<span data-ttu-id="14ff7-458">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-459">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-460">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-461">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-462">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14ff7-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-464">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-464">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="14ff7-465">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-465">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-466">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-466">Read mode</span></span>

<span data-ttu-id="14ff7-467">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="14ff7-467">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="14ff7-468">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-468">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-469">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-469">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-470">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-470">Compose mode</span></span>

<span data-ttu-id="14ff7-471">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="14ff7-471">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="14ff7-472">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-473">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="14ff7-473">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14ff7-474">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-474">Get 500 members maximum.</span></span>
- <span data-ttu-id="14ff7-475">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="14ff7-475">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14ff7-476">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-476">Type</span></span>

*   <span data-ttu-id="14ff7-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-478">Requirements</span></span>

|<span data-ttu-id="14ff7-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-479">Requirement</span></span>| <span data-ttu-id="14ff7-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-482">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-482">1.0</span></span>|
|[<span data-ttu-id="14ff7-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-484">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-486">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-486">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="14ff7-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-p128">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-490">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-490">Type</span></span>

*   [<span data-ttu-id="14ff7-491">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14ff7-491">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14ff7-492">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-492">Requirements</span></span>

|<span data-ttu-id="14ff7-493">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-493">Requirement</span></span>| <span data-ttu-id="14ff7-494">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-495">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-496">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-496">1.0</span></span>|
|[<span data-ttu-id="14ff7-497">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-498">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-499">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-500">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-500">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-501">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-501">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14ff7-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-503">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-503">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="14ff7-504">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-504">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-505">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-505">Read mode</span></span>

<span data-ttu-id="14ff7-506">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="14ff7-506">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="14ff7-507">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-507">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-508">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-508">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-509">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-509">Compose mode</span></span>

<span data-ttu-id="14ff7-510">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="14ff7-510">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="14ff7-511">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-512">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="14ff7-512">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14ff7-513">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-513">Get 500 members maximum.</span></span>
- <span data-ttu-id="14ff7-514">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="14ff7-514">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="14ff7-515">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-515">Type</span></span>

*   <span data-ttu-id="14ff7-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-517">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-517">Requirements</span></span>

|<span data-ttu-id="14ff7-518">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-518">Requirement</span></span>| <span data-ttu-id="14ff7-519">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-520">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-521">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-521">1.0</span></span>|
|[<span data-ttu-id="14ff7-522">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-523">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-524">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-525">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-525">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="14ff7-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-p132">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="14ff7-p133">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-531">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-531">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="14ff7-532">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-532">Type</span></span>

*   [<span data-ttu-id="14ff7-533">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14ff7-533">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14ff7-534">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-534">Requirements</span></span>

|<span data-ttu-id="14ff7-535">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-535">Requirement</span></span>| <span data-ttu-id="14ff7-536">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-537">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-538">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-538">1.0</span></span>|
|[<span data-ttu-id="14ff7-539">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-540">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-541">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-542">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-542">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-543">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-543">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="14ff7-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-545">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-545">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="14ff7-p134">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-548">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-548">Read mode</span></span>

<span data-ttu-id="14ff7-549">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-549">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-550">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-550">Compose mode</span></span>

<span data-ttu-id="14ff7-551">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-551">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="14ff7-552">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="14ff7-552">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="14ff7-553">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-553">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="14ff7-554">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-554">Type</span></span>

*   <span data-ttu-id="14ff7-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-556">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-556">Requirements</span></span>

|<span data-ttu-id="14ff7-557">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-557">Requirement</span></span>| <span data-ttu-id="14ff7-558">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-559">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-560">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-560">1.0</span></span>|
|[<span data-ttu-id="14ff7-561">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-562">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-563">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-564">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-564">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="14ff7-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-566">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="14ff7-566">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="14ff7-567">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="14ff7-567">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-568">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-568">Read mode</span></span>

<span data-ttu-id="14ff7-p135">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-571">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-571">Compose mode</span></span>

<span data-ttu-id="14ff7-572">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="14ff7-572">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="14ff7-573">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-573">Type</span></span>

*   <span data-ttu-id="14ff7-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-575">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-575">Requirements</span></span>

|<span data-ttu-id="14ff7-576">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-576">Requirement</span></span>| <span data-ttu-id="14ff7-577">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-577">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-578">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-578">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-579">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-579">1.0</span></span>|
|[<span data-ttu-id="14ff7-580">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-580">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-581">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-581">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-582">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-582">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-583">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-583">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14ff7-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14ff7-585">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-585">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="14ff7-586">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-586">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14ff7-587">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-587">Read mode</span></span>

<span data-ttu-id="14ff7-588">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-588">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="14ff7-589">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-589">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-590">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-590">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="14ff7-591">Mode composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-591">Compose mode</span></span>

<span data-ttu-id="14ff7-592">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="14ff7-592">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="14ff7-593">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="14ff7-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14ff7-594">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="14ff7-594">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14ff7-595">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="14ff7-595">Get 500 members maximum.</span></span>
- <span data-ttu-id="14ff7-596">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="14ff7-596">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14ff7-597">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-597">Type</span></span>

*   <span data-ttu-id="14ff7-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-599">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-599">Requirements</span></span>

|<span data-ttu-id="14ff7-600">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-600">Requirement</span></span>| <span data-ttu-id="14ff7-601">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-602">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-603">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-603">1.0</span></span>|
|[<span data-ttu-id="14ff7-604">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-605">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-606">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-607">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-607">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="14ff7-608">Méthodes</span><span class="sxs-lookup"><span data-stu-id="14ff7-608">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="14ff7-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14ff7-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="14ff7-610">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="14ff7-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="14ff7-611">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="14ff7-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="14ff7-612">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="14ff7-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-613">Paramètres</span><span class="sxs-lookup"><span data-stu-id="14ff7-613">Parameters</span></span>

|<span data-ttu-id="14ff7-614">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-614">Name</span></span>| <span data-ttu-id="14ff7-615">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-615">Type</span></span>| <span data-ttu-id="14ff7-616">Attributs</span><span class="sxs-lookup"><span data-stu-id="14ff7-616">Attributes</span></span>| <span data-ttu-id="14ff7-617">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="14ff7-618">Chaîne</span><span class="sxs-lookup"><span data-stu-id="14ff7-618">String</span></span>||<span data-ttu-id="14ff7-p139">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="14ff7-621">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-621">String</span></span>||<span data-ttu-id="14ff7-p140">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="14ff7-624">Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-624">Object</span></span>| <span data-ttu-id="14ff7-625">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-625">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-626">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-626">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14ff7-627">Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-627">Object</span></span>| <span data-ttu-id="14ff7-628">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-628">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-629">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-629">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14ff7-630">fonction</span><span class="sxs-lookup"><span data-stu-id="14ff7-630">function</span></span>| <span data-ttu-id="14ff7-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-631">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-632">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14ff7-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14ff7-633">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-633">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="14ff7-634">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="14ff7-634">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14ff7-635">Erreurs</span><span class="sxs-lookup"><span data-stu-id="14ff7-635">Errors</span></span>

| <span data-ttu-id="14ff7-636">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="14ff7-636">Error code</span></span> | <span data-ttu-id="14ff7-637">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-637">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="14ff7-638">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="14ff7-638">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="14ff7-639">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="14ff7-639">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="14ff7-640">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-640">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14ff7-641">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-641">Requirements</span></span>

|<span data-ttu-id="14ff7-642">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-642">Requirement</span></span>| <span data-ttu-id="14ff7-643">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-644">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-645">1.1</span><span class="sxs-lookup"><span data-stu-id="14ff7-645">1.1</span></span>|
|[<span data-ttu-id="14ff7-646">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-647">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-647">ReadWriteItem</span></span>|
|[<span data-ttu-id="14ff7-648">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-649">Composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-649">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-650">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-650">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="14ff7-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14ff7-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="14ff7-652">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-652">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="14ff7-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="14ff7-656">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="14ff7-656">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="14ff7-657">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="14ff7-657">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-658">Paramètres</span><span class="sxs-lookup"><span data-stu-id="14ff7-658">Parameters</span></span>

|<span data-ttu-id="14ff7-659">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-659">Name</span></span>| <span data-ttu-id="14ff7-660">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-660">Type</span></span>| <span data-ttu-id="14ff7-661">Attributs</span><span class="sxs-lookup"><span data-stu-id="14ff7-661">Attributes</span></span>| <span data-ttu-id="14ff7-662">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-662">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="14ff7-663">Chaîne</span><span class="sxs-lookup"><span data-stu-id="14ff7-663">String</span></span>||<span data-ttu-id="14ff7-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="14ff7-666">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-666">String</span></span>||<span data-ttu-id="14ff7-667">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="14ff7-667">The subject of the item to be attached.</span></span> <span data-ttu-id="14ff7-668">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="14ff7-668">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="14ff7-669">Object</span><span class="sxs-lookup"><span data-stu-id="14ff7-669">Object</span></span>| <span data-ttu-id="14ff7-670">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-670">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-671">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-671">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14ff7-672">Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-672">Object</span></span>| <span data-ttu-id="14ff7-673">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-673">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-674">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-674">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14ff7-675">fonction</span><span class="sxs-lookup"><span data-stu-id="14ff7-675">function</span></span>| <span data-ttu-id="14ff7-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-676">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-677">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14ff7-677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14ff7-678">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-678">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="14ff7-679">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="14ff7-679">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14ff7-680">Erreurs</span><span class="sxs-lookup"><span data-stu-id="14ff7-680">Errors</span></span>

| <span data-ttu-id="14ff7-681">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="14ff7-681">Error code</span></span> | <span data-ttu-id="14ff7-682">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-682">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="14ff7-683">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-683">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14ff7-684">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-684">Requirements</span></span>

|<span data-ttu-id="14ff7-685">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-685">Requirement</span></span>| <span data-ttu-id="14ff7-686">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-686">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-687">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-687">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-688">1.1</span><span class="sxs-lookup"><span data-stu-id="14ff7-688">1.1</span></span>|
|[<span data-ttu-id="14ff7-689">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-689">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-690">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-690">ReadWriteItem</span></span>|
|[<span data-ttu-id="14ff7-691">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-691">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-692">Composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-692">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-693">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-693">Example</span></span>

<span data-ttu-id="14ff7-694">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-694">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="14ff7-695">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="14ff7-695">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="14ff7-696">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="14ff7-696">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-697">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-697">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14ff7-698">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-698">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="14ff7-699">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="14ff7-699">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-700">La possibilité d’inclure des pièces jointes dans `displayReplyAllForm` l’appel à n’est pas prise en charge dans l’ensemble de conditions requises 1,1.</span><span class="sxs-lookup"><span data-stu-id="14ff7-700">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="14ff7-701">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="14ff7-701">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-702">Parameters</span><span class="sxs-lookup"><span data-stu-id="14ff7-702">Parameters</span></span>

|<span data-ttu-id="14ff7-703">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-703">Name</span></span>| <span data-ttu-id="14ff7-704">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-704">Type</span></span>| <span data-ttu-id="14ff7-705">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-705">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="14ff7-706">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="14ff7-706">String &#124; Object</span></span>| |<span data-ttu-id="14ff7-p145">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="14ff7-709">**OU**</span><span class="sxs-lookup"><span data-stu-id="14ff7-709">**OR**</span></span><br/><span data-ttu-id="14ff7-p146">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="14ff7-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="14ff7-712">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-712">String</span></span> | <span data-ttu-id="14ff7-713">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-713">&lt;optional&gt;</span></span> | <span data-ttu-id="14ff7-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="14ff7-716">fonction</span><span class="sxs-lookup"><span data-stu-id="14ff7-716">function</span></span> | <span data-ttu-id="14ff7-717">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-717">&lt;optional&gt;</span></span> | <span data-ttu-id="14ff7-718">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14ff7-718">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14ff7-719">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-719">Requirements</span></span>

|<span data-ttu-id="14ff7-720">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-720">Requirement</span></span>| <span data-ttu-id="14ff7-721">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-721">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-722">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-722">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-723">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-723">1.0</span></span>|
|[<span data-ttu-id="14ff7-724">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-724">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-725">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-725">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-726">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-726">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-727">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-727">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="14ff7-728">Exemples</span><span class="sxs-lookup"><span data-stu-id="14ff7-728">Examples</span></span>

<span data-ttu-id="14ff7-729">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-729">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="14ff7-730">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="14ff7-730">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="14ff7-731">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="14ff7-731">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="14ff7-732">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-732">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="14ff7-733">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="14ff7-733">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="14ff7-734">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="14ff7-734">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-735">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-735">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14ff7-736">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-736">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="14ff7-737">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="14ff7-737">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-738">La possibilité d’inclure des pièces jointes dans `displayReplyForm` l’appel à n’est pas prise en charge dans l’ensemble de conditions requises 1,1.</span><span class="sxs-lookup"><span data-stu-id="14ff7-738">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="14ff7-739">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="14ff7-739">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-740">Parameters</span><span class="sxs-lookup"><span data-stu-id="14ff7-740">Parameters</span></span>

|<span data-ttu-id="14ff7-741">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-741">Name</span></span>| <span data-ttu-id="14ff7-742">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-742">Type</span></span>| <span data-ttu-id="14ff7-743">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-743">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="14ff7-744">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="14ff7-744">String &#124; Object</span></span>| | <span data-ttu-id="14ff7-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="14ff7-747">**OU**</span><span class="sxs-lookup"><span data-stu-id="14ff7-747">**OR**</span></span><br/><span data-ttu-id="14ff7-p150">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="14ff7-p150">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="14ff7-750">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-750">String</span></span> | <span data-ttu-id="14ff7-751">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-751">&lt;optional&gt;</span></span> | <span data-ttu-id="14ff7-p151">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="14ff7-754">fonction</span><span class="sxs-lookup"><span data-stu-id="14ff7-754">function</span></span> | <span data-ttu-id="14ff7-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-755">&lt;optional&gt;</span></span> | <span data-ttu-id="14ff7-756">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14ff7-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14ff7-757">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-757">Requirements</span></span>

|<span data-ttu-id="14ff7-758">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-758">Requirement</span></span>| <span data-ttu-id="14ff7-759">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-760">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-761">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-761">1.0</span></span>|
|[<span data-ttu-id="14ff7-762">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-763">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-764">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-765">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="14ff7-766">Exemples</span><span class="sxs-lookup"><span data-stu-id="14ff7-766">Examples</span></span>

<span data-ttu-id="14ff7-767">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-767">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="14ff7-768">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="14ff7-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="14ff7-769">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="14ff7-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="14ff7-770">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-770">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="14ff7-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="14ff7-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="14ff7-772">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="14ff7-772">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-773">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-773">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-774">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-774">Requirements</span></span>

|<span data-ttu-id="14ff7-775">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-775">Requirement</span></span>| <span data-ttu-id="14ff7-776">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-776">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-777">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-777">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-778">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-778">1.0</span></span>|
|[<span data-ttu-id="14ff7-779">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-779">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-780">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-780">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-781">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-781">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-782">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-782">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14ff7-783">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="14ff7-783">Returns:</span></span>

<span data-ttu-id="14ff7-784">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14ff7-784">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="14ff7-785">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-785">Example</span></span>

<span data-ttu-id="14ff7-786">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-786">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="14ff7-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="14ff7-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="14ff7-788">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="14ff7-788">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-789">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-789">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-790">Paramètres</span><span class="sxs-lookup"><span data-stu-id="14ff7-790">Parameters</span></span>

|<span data-ttu-id="14ff7-791">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-791">Name</span></span>| <span data-ttu-id="14ff7-792">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-792">Type</span></span>| <span data-ttu-id="14ff7-793">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-793">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="14ff7-794">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="14ff7-794">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="14ff7-795">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="14ff7-795">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14ff7-796">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-796">Requirements</span></span>

|<span data-ttu-id="14ff7-797">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-797">Requirement</span></span>| <span data-ttu-id="14ff7-798">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-799">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-800">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-800">1.0</span></span>|
|[<span data-ttu-id="14ff7-801">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-802">Restreinte</span><span class="sxs-lookup"><span data-stu-id="14ff7-802">Restricted</span></span>|
|[<span data-ttu-id="14ff7-803">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-804">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14ff7-805">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="14ff7-805">Returns:</span></span>

<span data-ttu-id="14ff7-806">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="14ff7-806">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="14ff7-807">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="14ff7-807">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="14ff7-808">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-808">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="14ff7-809">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="14ff7-809">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="14ff7-810">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="14ff7-810">Value of `entityType`</span></span> | <span data-ttu-id="14ff7-811">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="14ff7-811">Type of objects in returned array</span></span> | <span data-ttu-id="14ff7-812">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="14ff7-812">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="14ff7-813">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-813">String</span></span> | <span data-ttu-id="14ff7-814">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="14ff7-814">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="14ff7-815">Contact</span><span class="sxs-lookup"><span data-stu-id="14ff7-815">Contact</span></span> | <span data-ttu-id="14ff7-816">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14ff7-816">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="14ff7-817">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-817">String</span></span> | <span data-ttu-id="14ff7-818">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14ff7-818">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="14ff7-819">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="14ff7-819">MeetingSuggestion</span></span> | <span data-ttu-id="14ff7-820">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14ff7-820">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="14ff7-821">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="14ff7-821">PhoneNumber</span></span> | <span data-ttu-id="14ff7-822">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="14ff7-822">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="14ff7-823">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="14ff7-823">TaskSuggestion</span></span> | <span data-ttu-id="14ff7-824">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14ff7-824">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="14ff7-825">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-825">String</span></span> | <span data-ttu-id="14ff7-826">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="14ff7-826">**Restricted**</span></span> |

<span data-ttu-id="14ff7-827">Type :  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="14ff7-827">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="14ff7-828">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-828">Example</span></span>

<span data-ttu-id="14ff7-829">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="14ff7-829">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="14ff7-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="14ff7-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="14ff7-831">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="14ff7-831">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-832">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-832">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14ff7-833">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="14ff7-833">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-834">Parameters</span><span class="sxs-lookup"><span data-stu-id="14ff7-834">Parameters</span></span>

|<span data-ttu-id="14ff7-835">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-835">Name</span></span>| <span data-ttu-id="14ff7-836">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-836">Type</span></span>| <span data-ttu-id="14ff7-837">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-837">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="14ff7-838">Chaîne</span><span class="sxs-lookup"><span data-stu-id="14ff7-838">String</span></span>|<span data-ttu-id="14ff7-839">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="14ff7-839">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14ff7-840">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-840">Requirements</span></span>

|<span data-ttu-id="14ff7-841">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-841">Requirement</span></span>| <span data-ttu-id="14ff7-842">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-843">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-844">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-844">1.0</span></span>|
|[<span data-ttu-id="14ff7-845">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-846">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-847">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-848">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14ff7-849">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="14ff7-849">Returns:</span></span>

<span data-ttu-id="14ff7-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="14ff7-852">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="14ff7-852">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="14ff7-853">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="14ff7-853">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="14ff7-854">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="14ff7-854">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-855">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14ff7-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="14ff7-859">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="14ff7-859">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="14ff7-860">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-860">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="14ff7-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14ff7-863">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-863">Requirements</span></span>

|<span data-ttu-id="14ff7-864">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-864">Requirement</span></span>| <span data-ttu-id="14ff7-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-866">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-867">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-867">1.0</span></span>|
|[<span data-ttu-id="14ff7-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-869">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14ff7-872">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="14ff7-872">Returns:</span></span>

<span data-ttu-id="14ff7-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="14ff7-875">Type : Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-875">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="14ff7-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-876">Example</span></span>

<span data-ttu-id="14ff7-877">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="14ff7-877">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="14ff7-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="14ff7-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="14ff7-879">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="14ff7-879">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14ff7-880">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="14ff7-880">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14ff7-881">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="14ff7-881">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="14ff7-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-884">Parameters</span><span class="sxs-lookup"><span data-stu-id="14ff7-884">Parameters</span></span>

|<span data-ttu-id="14ff7-885">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-885">Name</span></span>| <span data-ttu-id="14ff7-886">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-886">Type</span></span>| <span data-ttu-id="14ff7-887">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="14ff7-888">Chaîne</span><span class="sxs-lookup"><span data-stu-id="14ff7-888">String</span></span>|<span data-ttu-id="14ff7-889">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="14ff7-889">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14ff7-890">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-890">Requirements</span></span>

|<span data-ttu-id="14ff7-891">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-891">Requirement</span></span>| <span data-ttu-id="14ff7-892">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-893">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-894">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-894">1.0</span></span>|
|[<span data-ttu-id="14ff7-895">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-896">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-897">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-898">Lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14ff7-899">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="14ff7-899">Returns:</span></span>

<span data-ttu-id="14ff7-900">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="14ff7-900">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="14ff7-901">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="14ff7-901">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="14ff7-902">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-902">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="14ff7-903">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="14ff7-903">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="14ff7-904">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="14ff7-904">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="14ff7-p158">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p158">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-908">Parameters</span><span class="sxs-lookup"><span data-stu-id="14ff7-908">Parameters</span></span>

|<span data-ttu-id="14ff7-909">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-909">Name</span></span>| <span data-ttu-id="14ff7-910">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-910">Type</span></span>| <span data-ttu-id="14ff7-911">Attributs</span><span class="sxs-lookup"><span data-stu-id="14ff7-911">Attributes</span></span>| <span data-ttu-id="14ff7-912">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-912">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="14ff7-913">function</span><span class="sxs-lookup"><span data-stu-id="14ff7-913">function</span></span>||<span data-ttu-id="14ff7-914">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14ff7-914">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14ff7-915">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14ff7-915">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="14ff7-916">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="14ff7-916">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="14ff7-917">Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-917">Object</span></span>| <span data-ttu-id="14ff7-918">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-918">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-919">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-919">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="14ff7-920">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-920">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14ff7-921">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-921">Requirements</span></span>

|<span data-ttu-id="14ff7-922">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-922">Requirement</span></span>| <span data-ttu-id="14ff7-923">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-923">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-924">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-924">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-925">1.0</span><span class="sxs-lookup"><span data-stu-id="14ff7-925">1.0</span></span>|
|[<span data-ttu-id="14ff7-926">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-926">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-927">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-927">ReadItem</span></span>|
|[<span data-ttu-id="14ff7-928">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-928">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-929">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14ff7-929">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-930">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-930">Example</span></span>

<span data-ttu-id="14ff7-p161">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="14ff7-p161">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="14ff7-934">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14ff7-934">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="14ff7-935">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="14ff7-935">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="14ff7-936">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="14ff7-936">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="14ff7-937">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="14ff7-937">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="14ff7-938">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="14ff7-938">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="14ff7-939">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="14ff7-939">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14ff7-940">Paramètres</span><span class="sxs-lookup"><span data-stu-id="14ff7-940">Parameters</span></span>

|<span data-ttu-id="14ff7-941">Nom</span><span class="sxs-lookup"><span data-stu-id="14ff7-941">Name</span></span>| <span data-ttu-id="14ff7-942">Type</span><span class="sxs-lookup"><span data-stu-id="14ff7-942">Type</span></span>| <span data-ttu-id="14ff7-943">Attributs</span><span class="sxs-lookup"><span data-stu-id="14ff7-943">Attributes</span></span>| <span data-ttu-id="14ff7-944">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-944">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="14ff7-945">String</span><span class="sxs-lookup"><span data-stu-id="14ff7-945">String</span></span>||<span data-ttu-id="14ff7-946">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="14ff7-946">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="14ff7-947">Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-947">Object</span></span>| <span data-ttu-id="14ff7-948">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-948">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-949">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="14ff7-949">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14ff7-950">Objet</span><span class="sxs-lookup"><span data-stu-id="14ff7-950">Object</span></span>| <span data-ttu-id="14ff7-951">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-951">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-952">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="14ff7-952">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14ff7-953">fonction</span><span class="sxs-lookup"><span data-stu-id="14ff7-953">function</span></span>| <span data-ttu-id="14ff7-954">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="14ff7-954">&lt;optional&gt;</span></span>|<span data-ttu-id="14ff7-955">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14ff7-955">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14ff7-956">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="14ff7-956">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14ff7-957">Erreurs</span><span class="sxs-lookup"><span data-stu-id="14ff7-957">Errors</span></span>

| <span data-ttu-id="14ff7-958">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="14ff7-958">Error code</span></span> | <span data-ttu-id="14ff7-959">Description</span><span class="sxs-lookup"><span data-stu-id="14ff7-959">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="14ff7-960">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="14ff7-960">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14ff7-961">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14ff7-961">Requirements</span></span>

|<span data-ttu-id="14ff7-962">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14ff7-962">Requirement</span></span>| <span data-ttu-id="14ff7-963">Valeur</span><span class="sxs-lookup"><span data-stu-id="14ff7-963">Value</span></span>|
|---|---|
|[<span data-ttu-id="14ff7-964">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14ff7-964">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14ff7-965">1.1</span><span class="sxs-lookup"><span data-stu-id="14ff7-965">1.1</span></span>|
|[<span data-ttu-id="14ff7-966">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14ff7-966">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14ff7-967">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14ff7-967">ReadWriteItem</span></span>|
|[<span data-ttu-id="14ff7-968">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14ff7-968">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14ff7-969">Composition</span><span class="sxs-lookup"><span data-stu-id="14ff7-969">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14ff7-970">Exemple</span><span class="sxs-lookup"><span data-stu-id="14ff7-970">Example</span></span>

<span data-ttu-id="14ff7-971">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="14ff7-971">The following code removes an attachment with an identifier of '0'.</span></span>

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
